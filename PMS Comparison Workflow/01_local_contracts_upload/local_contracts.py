"""
Local folder analysis + upload to Supabase Storage and public.contract_files.

Expected layout (fixed):
  LOCAL_CONTRACTS_ROOT/
    <NC_xxxx_CUSTOMER>/
      02. CONTRACT/
        *.pdf | *.doc | *.docx   # direct files only — subfolders under CONTRACT are ignored

Upload picks **one file per customer folder**:
  - Prefer a single PDF; if several PDFs, prefer names with Complete / Docusign / signed / firmado;
    deprioritize copy / duplicate / do not edit / borrador.
  - If no PDF, use one Word file (same scoring).
  - First UPLOAD_BATCH_SIZE uploads go to BATCH1_PREFIX (default "To Fill 1"), the rest to BATCH2_PREFIX ("To Fill 2").

Commands:
  py local_contracts.py analyze
  py local_contracts.py upload
  py local_contracts.py upload --only-missing   # skip paths already in contract_files (DB)
  py local_contracts.py upload --no-batch-split  # use SUPABASE_STORAGE_PREFIX for every file (legacy)
  py local_contracts.py upload --start 0 --limit 600   # batch 1 only → To Fill 1
  py local_contracts.py upload --start 600             # batch 2 only → To Fill 2 (after batch 1 done)
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import random
import re
import sys
import time
import unicodedata
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

# Override existing env vars so .env wins (fixes stale wrong SUPABASE_URL in Windows).
load_dotenv(override=True)

from supabase import Client, ClientOptions, create_client

ROOT = Path(__file__).resolve().parent
REPORT_PATH = ROOT / "analysis_report.json"
UPLOAD_CHECKPOINT = ROOT / "upload_checkpoint.json"
# Exact folder name under each customer (case-insensitive).
CONTRACT_SUBDIR_NAME = "02. CONTRACT"
ALLOWED_EXT = frozenset({".pdf", ".doc", ".docx"})
MIME_FALLBACK = {
    ".pdf": "application/pdf",
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

# Prefer “final” contract PDFs when multiple exist.
_PDF_POSITIVE = (
    "complete",
    "completed",
    "docusign",
    "signed",
    "firmado",
    "contrato firmado",
)
_PDF_NEGATIVE = (
    "copy of",
    "copy ",
    "duplicate",
    "do not edit",
    "borrador",
)

# Spanish/Latin PDFs often use typographic quotes and dashes; normalize before NFKD.
_STORAGE_PUNCT = str.maketrans(
    {
        "\u2018": "'",  # ‘
        "\u2019": "'",  # ’
        "\u201b": "'",  # ‛
        "\u201c": '"',  # “
        "\u201d": '"',  # ”
        "\u00ab": '"',  # «
        "\u00bb": '"',  # »
        "\u2013": "-",  # en dash
        "\u2014": "-",  # em dash
        "\u00a0": " ",  # nbsp
    }
)


def safe_segment(name: str) -> str:
    """
    Storage object keys must be ASCII-safe (Supabase returns InvalidKey otherwise).

    Spanish (and other Latin) names in folders/files are *folded* to ASCII here:
    García→Garcia, CONTRATACIÓN→CONTRATACION, Muñoz→Munoz (ñ decomposes to n).

    Original spelling is unchanged on disk and in Postgres (local_relative_path,
    file_name on contract_files); only the Storage path uses this slug.
    """
    s = name.strip()
    if not s:
        return "unnamed"
    s = s.translate(_STORAGE_PUNCT)
    # Unicode NFKD splits letters from accents; strip combining marks (café → cafe).
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = "".join(ch if ord(ch) < 128 else "_" for ch in s)
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = s.replace("..", "_")
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "unnamed"


def find_contract_dir(parent: Path) -> Path | None:
    """Direct child only: <customer>/02. CONTRACT"""
    if not parent.is_dir():
        return None
    want = CONTRACT_SUBDIR_NAME.casefold()
    try:
        for child in parent.iterdir():
            if child.is_dir() and child.name.strip().casefold() == want:
                return child
    except OSError:
        return None
    return None


def _contract_file_score(path: Path) -> tuple:
    """Higher tuple = better when using max()."""
    n = path.name.casefold()
    pos = any(k in n for k in _PDF_POSITIVE)
    neg = any(k in n for k in _PDF_NEGATIVE)
    try:
        sz = path.stat().st_size
    except OSError:
        sz = 0
    return (pos, not neg, sz)


def pick_one_of(paths: list[Path]) -> Path:
    if len(paths) == 1:
        return paths[0]
    return max(paths, key=_contract_file_score)


def pick_one_contract_file_direct_only(cdir: Path) -> Path | None:
    """
    Only files directly inside 02. CONTRACT/ (no subfolders).
    Prefer one PDF; if none, one .doc/.docx. If several, use scoring.
    """
    pdfs: list[Path] = []
    words: list[Path] = []
    try:
        for f in cdir.iterdir():
            if not f.is_file():
                continue
            ext = f.suffix.lower()
            if ext == ".pdf":
                pdfs.append(f)
            elif ext in (".doc", ".docx"):
                words.append(f)
    except OSError:
        return None
    if pdfs:
        return pick_one_of(pdfs)
    if words:
        return pick_one_of(words)
    return None


def analyze(root: Path) -> dict[str, Any]:
    if not root.is_dir():
        raise FileNotFoundError(f"LOCAL_CONTRACTS_ROOT is not a directory: {root}")

    entries: list[dict[str, Any]] = []
    missing_contract: list[str] = []
    contract_folder_but_no_files: list[str] = []

    for child in sorted(root.iterdir(), key=lambda p: p.name.lower()):
        if not child.is_dir():
            continue
        name = child.name
        cdir = find_contract_dir(child)
        if not cdir:
            missing_contract.append(name)
            continue

        f = pick_one_contract_file_direct_only(cdir)
        if not f:
            contract_folder_but_no_files.append(name)
            continue
        ext = f.suffix.lower()
        try:
            st = f.stat()
            size = st.st_size
            mtime = st.st_mtime
        except OSError:
            contract_folder_but_no_files.append(name)
            continue
        rel = f.relative_to(root)
        mime = MIME_FALLBACK.get(ext) or (
            mimetypes.guess_type(f.name)[0] or "application/octet-stream"
        )
        entries.append(
            {
                "parent_folder": name,
                "file_name": f.name,
                "file_ext": ext,
                "size_bytes": size,
                "mtime_unix": mtime,
                "local_relative_path": str(rel).replace("\\", "/"),
                "mime_type": mime,
            }
        )

    top_dirs = sum(1 for p in root.iterdir() if p.is_dir())
    return {
        "local_root": str(root.resolve()),
        "total_top_level_dirs": top_dirs,
        "folders_with_contract_subfolder": top_dirs - len(missing_contract),
        "folders_missing_contract_subfolder": missing_contract,
        "folders_contract_empty_or_wrong_ext": contract_folder_but_no_files,
        "parent_folders_with_uploadable_files": len({e["parent_folder"] for e in entries}),
        "contract_files_count": len(entries),
        "files": entries,
    }


def load_upload_done() -> set[str]:
    if not UPLOAD_CHECKPOINT.is_file():
        return set()
    try:
        data = json.loads(UPLOAD_CHECKPOINT.read_text(encoding="utf-8"))
        return set(data.get("done", []))
    except (json.JSONDecodeError, OSError):
        return set()


def save_upload_done(done: set[str]) -> None:
    UPLOAD_CHECKPOINT.write_text(
        json.dumps({"done": sorted(done)}, indent=2),
        encoding="utf-8",
    )


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _is_transient_upload_error(exc: BaseException) -> bool:
    """Network blips, timeouts, and common rate-limit / overload responses."""
    if isinstance(exc, (TimeoutError, ConnectionError, BrokenPipeError)):
        return True
    # Windows: e.g. WinError 10054, errno 11001 DNS
    if isinstance(exc, OSError) and getattr(exc, "errno", None) in (
        11001,
        10054,
        10060,
        11002,
    ):
        return True
    msg = str(exc).lower()
    needles = (
        "429",
        "503",
        "502",
        "504",
        "timeout",
        "timed out",
        "temporarily unavailable",
        "rate limit",
        "too many requests",
        "connection reset",
        "connection aborted",
        "read operation timed out",
        "remote end closed",
    )
    if any(n in msg for n in needles):
        return True
    try:
        import httpx

        if isinstance(
            exc,
            (
                httpx.TimeoutException,
                httpx.ConnectError,
                httpx.ReadError,
                httpx.RemoteProtocolError,
                httpx.WriteError,
            ),
        ):
            return True
    except ImportError:
        pass
    return False


def _norm_rel_path(p: str) -> str:
    return str(p).replace("\\", "/")


def fetch_existing_contract_paths(client: Client) -> set[str]:
    """
    All local_relative_path / file_name values already recorded in contract_files.
    Paginated (PostgREST default page size is 1000).
    """
    out: set[str] = set()
    page = 1000
    offset = 0
    while True:
        resp = (
            client.table("contract_files")
            .select("local_relative_path,file_name")
            .range(offset, offset + page - 1)
            .execute()
        )
        rows = resp.data or []
        if not rows:
            break
        for row in rows:
            p = row.get("local_relative_path") or row.get("file_name")
            if p:
                out.add(_norm_rel_path(p))
        if len(rows) < page:
            break
        offset += page
    return out


def get_supabase() -> tuple[Client, str]:
    url = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    bucket = os.environ.get("SUPABASE_BUCKET", "contracts").strip()
    if not url or not key:
        raise SystemExit(
            "Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY in .env (service role JWT)."
        )
    # Default in supabase-py is storage_client_timeout=20s — too low for large PDFs.
    storage_timeout = _env_int("SUPABASE_STORAGE_TIMEOUT_SECONDS", 300)
    postgrest_timeout = _env_int("SUPABASE_POSTGREST_TIMEOUT_SECONDS", 120)
    opts = ClientOptions(
        storage_client_timeout=storage_timeout,
        postgrest_client_timeout=postgrest_timeout,
    )
    return create_client(url, key, options=opts), bucket


def cmd_analyze(args: argparse.Namespace) -> int:
    root = Path(args.root or os.environ.get("LOCAL_CONTRACTS_ROOT", "contracts_data"))
    root = root.expanduser().resolve()
    report = analyze(root)
    REPORT_PATH.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(f"Root: {report['local_root']}")
    print(f"Top-level folders: {report['total_top_level_dirs']}")
    print(f"With subfolder \"{CONTRACT_SUBDIR_NAME}\": {report['folders_with_contract_subfolder']}")
    print(f"Missing \"{CONTRACT_SUBDIR_NAME}\": {len(report['folders_missing_contract_subfolder'])}")
    print(
        f"Parent folders with a chosen contract file (direct in CONTRACT only, one per folder): {report['parent_folders_with_uploadable_files']}"
    )
    print(f"Chosen files to upload (max one per customer): {report['contract_files_count']}")
    print(f"Report written to {REPORT_PATH}")
    miss = report["folders_missing_contract_subfolder"]
    if miss and len(miss) <= 30:
        print(f"Folders without \"{CONTRACT_SUBDIR_NAME}\":", ", ".join(miss))
    elif miss:
        print(f"First 20 folders without \"{CONTRACT_SUBDIR_NAME}\": {', '.join(miss[:20])} ...")
    return 0


def _upload_batch_split_enabled(args: argparse.Namespace) -> bool:
    if getattr(args, "no_batch_split", False):
        return False
    return os.environ.get("USE_UPLOAD_BATCH_SPLIT", "true").lower() in (
        "1",
        "true",
        "yes",
    )


def _upload_prefix_for_index(idx: int, args: argparse.Namespace) -> str:
    """Storage folder under bucket: batch split, or legacy SUPABASE_STORAGE_PREFIX."""
    if _upload_batch_split_enabled(args):
        batch = (
            args.batch_size
            if getattr(args, "batch_size", None) is not None
            else _env_int("UPLOAD_BATCH_SIZE", 600)
        )
        raw1 = getattr(args, "batch1_prefix", None) or os.environ.get(
            "BATCH1_PREFIX", "To Fill 1"
        )
        raw2 = getattr(args, "batch2_prefix", None) or os.environ.get(
            "BATCH2_PREFIX", "To Fill 2"
        )
        p = raw1.strip().strip("/") if idx < batch else raw2.strip().strip("/")
        return p
    return os.environ.get("SUPABASE_STORAGE_PREFIX", "").strip().strip("/")


def cmd_upload(args: argparse.Namespace) -> int:
    root = Path(args.root or os.environ.get("LOCAL_CONTRACTS_ROOT", "contracts_data"))
    root = root.expanduser().resolve()
    use_cp = os.environ.get("USE_UPLOAD_CHECKPOINT", "true").lower() in (
        "1",
        "true",
        "yes",
    )
    done = load_upload_done() if use_cp else set()

    report = analyze(root)
    if not report["files"]:
        print("No files to upload. Run analyze and check folder layout.", file=sys.stderr)
        return 1

    files_sorted = sorted(report["files"], key=lambda x: x["parent_folder"].lower())

    start = max(0, int(getattr(args, "start", 0) or 0))
    lim = getattr(args, "limit", None)
    if lim is not None:
        lim = int(lim)
        if lim < 0:
            lim = 0
    if start >= len(files_sorted):
        print(
            f"Nothing to upload: --start {start} is past end ({len(files_sorted)} file(s) total).",
            file=sys.stderr,
        )
        return 1
    files_window = files_sorted[start:]
    if lim is not None:
        files_window = files_window[:lim]

    client, bucket = get_supabase()

    only_missing = getattr(args, "only_missing", False)
    existing_db: set[str] = set()
    if only_missing:
        existing_db = fetch_existing_contract_paths(client)
        print(f"DB already has {len(existing_db)} contract file path(s); will skip those.")

    if _upload_batch_split_enabled(args):
        bs = (
            args.batch_size
            if getattr(args, "batch_size", None) is not None
            else _env_int("UPLOAD_BATCH_SIZE", 600)
        )
        p1 = (getattr(args, "batch1_prefix", None) or os.environ.get("BATCH1_PREFIX", "To Fill 1")).strip()
        p2 = (getattr(args, "batch2_prefix", None) or os.environ.get("BATCH2_PREFIX", "To Fill 2")).strip()
        print(
            f"Batch split rule: index < {bs} → \"{p1}\", else → \"{p2}\" "
            f"(total customers: {len(files_sorted)})."
        )
        print(
            f"This run: rows [{start}:{start + len(files_window)}] = {len(files_window)} file(s) to process."
        )
    else:
        leg = os.environ.get("SUPABASE_STORAGE_PREFIX", "").strip() or "(bucket root)"
        print(f"Legacy mode: SUPABASE_STORAGE_PREFIX = {leg}")
        print(
            f"This run: rows [{start}:{start + len(files_window)}] = {len(files_window)} file(s) to process."
        )

    sleep_between = _env_float("UPLOAD_SLEEP_SECONDS", 0.0)
    max_retries = _env_int("UPLOAD_MAX_RETRIES", 5)
    retry_base = _env_float("UPLOAD_RETRY_BASE_SECONDS", 1.5)

    ok = 0
    err = 0
    skipped_db = 0
    skipped_cp = 0
    for offset, item in enumerate(files_window):
        idx = start + offset
        parent = item["parent_folder"]
        key = item["local_relative_path"]
        if use_cp and key in done:
            skipped_cp += 1
            continue
        rel_norm = _norm_rel_path(key)
        if only_missing and rel_norm in existing_db:
            skipped_db += 1
            continue

        local_path = root / item["local_relative_path"]
        if not local_path.is_file():
            print(f"[skip missing] {local_path}", file=sys.stderr)
            err += 1
            continue

        # Flat under prefix: no per-customer folders — one object name per file (path parts joined with __).
        prefix = _upload_prefix_for_index(idx, args)
        rel = item["local_relative_path"].replace("\\", "/")
        storage_rel = "__".join(safe_segment(p) for p in rel.split("/") if p)
        if prefix:
            prefix_safe = "/".join(
                safe_segment(p) for p in prefix.replace("\\", "/").split("/") if p
            )
            storage_path = f"{prefix_safe}/{storage_rel}" if prefix_safe else storage_rel
        else:
            storage_path = storage_rel
        mime = item["mime_type"]
        data = local_path.read_bytes()

        rel_key = item["local_relative_path"].replace("\\", "/")
        row = {
            "parent_folder": parent,
            "file_name": rel_key,
            "file_ext": item["file_ext"],
            "size_bytes": item["size_bytes"],
            "storage_path": storage_path,
            "local_relative_path": item["local_relative_path"],
            "mime_type": mime,
        }

        for attempt in range(max_retries + 1):
            try:
                client.storage.from_(bucket).upload(
                    path=storage_path,
                    file=data,
                    file_options={
                        "content-type": mime,
                        "upsert": "true",
                    },
                )
                client.table("contract_files").upsert(
                    row,
                    on_conflict="parent_folder,file_name",
                ).execute()

                ok += 1
                if use_cp:
                    done.add(key)
                    save_upload_done(done)
                print(f"[ok] {storage_path}")
                break
            except Exception as e:
                last_exc = e
                if attempt < max_retries and _is_transient_upload_error(e):
                    delay = retry_base * (2**attempt) + random.uniform(0, 0.35)
                    print(
                        f"[retry {attempt + 1}/{max_retries}] {storage_path}: {e!s}",
                        file=sys.stderr,
                    )
                    time.sleep(delay)
                    continue
                err += 1
                print(f"[err] {local_path}: {e}", file=sys.stderr)
                break

        if sleep_between > 0:
            time.sleep(sleep_between)

    if only_missing:
        print(
            f"Done. Uploaded: {ok}, skipped (checkpoint): {skipped_cp}, "
            f"skipped (already in DB): {skipped_db}, errors: {err}"
        )
    else:
        print(
            f"Done. Uploaded: {ok}, skipped (checkpoint): {skipped_cp}, errors: {err}"
        )
    return 0 if err == 0 else 2


def main() -> int:
    p = argparse.ArgumentParser(description="Analyze local contracts tree, upload to Supabase.")
    sub = p.add_subparsers(dest="cmd", required=True)

    pa = sub.add_parser("analyze", help="Scan LOCAL_CONTRACTS_ROOT and write analysis_report.json")
    pa.add_argument(
        "--root",
        help="Override LOCAL_CONTRACTS_ROOT (default: env or ./contracts_data)",
    )
    pa.set_defaults(func=cmd_analyze)

    pu = sub.add_parser("upload", help="Upload files and upsert rows in contract_files")
    pu.add_argument("--root", help="Override LOCAL_CONTRACTS_ROOT")
    pu.add_argument(
        "--only-missing",
        action="store_true",
        help="Skip files whose path is already in contract_files (re-run after partial failures).",
    )
    pu.add_argument(
        "--no-batch-split",
        action="store_true",
        help="Do not split To Fill 1 / To Fill 2; use SUPABASE_STORAGE_PREFIX for every file.",
    )
    pu.add_argument(
        "--batch-size",
        type=int,
        default=None,
        help="First N files (sorted by folder name) go to batch 1 (default env UPLOAD_BATCH_SIZE or 600).",
    )
    pu.add_argument(
        "--batch1-prefix",
        default=None,
        help="Storage subfolder for first batch (default env BATCH1_PREFIX or 'To Fill 1').",
    )
    pu.add_argument(
        "--batch2-prefix",
        default=None,
        help="Storage subfolder for second batch (default env BATCH2_PREFIX or 'To Fill 2').",
    )
    pu.add_argument(
        "--start",
        type=int,
        default=0,
        help="0-based index into sorted-by-folder list. Use 600 to start batch 2 (index 600+).",
    )
    pu.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Max files to upload this run (e.g. 600 for first batch only). Omit = no cap.",
    )
    pu.set_defaults(func=cmd_upload)

    args = p.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
