# Local Contracts Upload

## Purpose

`local_contracts.py` scans a local customer contracts folder, chooses one contract file per customer, uploads it to Supabase Storage, and records metadata in the `contract_files` table.

This is the script used for the "pull files from folder and push selected contract PDFs/docs to Supabase" step.

## Expected Folder Layout

```text
LOCAL_CONTRACTS_ROOT/
  <NC_xxxx_CUSTOMER>/
    02. CONTRACT/
      *.pdf | *.doc | *.docx
```

Only files directly inside `02. CONTRACT` are considered. Subfolders below `02. CONTRACT` are ignored.

## Selection Logic

- Prefer PDF files over Word files.
- If multiple PDFs exist, prefer names containing terms like `Complete`, `Docusign`, `signed`, `firmado`, or `contrato firmado`.
- Deprioritize names containing terms like `copy`, `duplicate`, `do not edit`, or `borrador`.
- If no PDF exists, select one `.doc` or `.docx` file using similar scoring.
- Upload only one selected file per customer folder.

## Supabase Output

The script uploads the selected file to the configured Supabase Storage bucket and upserts a row into `public.contract_files`.

Storage paths are normalized to ASCII-safe keys so Spanish names and special characters do not break Supabase object uploads.

## Main Commands

```powershell
py local_contracts.py analyze
py local_contracts.py upload
py local_contracts.py upload --only-missing
py local_contracts.py upload --start 0 --limit 600
py local_contracts.py upload --start 600
```

## Environment Variables

The script reads `.env` values such as:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_BUCKET`
- `LOCAL_CONTRACTS_ROOT`
- `UPLOAD_BATCH_SIZE`
- `BATCH1_PREFIX`
- `BATCH2_PREFIX`

