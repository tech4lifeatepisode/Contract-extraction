-- Run once in Supabase → SQL Editor (service role / dashboard).

create table if not exists public.contract_files (
  id uuid primary key default gen_random_uuid(),
  parent_folder text not null,
  file_name text not null,
  file_ext text,
  size_bytes bigint,
  storage_path text not null,
  local_relative_path text,
  mime_type text,
  created_at timestamptz not null default now(),
  unique (parent_folder, file_name)
);

alter table public.contract_files enable row level security;

-- Service role bypasses RLS; for anon/authenticated reads add policies as needed.
-- Example: allow authenticated read
-- create policy "read contract_files" on public.contract_files for select to authenticated using (true);

create index if not exists contract_files_parent_idx on public.contract_files (parent_folder);
