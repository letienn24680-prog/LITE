-- ============================================================================
-- LITE Document Workspace — Supabase Database Schema, Storage & Share RPCs
-- Chạy trực tiếp 1 lần trong Supabase Dashboard -> SQL Editor.
-- An toàn khi chạy lại nhiều lần (idempotent).
-- ============================================================================

-- 1. Bảng lưu trữ metadata và nội dung đã giải mã của tài liệu theo từng user
create table if not exists public.documents (
  id text not null,
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  title text not null,
  file_type text not null check (file_type in ('pptx', 'docx', 'xlsx', 'pdf')),
  file_name text,
  file_size text,
  file_size_bytes bigint,
  storage_path text,
  r2_object_key text,
  folder text,
  is_favorite boolean not null default false,
  is_shared boolean not null default false,
  page_count integer,
  content_preview text not null default '',
  parsed_content jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  last_accessed_at timestamptz not null default now(),
  primary key (user_id, id)
);

-- Bổ sung cột r2_object_key nếu bảng documents đã được tạo từ phiên bản trước
alter table public.documents
  add column if not exists r2_object_key text;

-- 2. Index phục vụ sắp xếp tài liệu mở/chỉnh sửa gần đây theo từng tài khoản
create index if not exists idx_documents_user_last_accessed
  on public.documents (user_id, last_accessed_at desc);

-- 3. Bật Row Level Security (RLS) cho bảng public.documents
alter table public.documents enable row level security;

drop policy if exists "documents_select_own" on public.documents;
create policy "documents_select_own"
  on public.documents
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "documents_insert_own" on public.documents;
create policy "documents_insert_own"
  on public.documents
  for insert
  to authenticated
  with check (auth.uid() = user_id);

drop policy if exists "documents_update_own" on public.documents;
create policy "documents_update_own"
  on public.documents
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

drop policy if exists "documents_delete_own" on public.documents;
create policy "documents_delete_own"
  on public.documents
  for delete
  to authenticated
  using (auth.uid() = user_id);

-- 4. Tạo Storage bucket riêng tư 'user-files' (giới hạn 50MB / tệp)
insert into storage.buckets (id, name, public, file_size_limit)
values ('user-files', 'user-files', false, 52428800)
on conflict (id) do update
set
  public = excluded.public,
  file_size_limit = excluded.file_size_limit;

-- 5. Chính sách RLS trên storage.objects cho bucket 'user-files'
-- Mỗi tài khoản chỉ được đọc/ghi/cập nhật/xóa trong thư mục {user_id}/... của chính mình
drop policy if exists "user_files_select_own" on storage.objects;
create policy "user_files_select_own"
  on storage.objects
  for select
  to authenticated
  using (
    bucket_id = 'user-files'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "user_files_insert_own" on storage.objects;
create policy "user_files_insert_own"
  on storage.objects
  for insert
  to authenticated
  with check (
    bucket_id = 'user-files'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "user_files_update_own" on storage.objects;
create policy "user_files_update_own"
  on storage.objects
  for update
  to authenticated
  using (
    bucket_id = 'user-files'
    and (storage.foldername(name))[1] = auth.uid()::text
  )
  with check (
    bucket_id = 'user-files'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

drop policy if exists "user_files_delete_own" on storage.objects;
create policy "user_files_delete_own"
  on storage.objects
  for delete
  to authenticated
  using (
    bucket_id = 'user-files'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

-- 6. Bảng lưu trữ liên kết chia sẻ công khai / có mật khẩu (document_shares)
create table if not exists public.document_shares (
  share_token text primary key,
  document_id text not null,
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  title text not null,
  file_type text not null check (file_type in ('pptx', 'docx', 'xlsx', 'pdf')),
  file_name text,
  file_size text,
  r2_object_key text,
  storage_path text,
  access_level text not null default 'private' check (access_level in ('private', 'link', 'password')),
  password_hash text,
  allow_download boolean not null default true,
  snapshot_payload jsonb not null default '{}'::jsonb,
  view_count integer not null default 0,
  revoked_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, document_id)
);

-- 7. Bật RLS cho public.document_shares (chỉ chủ sở hữu được truy vấn/sửa trực tiếp)
alter table public.document_shares enable row level security;

drop policy if exists "document_shares_select_own" on public.document_shares;
create policy "document_shares_select_own"
  on public.document_shares
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "document_shares_insert_own" on public.document_shares;
create policy "document_shares_insert_own"
  on public.document_shares
  for insert
  to authenticated
  with check (auth.uid() = user_id);

drop policy if exists "document_shares_update_own" on public.document_shares;
create policy "document_shares_update_own"
  on public.document_shares
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

drop policy if exists "document_shares_delete_own" on public.document_shares;
create policy "document_shares_delete_own"
  on public.document_shares
  for delete
  to authenticated
  using (auth.uid() = user_id);

-- 8. RPC công khai đọc metadata & snapshot khi link chia sẻ hợp lệ (SECURITY DEFINER)
-- Không trả snapshot_payload hay key tệp khi chưa mở khóa hoặc đã bị thu hồi / riêng tư.
create or replace function public.get_shared_document_meta(p_share_token text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_share public.document_shares%rowtype;
  v_next_views integer;
begin
  select *
  into v_share
  from public.document_shares
  where share_token = p_share_token
  limit 1;

  if not found then
    return jsonb_build_object(
      'status', 'not_found',
      'share_token', p_share_token
    );
  end if;

  if v_share.revoked_at is not null then
    return jsonb_build_object(
      'status', 'revoked',
      'share_token', v_share.share_token,
      'title', v_share.title,
      'file_type', v_share.file_type
    );
  end if;

  if v_share.access_level = 'private' then
    return jsonb_build_object(
      'status', 'private',
      'share_token', v_share.share_token,
      'title', v_share.title,
      'file_type', v_share.file_type
    );
  end if;

  if v_share.access_level = 'password' then
    return jsonb_build_object(
      'status', 'password_required',
      'share_token', v_share.share_token,
      'title', v_share.title,
      'file_type', v_share.file_type,
      'file_name', v_share.file_name,
      'file_size', v_share.file_size,
      'allow_download', v_share.allow_download,
      'view_count', coalesce(v_share.view_count, 0)
    );
  end if;

  update public.document_shares
  set view_count = coalesce(view_count, 0) + 1
  where share_token = v_share.share_token
  returning view_count into v_next_views;

  return jsonb_build_object(
    'status', 'ok',
    'share_token', v_share.share_token,
    'document_id', v_share.document_id,
    'title', v_share.title,
    'file_type', v_share.file_type,
    'file_name', v_share.file_name,
    'file_size', v_share.file_size,
    'allow_download', v_share.allow_download,
    'view_count', coalesce(v_next_views, 1),
    'r2_object_key', v_share.r2_object_key,
    'storage_path', v_share.storage_path,
    'snapshot_payload', coalesce(v_share.snapshot_payload, '{}'::jsonb)
  );
end;
$$;

-- 9. RPC xác thực mật khẩu (SHA-256 hex) cho link chia sẻ được bảo vệ
create or replace function public.unlock_shared_document(
  p_share_token text,
  p_password_hash text
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_share public.document_shares%rowtype;
  v_next_views integer;
begin
  select *
  into v_share
  from public.document_shares
  where share_token = p_share_token
  limit 1;

  if not found then
    return jsonb_build_object(
      'status', 'not_found',
      'share_token', p_share_token
    );
  end if;

  if v_share.revoked_at is not null then
    return jsonb_build_object(
      'status', 'revoked',
      'share_token', v_share.share_token,
      'title', v_share.title,
      'file_type', v_share.file_type
    );
  end if;

  if v_share.access_level = 'private' then
    return jsonb_build_object(
      'status', 'private',
      'share_token', v_share.share_token,
      'title', v_share.title,
      'file_type', v_share.file_type
    );
  end if;

  if v_share.access_level = 'password' then
    if coalesce(v_share.password_hash, '') = ''
      or coalesce(p_password_hash, '') <> v_share.password_hash then
      return jsonb_build_object(
        'status', 'password_invalid',
        'share_token', v_share.share_token,
        'title', v_share.title,
        'file_type', v_share.file_type,
        'file_name', v_share.file_name,
        'file_size', v_share.file_size,
        'allow_download', v_share.allow_download,
        'view_count', coalesce(v_share.view_count, 0)
      );
    end if;
  end if;

  update public.document_shares
  set view_count = coalesce(view_count, 0) + 1
  where share_token = v_share.share_token
  returning view_count into v_next_views;

  return jsonb_build_object(
    'status', 'ok',
    'share_token', v_share.share_token,
    'document_id', v_share.document_id,
    'title', v_share.title,
    'file_type', v_share.file_type,
    'file_name', v_share.file_name,
    'file_size', v_share.file_size,
    'allow_download', v_share.allow_download,
    'view_count', coalesce(v_next_views, 1),
    'r2_object_key', v_share.r2_object_key,
    'storage_path', v_share.storage_path,
    'snapshot_payload', coalesce(v_share.snapshot_payload, '{}'::jsonb)
  );
end;
$$;

revoke all on function public.get_shared_document_meta(text) from public;
grant execute on function public.get_shared_document_meta(text) to anon, authenticated;

revoke all on function public.unlock_shared_document(text, text) from public;
grant execute on function public.unlock_shared_document(text, text) to anon, authenticated;
