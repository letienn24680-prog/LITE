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


-- ============================================================================
-- 10. Cộng tác Thời gian thực (Google Docs/Sheets/Slides Collab + Realtime)
-- ============================================================================

-- Bổ sung cột permission ('view' | 'edit') vào bảng public.document_shares
alter table public.document_shares
  add column if not exists permission text not null default 'view'
  check (permission in ('view', 'edit'));

-- Bảng lưu danh sách tài liệu được chia sẻ với từng cộng tác viên ("Được chia sẻ với tôi")
create table if not exists public.document_collaborators (
  owner_id uuid not null references auth.users(id) on delete cascade,
  document_id text not null,
  collaborator_user_id uuid not null references auth.users(id) on delete cascade,
  share_token text not null,
  permission text not null default 'view' check (permission in ('view', 'edit')),
  last_opened_at timestamptz not null default now(),
  primary key (collaborator_user_id, owner_id, document_id)
);

create index if not exists idx_document_collaborators_user_opened
  on public.document_collaborators (collaborator_user_id, last_opened_at desc);

alter table public.document_collaborators enable row level security;

drop policy if exists "document_collaborators_select" on public.document_collaborators;
create policy "document_collaborators_select"
  on public.document_collaborators
  for select
  to authenticated
  using (auth.uid() = collaborator_user_id or auth.uid() = owner_id);

drop policy if exists "document_collaborators_insert_own" on public.document_collaborators;
create policy "document_collaborators_insert_own"
  on public.document_collaborators
  for insert
  to authenticated
  with check (auth.uid() = collaborator_user_id);

drop policy if exists "document_collaborators_update_own" on public.document_collaborators;
create policy "document_collaborators_update_own"
  on public.document_collaborators
  for update
  to authenticated
  using (auth.uid() = collaborator_user_id)
  with check (auth.uid() = collaborator_user_id);

drop policy if exists "document_collaborators_delete" on public.document_collaborators;
create policy "document_collaborators_delete"
  on public.document_collaborators
  for delete
  to authenticated
  using (auth.uid() = collaborator_user_id or auth.uid() = owner_id);

-- Cho phép cộng tác viên đã mở link hợp lệ theo dõi Realtime Postgres Changes của tài liệu
drop policy if exists "documents_select_shared_collaborator" on public.documents;
create policy "documents_select_shared_collaborator"
  on public.documents
  for select
  to authenticated
  using (
    exists (
      select 1
      from public.document_collaborators dc
      join public.document_shares ds
        on ds.share_token = dc.share_token
      where dc.collaborator_user_id = auth.uid()
        and dc.owner_id = documents.user_id
        and dc.document_id = documents.id
        and ds.revoked_at is null
        and ds.access_level <> 'private'
    )
  );

-- RPC giải mã liên kết cộng tác ?share=<token> trả về dữ liệu mới nhất từ public.documents
create or replace function public.resolve_collab_share(p_share_token text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_share public.document_shares%rowtype;
  v_doc public.documents%rowtype;
  v_has_doc boolean := false;
  v_next_views integer;
  v_perm text;
  v_uid uuid;
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

  v_perm := coalesce(v_share.permission, 'view');

  if v_share.revoked_at is not null then
    return jsonb_build_object(
      'status', 'revoked',
      'share_token', v_share.share_token,
      'owner_id', v_share.user_id,
      'document_id', v_share.document_id,
      'permission', v_perm
    );
  end if;

  if v_share.access_level = 'private' then
    return jsonb_build_object(
      'status', 'private',
      'share_token', v_share.share_token,
      'owner_id', v_share.user_id,
      'document_id', v_share.document_id,
      'permission', v_perm
    );
  end if;

  select *
  into v_doc
  from public.documents
  where user_id = v_share.user_id
    and id = v_share.document_id
  limit 1;

  v_has_doc := found;

  update public.document_shares
  set view_count = coalesce(view_count, 0) + 1
  where share_token = v_share.share_token
  returning view_count into v_next_views;

  v_uid := auth.uid();
  if v_uid is not null and v_uid <> v_share.user_id then
    insert into public.document_collaborators (
      owner_id,
      document_id,
      collaborator_user_id,
      share_token,
      permission,
      last_opened_at
    )
    values (
      v_share.user_id,
      v_share.document_id,
      v_uid,
      v_share.share_token,
      v_perm,
      now()
    )
    on conflict (collaborator_user_id, owner_id, document_id) do update
    set
      share_token = excluded.share_token,
      permission = excluded.permission,
      last_opened_at = now();
  end if;

  return jsonb_build_object(
    'status', 'ready',
    'share_token', v_share.share_token,
    'owner_id', v_share.user_id,
    'document_id', v_share.document_id,
    'permission', v_perm,
    'allow_download', v_share.allow_download,
    'view_count', coalesce(v_next_views, 1),
    'title', case when v_has_doc then v_doc.title else v_share.title end,
    'file_type', case when v_has_doc then v_doc.file_type else v_share.file_type end,
    'file_name', case when v_has_doc then coalesce(v_doc.file_name, v_share.file_name) else v_share.file_name end,
    'file_size', case when v_has_doc then coalesce(v_doc.file_size, v_share.file_size) else v_share.file_size end,
    'storage_path', case when v_has_doc then coalesce(v_doc.storage_path, v_share.storage_path) else v_share.storage_path end,
    'r2_object_key', case when v_has_doc then coalesce(v_doc.r2_object_key, v_share.r2_object_key) else v_share.r2_object_key end,
    'folder', case when v_has_doc then v_doc.folder else null end,
    'page_count', case when v_has_doc then v_doc.page_count else null end,
    'content_preview', case when v_has_doc then coalesce(v_doc.content_preview, '') else '' end,
    'parsed_content', case when v_has_doc then v_doc.parsed_content else null end,
    'snapshot_payload', coalesce(v_share.snapshot_payload, '{}'::jsonb),
    'updated_at', case when v_has_doc then v_doc.updated_at else v_share.updated_at end
  );
end;
$$;

-- RPC cho phép cộng tác viên (kể cả khách chưa đăng nhập có link edit) lưu bền vững nội dung lên DB
create or replace function public.update_shared_document_content(
  p_share_token text,
  p_title text,
  p_content_preview text,
  p_parsed_content jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_share public.document_shares%rowtype;
  v_now timestamptz := now();
  v_next_title text;
begin
  select *
  into v_share
  from public.document_shares
  where share_token = p_share_token
  limit 1;

  if not found then
    return jsonb_build_object('ok', false, 'error', 'Liên kết chia sẻ không tồn tại.');
  end if;

  if v_share.revoked_at is not null or v_share.access_level = 'private' then
    return jsonb_build_object('ok', false, 'error', 'Liên kết chia sẻ đã bị thu hồi.');
  end if;

  if coalesce(v_share.permission, 'view') <> 'edit' then
    return jsonb_build_object('ok', false, 'error', 'Liên kết này chỉ cho phép xem, không có quyền chỉnh sửa.');
  end if;

  v_next_title := coalesce(nullif(trim(p_title), ''), v_share.title);

  update public.documents
  set
    title = v_next_title,
    content_preview = coalesce(p_content_preview, content_preview),
    parsed_content = coalesce(p_parsed_content, parsed_content),
    updated_at = v_now
  where user_id = v_share.user_id
    and id = v_share.document_id;

  update public.document_shares
  set
    title = v_next_title,
    snapshot_payload = coalesce(v_share.snapshot_payload, '{}'::jsonb)
      || coalesce(p_parsed_content, '{}'::jsonb)
      || jsonb_build_object(
        'title', v_next_title,
        'contentPreview', coalesce(p_content_preview, '')
      ),
    updated_at = v_now
  where share_token = v_share.share_token;

  return jsonb_build_object(
    'ok', true,
    'updated_at', v_now
  );
end;
$$;

-- RPC lấy danh sách tài liệu được chia sẻ với người dùng đang đăng nhập
create or replace function public.get_shared_with_me_documents()
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_uid uuid := auth.uid();
  v_items jsonb;
begin
  if v_uid is null then
    return '[]'::jsonb;
  end if;

  select coalesce(
    jsonb_agg(
      jsonb_build_object(
        'share_token', ds.share_token,
        'owner_id', dc.owner_id,
        'document_id', dc.document_id,
        'permission', coalesce(ds.permission, dc.permission, 'view'),
        'last_opened_at', dc.last_opened_at,
        'title', coalesce(d.title, ds.title),
        'file_type', coalesce(d.file_type, ds.file_type),
        'file_name', coalesce(d.file_name, ds.file_name),
        'file_size', coalesce(d.file_size, ds.file_size),
        'storage_path', coalesce(d.storage_path, ds.storage_path),
        'r2_object_key', coalesce(d.r2_object_key, ds.r2_object_key),
        'folder', d.folder,
        'page_count', d.page_count,
        'content_preview', coalesce(d.content_preview, ''),
        'parsed_content', d.parsed_content,
        'snapshot_payload', coalesce(ds.snapshot_payload, '{}'::jsonb),
        'updated_at', coalesce(d.updated_at, ds.updated_at)
      )
      order by dc.last_opened_at desc
    ),
    '[]'::jsonb
  )
  into v_items
  from public.document_collaborators dc
  join public.document_shares ds
    on ds.share_token = dc.share_token
  left join public.documents d
    on d.user_id = dc.owner_id
   and d.id = dc.document_id
  where dc.collaborator_user_id = v_uid
    and dc.owner_id <> v_uid
    and ds.revoked_at is null
    and ds.access_level <> 'private';

  return coalesce(v_items, '[]'::jsonb);
end;
$$;

revoke all on function public.resolve_collab_share(text) from public;
grant execute on function public.resolve_collab_share(text) to anon, authenticated;

revoke all on function public.update_shared_document_content(text, text, text, jsonb) from public;
grant execute on function public.update_shared_document_content(text, text, text, jsonb) to anon, authenticated;

revoke all on function public.get_shared_with_me_documents() from public;
grant execute on function public.get_shared_with_me_documents() to authenticated;

-- Bật Supabase Realtime Postgres Changes cho bảng public.documents
do $$
begin
  alter publication supabase_realtime add table public.documents;
exception
  when duplicate_object then null;
  when undefined_object then null;
end
$$;
