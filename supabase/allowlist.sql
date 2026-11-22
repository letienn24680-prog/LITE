-- Chạy đoạn SQL này 1 lần trong Supabase -> SQL Editor
create table if not exists public.app_allowed_emails (
  email text primary key,
  status text not null default 'approved' check (status in ('approved', 'pending')),
  note text,
  added_by text,
  created_at timestamptz not null default now()
);

alter table public.app_allowed_emails enable row level security;

-- Cho phép kiểm tra trạng thái email & gửi yêu cầu xin cấp quyền từ app
drop policy if exists "Allow public read allowed emails" on public.app_allowed_emails;
create policy "Allow public read allowed emails"
  on public.app_allowed_emails for select
  using (true);

drop policy if exists "Allow public request pending email" on public.app_allowed_emails;
create policy "Allow public request pending email"
  on public.app_allowed_emails for insert
  with check (status = 'pending' or auth.role() = 'authenticated');

drop policy if exists "Allow authenticated manage emails" on public.app_allowed_emails;
create policy "Allow authenticated manage emails"
  on public.app_allowed_emails for all
  using (auth.role() = 'authenticated')
  with check (auth.role() = 'authenticated');
