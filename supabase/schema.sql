-- Boiler Construction Management System
-- Run this whole file in Supabase > SQL Editor > New query.
create extension if not exists pgcrypto;
create table if not exists public.dpr_entries (
 id uuid primary key default gen_random_uuid(),
 user_id uuid not null references auth.users(id) on delete cascade,
 work_date date not null,
 department text not null check (department in ('Structure','Pressure Part','Non Pressure Part','Rotary Part','LP Piping','Trim Pipe','Bunker')),
 pg text default '', ma text default '', du text default '', description text default '',
 quantity numeric(14,3) not null default 0, weight numeric(14,3) not null default 0,
 pre_assembly numeric(14,3) not null default 0, placement numeric(14,3) not null default 0,
 alignment numeric(14,3) not null default 0, bolting numeric(14,3) not null default 0,
 remark text default '', created_at timestamptz not null default now()
);
create table if not exists public.ra_bills (
 id uuid primary key default gen_random_uuid(), user_id uuid not null references auth.users(id) on delete cascade,
 bill_month text not null, bill_year integer not null, quantity numeric(14,3) default 0,
 weight numeric(14,3) default 0, created_at timestamptz not null default now()
);
create table if not exists public.customer_bills (
 id uuid primary key default gen_random_uuid(), user_id uuid not null references auth.users(id) on delete cascade,
 bill_no text not null, customer_name text not null, bill_date date not null, description text default '',
 quantity numeric(14,3) default 0, rate numeric(14,2) default 0, amount numeric(14,2) default 0,
 gst numeric(5,2) default 18, gst_amount numeric(14,2) default 0, total_amount numeric(14,2) default 0,
 created_at timestamptz not null default now()
);
create table if not exists public.drawings (
 id uuid primary key default gen_random_uuid(), user_id uuid not null references auth.users(id) on delete cascade,
 file_name text not null, storage_path text not null, department text not null default 'General',
 created_at timestamptz not null default now()
);
alter table public.dpr_entries enable row level security;
alter table public.ra_bills enable row level security;
alter table public.customer_bills enable row level security;
alter table public.drawings enable row level security;
drop policy if exists "DPR owner access" on public.dpr_entries;
create policy "DPR owner access" on public.dpr_entries for all to authenticated using (auth.uid()=user_id) with check (auth.uid()=user_id);
drop policy if exists "RA owner access" on public.ra_bills;
create policy "RA owner access" on public.ra_bills for all to authenticated using (auth.uid()=user_id) with check (auth.uid()=user_id);
drop policy if exists "Customer bill owner access" on public.customer_bills;
create policy "Customer bill owner access" on public.customer_bills for all to authenticated using (auth.uid()=user_id) with check (auth.uid()=user_id);
drop policy if exists "Drawing owner access" on public.drawings;
create policy "Drawing owner access" on public.drawings for all to authenticated using (auth.uid()=user_id) with check (auth.uid()=user_id);
insert into storage.buckets(id,name,public,file_size_limit,allowed_mime_types)
values ('drawings','drawings',false,52428800,array['application/pdf','image/png','image/jpeg','application/acad','application/octet-stream'])
on conflict (id) do nothing;
drop policy if exists "Private drawing access" on storage.objects;
create policy "Private drawing access" on storage.objects for all to authenticated
using (bucket_id='drawings' and (storage.foldername(name))[1]=auth.uid()::text)
with check (bucket_id='drawings' and (storage.foldername(name))[1]=auth.uid()::text);
