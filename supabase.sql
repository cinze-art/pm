-- PEMUDA PEMUDI MAWE — JIMPITAN FLEXIBLE V3
-- Jalankan seluruh file ini di Supabase SQL Editor.
-- Schema ini mempertahankan data KK/agenda/organisasi lama dan menambahkan
-- program, jadwal, sesi penarikan, transaksi kumulatif, serta audit koreksi.

create extension if not exists pgcrypto;

create table if not exists public.rt (
  id int primary key,
  nama text not null
);
insert into public.rt(id,nama) values
(1,'RT 01'),(2,'RT 02'),(3,'RT 03'),(4,'RT 04'),(5,'RT 05'),(6,'RT 06'),(7,'HUNTAP')
on conflict (id) do update set nama=excluded.nama;

create table if not exists public.kk (
 id uuid primary key default gen_random_uuid(),
 nama text not null,
 rt int not null references public.rt(id),
 jenis text not null default 'umum',
 aktif boolean not null default true,
 created_at timestamptz not null default now()
);

-- Program tidak mengunci jenis/frekuensi.
create table if not exists public.jimpitan_program (
 id uuid primary key default gen_random_uuid(),
 nama_program text not null,
 jenis text not null,
 target_nominal numeric(14,2) not null default 0 check(target_nominal >= 0),
 tanggal_mulai date not null,
 tanggal_selesai date not null,
 frekuensi text not null check(frekuensi in ('weekly','biweekly','monthly','custom')),
 custom_dates jsonb not null default '[]'::jsonb,
 rt_ids jsonb not null default '[]'::jsonb,
 aktif boolean not null default true,
 created_by uuid references auth.users(id),
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now(),
 check(tanggal_selesai >= tanggal_mulai)
);

-- Satu tanggal + satu RT = satu jadwal.
create table if not exists public.jimpitan_schedule (
 id uuid primary key default gen_random_uuid(),
 program_id uuid not null references public.jimpitan_program(id) on delete restrict,
 tanggal date not null,
 rt int not null references public.rt(id),
 status text not null default 'scheduled' check(status in ('scheduled','active','finished','cancelled')),
 created_at timestamptz not null default now(),
 unique(program_id,tanggal,rt)
);

-- Penarik dibuat saat jadwal dimulai, bukan dipilih Admin.
create table if not exists public.penarikan_session (
 id uuid primary key default gen_random_uuid(),
 schedule_id uuid not null references public.jimpitan_schedule(id) on delete restrict,
 nama_penarik text not null,
 identitas text,
 started_at timestamptz not null default now(),
 finished_at timestamptz,
 status text not null default 'active' check(status in ('active','finished','cancelled')),
 created_at timestamptz not null default now()
);

-- Satu KK hanya satu transaksi untuk satu jadwal.
-- Nominal boleh 0 dan tetap menjadi transaksi yang tercatat.
create table if not exists public.jimpitan_transaction (
 id uuid primary key default gen_random_uuid(),
 schedule_id uuid not null references public.jimpitan_schedule(id) on delete restrict,
 session_id uuid not null references public.penarikan_session(id) on delete restrict,
 kk_id uuid not null references public.kk(id) on delete restrict,
 nominal numeric(14,2) not null default 0 check(nominal >= 0),
 confirmed_at timestamptz not null default now(),
 confirmed_by uuid references auth.users(id),
 updated_at timestamptz not null default now(),
 updated_by uuid references auth.users(id),
 unique(schedule_id,kk_id)
);

-- Audit koreksi: nilai lama tidak dihapus dari riwayat.
create table if not exists public.jimpitan_correction (
 id uuid primary key default gen_random_uuid(),
 transaction_id uuid not null references public.jimpitan_transaction(id) on delete restrict,
 old_nominal numeric(14,2) not null,
 new_nominal numeric(14,2) not null,
 reason text not null,
 corrected_by uuid not null references auth.users(id),
 corrected_at timestamptz not null default now()
);

create table if not exists public.agenda (
 id uuid primary key default gen_random_uuid(),
 tgl int not null,
 bulan text not null,
 tahun int not null,
 acara text not null,
 tempat text not null,
 jam text not null,
 deskripsi text,
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now()
);

create table if not exists public.organisasi (
 id int primary key default 1 check(id=1),
 ketua text, wakil text, sekretaris text, bendahara text,
 pembina text, sekretaris2 text, bendahara2 text,
 bidang jsonb not null default '{}'::jsonb,
 koordinator jsonb not null default '{}'::jsonb,
 visi text, misi jsonb not null default '[]'::jsonb,
 updated_at timestamptz not null default now()
);

alter table public.organisasi add column if not exists pembina text;
alter table public.organisasi add column if not exists sekretaris2 text;
alter table public.organisasi add column if not exists bendahara2 text;
alter table public.organisasi add column if not exists bidang jsonb not null default '{}'::jsonb;

insert into public.organisasi(id,ketua,wakil,sekretaris,bendahara,koordinator,visi,misi)
values(
1,'Budi Santoso','Siti Aminah','Parno','Agus',
'{"1":"Karman RT 01","2":"Jono RT 02","3":"Sutrisno RT 03","4":"Wagiman RT 04","5":"Supri RT 05","6":"Marno RT 06","7":"Tukijo HUNTAP"}',
'Terwujudnya Pemuda Pemudi Mawe yang Guyub, Rukun, Sejahtera',
'["Guyub","Rukun","Sejahtera"]'
)
on conflict (id) do nothing;

-- Indeks
create index if not exists kk_rt_idx on public.kk(rt);
create index if not exists kk_nama_idx on public.kk(nama);
create index if not exists schedule_tanggal_idx on public.jimpitan_schedule(tanggal);
create index if not exists schedule_program_idx on public.jimpitan_schedule(program_id);
create index if not exists schedule_rt_idx on public.jimpitan_schedule(rt);
create index if not exists transaction_kk_idx on public.jimpitan_transaction(kk_id);
create index if not exists transaction_schedule_idx on public.jimpitan_transaction(schedule_id);
create index if not exists transaction_confirmed_idx on public.jimpitan_transaction(confirmed_at);
create index if not exists correction_transaction_idx on public.jimpitan_correction(transaction_id);

-- RLS
alter table public.rt enable row level security;
alter table public.kk enable row level security;
alter table public.jimpitan_program enable row level security;
alter table public.jimpitan_schedule enable row level security;
alter table public.penarikan_session enable row level security;
alter table public.jimpitan_transaction enable row level security;
alter table public.jimpitan_correction enable row level security;
alter table public.agenda enable row level security;
alter table public.organisasi enable row level security;

-- Bersihkan policy V3 jika pernah dijalankan.
drop policy if exists rt_public_read on public.rt;
drop policy if exists kk_public_read on public.kk;
drop policy if exists kk_auth_write on public.kk;
drop policy if exists program_public_read on public.jimpitan_program;
drop policy if exists program_auth_write on public.jimpitan_program;
drop policy if exists schedule_public_read on public.jimpitan_schedule;
drop policy if exists schedule_auth_write on public.jimpitan_schedule;
drop policy if exists session_public_read on public.penarikan_session;
drop policy if exists session_public_insert on public.penarikan_session;
drop policy if exists session_auth_write on public.penarikan_session;
drop policy if exists transaction_public_read on public.jimpitan_transaction;
drop policy if exists transaction_public_insert on public.jimpitan_transaction;
drop policy if exists transaction_auth_write on public.jimpitan_transaction;
drop policy if exists correction_auth_all on public.jimpitan_correction;
drop policy if exists agenda_public_read on public.agenda;
drop policy if exists agenda_auth_write on public.agenda;
drop policy if exists org_public_read on public.organisasi;
drop policy if exists org_auth_write on public.organisasi;

create policy rt_public_read on public.rt for select using (true);
create policy kk_public_read on public.kk for select using (aktif=true);
create policy kk_auth_write on public.kk for all to authenticated using (true) with check (true);

create policy program_public_read on public.jimpitan_program for select using (true);
create policy program_auth_write on public.jimpitan_program for all to authenticated using (true) with check (true);

create policy schedule_public_read on public.jimpitan_schedule for select using (true);
create policy schedule_auth_write on public.jimpitan_schedule for all to authenticated using (true) with check (true);

-- Penarik adalah orang yang berada di lapangan, jadi sesi boleh dibuat oleh publik.
create policy session_public_read on public.penarikan_session for select using (true);
create policy session_public_insert on public.penarikan_session for insert with check (true);
create policy session_auth_write on public.penarikan_session for update to authenticated using (true) with check (true);

-- Transaksi lapangan boleh dibuat oleh penarik tanpa akun.
create policy transaction_public_read on public.jimpitan_transaction for select using (true);
create policy transaction_public_insert on public.jimpitan_transaction for insert with check (true);
create policy transaction_auth_write on public.jimpitan_transaction for update to authenticated using (true) with check (true);

-- Koreksi hanya Admin/Auth.
create policy correction_auth_all on public.jimpitan_correction for all to authenticated using (true) with check (true);

create policy agenda_public_read on public.agenda for select using (true);
create policy agenda_auth_write on public.agenda for all to authenticated using (true) with check (true);
create policy org_public_read on public.organisasi for select using (true);
create policy org_auth_write on public.organisasi for all to authenticated using (true) with check (true);

-- Fallback compatibility: tabel jimpitan_bulanan lama tidak dihapus.
-- Data lama tetap ada sebagai arsip jika tabel tersebut sudah pernah dibuat.
