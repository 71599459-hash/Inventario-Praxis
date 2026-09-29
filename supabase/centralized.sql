-- Inventario TIC Colegio Praxis — backend centralizado Supabase
-- Ejecutar en un proyecto Supabase privado.

create extension if not exists pgcrypto;

do $$
begin
  if not exists (select 1 from pg_type where typname = 'praxis_role') then
    create type praxis_role as enum ('ADMIN_TIC','EDITOR','CONSULTA');
  end if;
end
$$;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text,
  nombre text,
  role praxis_role not null default 'CONSULTA',
  activo boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.inventory_state (
  id integer primary key default 1 check (id = 1),
  state jsonb,
  version bigint not null default 0,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id)
);

insert into public.inventory_state(id,state,version)
values (1,null,0)
on conflict (id) do nothing;

create table if not exists public.inventory_sync (
  id integer primary key default 1 check (id = 1),
  version bigint not null default 0,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id)
);

insert into public.inventory_sync(id,version)
values (1,0)
on conflict (id) do nothing;

create table if not exists public.audit_events (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  user_id uuid references auth.users(id),
  action text not null,
  version_before bigint,
  version_after bigint,
  summary jsonb not null default '{}'::jsonb
);

create index if not exists audit_events_created_idx on public.audit_events(created_at desc);
create index if not exists audit_events_user_idx on public.audit_events(user_id, created_at desc);

create or replace function public.current_praxis_role()
returns praxis_role
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    (select role from public.profiles where id = auth.uid() and activo = true),
    'CONSULTA'::praxis_role
  )
$$;

create or replace function public.handle_new_praxis_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles(id,email,nombre,role)
  values (
    new.id,
    new.email,
    coalesce(new.raw_user_meta_data->>'nombre', split_part(new.email,'@',1)),
    'CONSULTA'
  )
  on conflict (id) do update
    set email=excluded.email,
        nombre=coalesce(nullif(public.profiles.nombre,''),excluded.nombre),
        updated_at=now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created_praxis on auth.users;
create trigger on_auth_user_created_praxis
after insert on auth.users
for each row execute function public.handle_new_praxis_user();

create or replace function public.bootstrap_first_admin()
returns boolean
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  v_email text;
  v_confirmed timestamptz;
begin
  if auth.uid() is null then
    raise exception 'AUTH_REQUIRED';
  end if;

  select email, email_confirmed_at
    into v_email, v_confirmed
  from auth.users
  where id = auth.uid();

  if v_email is null or lower(v_email) not like '%@praxis.edu.pe' then
    return false;
  end if;

  if v_confirmed is null then
    return false;
  end if;

  if exists(select 1 from public.profiles where role='ADMIN_TIC' and activo=true) then
    return false;
  end if;

  update public.profiles
     set role='ADMIN_TIC', updated_at=now()
   where id=auth.uid();

  insert into public.audit_events(user_id,action,summary)
  values(auth.uid(),'BOOTSTRAP_ADMIN',jsonb_build_object('email',v_email));

  return true;
end;
$$;

create or replace function public.save_inventory_state(
  p_payload jsonb,
  p_action text default 'ACTUALIZAR_INVENTARIO',
  p_expected_version bigint default null
)
returns table(new_version bigint, updated_at timestamptz)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_role praxis_role;
  v_before bigint;
  v_after bigint;
  v_updated timestamptz;
begin
  if auth.uid() is null then
    raise exception 'AUTH_REQUIRED';
  end if;

  select public.current_praxis_role() into v_role;
  if v_role not in ('ADMIN_TIC','EDITOR') then
    raise exception 'READ_ONLY_ROLE';
  end if;

  select version into v_before
  from public.inventory_state
  where id = 1
  for update;

  if p_expected_version is not null and p_expected_version <> v_before then
    raise exception 'VERSION_CONFLICT expected %, current %', p_expected_version, v_before;
  end if;

  v_after := v_before + 1;
  v_updated := now();

  update public.inventory_state
     set state = p_payload,
         version = v_after,
         updated_at = v_updated,
         updated_by = auth.uid()
   where id = 1;

  update public.inventory_sync
     set version = v_after,
         updated_at = v_updated,
         updated_by = auth.uid()
   where id = 1;

  insert into public.audit_events(user_id,action,version_before,version_after,summary)
  values (
    auth.uid(),
    coalesce(nullif(p_action,''),'ACTUALIZAR_INVENTARIO'),
    v_before,
    v_after,
    jsonb_build_object(
      'masterRecords', jsonb_array_length(coalesce(p_payload->'inventory','[]'::jsonb)),
      'webMovements', jsonb_array_length(coalesce(p_payload->'webMovements','[]'::jsonb)),
      'fileName', p_payload->>'fileName'
    )
  );

  return query select v_after, v_updated;
end;
$$;

create or replace function public.set_user_role(p_user uuid, p_role praxis_role)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if auth.uid() is null or public.current_praxis_role() <> 'ADMIN_TIC' then
    raise exception 'ADMIN_REQUIRED';
  end if;

  if p_user = auth.uid()
     and p_role <> 'ADMIN_TIC'
     and (select count(*) from public.profiles where role='ADMIN_TIC' and activo=true) <= 1 then
    raise exception 'LAST_ADMIN';
  end if;

  update public.profiles
     set role=p_role, updated_at=now()
   where id=p_user;

  insert into public.audit_events(user_id,action,summary)
  values(auth.uid(),'CAMBIAR_ROL',jsonb_build_object('target_user',p_user,'new_role',p_role));
end;
$$;

alter table public.profiles enable row level security;
alter table public.inventory_state enable row level security;
alter table public.inventory_sync enable row level security;
alter table public.audit_events enable row level security;

drop policy if exists profiles_read_self_or_admin on public.profiles;
create policy profiles_read_self_or_admin on public.profiles
for select to authenticated
using (id = auth.uid() or public.current_praxis_role() = 'ADMIN_TIC');

drop policy if exists profiles_update_admin on public.profiles;
create policy profiles_update_admin on public.profiles
for update to authenticated
using (public.current_praxis_role() = 'ADMIN_TIC')
with check (public.current_praxis_role() = 'ADMIN_TIC');

drop policy if exists inventory_read_authenticated on public.inventory_state;
create policy inventory_read_authenticated on public.inventory_state
for select to authenticated
using (true);

drop policy if exists inventory_sync_read_authenticated on public.inventory_sync;
create policy inventory_sync_read_authenticated on public.inventory_sync
for select to authenticated
using (true);

drop policy if exists audit_read_admin on public.audit_events;
create policy audit_read_admin on public.audit_events
for select to authenticated
using (public.current_praxis_role() = 'ADMIN_TIC');

revoke all on function public.bootstrap_first_admin() from public, anon;
revoke all on function public.save_inventory_state(jsonb,text,bigint) from public, anon;
revoke all on function public.set_user_role(uuid,praxis_role) from public, anon;
revoke all on function public.current_praxis_role() from public, anon;

grant usage on schema public to authenticated;
grant select on public.inventory_state to authenticated;
grant select on public.inventory_sync to authenticated;
grant select on public.profiles to authenticated;
grant select on public.audit_events to authenticated;
grant execute on function public.bootstrap_first_admin() to authenticated;
grant execute on function public.current_praxis_role() to authenticated;
grant execute on function public.save_inventory_state(jsonb,text,bigint) to authenticated;
grant execute on function public.set_user_role(uuid,praxis_role) to authenticated;

-- Realtime usa una tabla liviana de versión. Al recibir un cambio, el cliente
-- vuelve a leer inventory_state; así no se envía el JSON completo en el evento.
do $$
begin
  if exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='inventory_state'
  ) then
    alter publication supabase_realtime drop table public.inventory_state;
  end if;

  if not exists (
    select 1 from pg_publication_tables
    where pubname='supabase_realtime' and schemaname='public' and tablename='inventory_sync'
  ) then
    alter publication supabase_realtime add table public.inventory_sync;
  end if;
end
$$;
