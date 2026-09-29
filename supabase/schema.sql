-- Esquema base para la futura versión multiusuario del Inventario TIC Praxis.
-- Ejecutar en un proyecto privado de Supabase/PostgreSQL.

create extension if not exists pgcrypto;

create table if not exists sedes (
  id uuid primary key default gen_random_uuid(),
  nombre text not null unique,
  activo boolean not null default true,
  created_at timestamptz not null default now()
);

create table if not exists colaboradores (
  id uuid primary key default gen_random_uuid(),
  dni text unique,
  nombres text not null,
  sede_id uuid references sedes(id),
  area text,
  cargo text,
  activo boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists equipos (
  id uuid primary key default gen_random_uuid(),
  -- Puede ser NULL durante la etapa de depuración. Cuando existe, debe ser único.
  codigo_tic text,
  codigo_padre_tic text,
  equipo text not null,
  descripcion text,
  marca text,
  modelo text,
  serie text,
  sede_id uuid references sedes(id),
  area text,
  responsable_id uuid references colaboradores(id),
  estado text,
  condicion text,
  situacion text,
  ubicacion_tipo text not null default 'SEDE'
    check (ubicacion_tipo in ('ASIGNADO','ALMACEN','SEDE')),
  observaciones text,
  origen text,
  origen_fila integer,
  necesita_revision boolean not null default false,
  ubicaciones_conflictivas jsonb not null default '[]'::jsonb,
  activo boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists equipos_codigo_tic_unique
  on equipos (upper(codigo_tic))
  where codigo_tic is not null and btrim(codigo_tic) <> '';

create unique index if not exists equipos_serie_unique
  on equipos (upper(serie))
  where serie is not null and btrim(serie) <> '';

create table if not exists movimientos (
  id uuid primary key default gen_random_uuid(),
  equipo_id uuid not null references equipos(id) on delete cascade,
  fecha timestamptz not null default now(),
  sede_origen_id uuid references sedes(id),
  area_origen text,
  responsable_origen_id uuid references colaboradores(id),
  sede_destino_id uuid references sedes(id),
  area_destino text,
  responsable_destino_id uuid references colaboradores(id),
  tipo text not null,
  motivo text,
  observaciones text,
  usuario_id uuid,
  created_at timestamptz not null default now()
);

create table if not exists importaciones (
  id uuid primary key default gen_random_uuid(),
  nombre_archivo text,
  registros_leidos integer not null default 0,
  registros_maestros integer not null default 0,
  codigos_consolidados integer not null default 0,
  conflictos_ubicacion integer not null default 0,
  registros_sin_codigo integer not null default 0,
  usuario_id uuid,
  created_at timestamptz not null default now()
);

-- Conserva cada aparición del Excel aunque luego varias filas se consoliden
-- en un único registro maestro de equipos.
create table if not exists apariciones_excel (
  id uuid primary key default gen_random_uuid(),
  importacion_id uuid references importaciones(id) on delete cascade,
  equipo_id uuid references equipos(id) on delete set null,
  codigo_tic text,
  hoja text not null,
  fila integer,
  sede_id uuid references sedes(id),
  area text,
  responsable_texto text,
  dni_texto text,
  estado text,
  fecha_texto text,
  observaciones text,
  datos_origen jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists equipos_sede_idx on equipos(sede_id);
create index if not exists equipos_responsable_idx on equipos(responsable_id);
create index if not exists movimientos_equipo_idx on movimientos(equipo_id, fecha desc);
create index if not exists apariciones_excel_codigo_idx on apariciones_excel(codigo_tic);
create index if not exists apariciones_excel_equipo_idx on apariciones_excel(equipo_id);

alter table sedes enable row level security;
alter table colaboradores enable row level security;
alter table equipos enable row level security;
alter table movimientos enable row level security;
alter table importaciones enable row level security;
alter table apariciones_excel enable row level security;

-- Las políticas deben definirse según los roles institucionales
-- (administrador_tic, editor, consulta).
