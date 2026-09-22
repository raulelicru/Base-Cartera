-- Esquema de la app Base de Cartera (aplicado al proyecto Supabase "base-cartera").

-- Catálogo Nacional de Códigos Postales (1 fila por CP)
create table public.catalogo_cp (
  cp text primary key check (cp ~ '^\d{5}$'),
  municipio text,
  estado text,
  zona text,
  actualizado_en timestamptz not null default now()
);

-- Cada corrida de generación de la Base de Gestión
create table public.corridas (
  id bigint generated always as identity primary key,
  campania_trabajo integer not null,
  archivo_cartera text,
  total_filas integer not null default 0,
  filas_revision integer not null default 0,
  resumen jsonb not null default '{}'::jsonb,
  creado_en timestamptz not null default now()
);
create index corridas_campania_idx on public.corridas (campania_trabajo, creado_en desc);

-- Filas de la Base de Gestión (22 columnas de la especificación + revisión)
create table public.base_gestion (
  id bigint generated always as identity primary key,
  corrida_id bigint not null references public.corridas (id) on delete cascade,
  fila integer not null,
  zona text,
  region text,
  ruta text,
  division text,
  id_cobrador text,
  no_dama text,
  direccion text,
  direccion_calle text,
  colonia text,
  municipio_poblacion text,
  cp text,
  estado text,
  zona_urbano_rural text,
  referencia text,
  anio_saldo text,
  campania_saldo text,
  digito_verificador text,
  concatenado text,
  fecha_cierre date,
  morosidad text,
  campania_trabajo integer not null,
  referencia_pago text,
  requiere_revision boolean not null default false,
  motivo_revision text,
  unique (corrida_id, fila)
);
create index base_gestion_no_dama_idx on public.base_gestion (no_dama);
create index base_gestion_revision_idx on public.base_gestion (corrida_id) where requiere_revision;

-- RLS activo y sin políticas: sólo la llave secreta (service_role) del servidor Streamlit puede leer/escribir.
alter table public.catalogo_cp enable row level security;
alter table public.corridas enable row level security;
alter table public.base_gestion enable row level security;

-- Bucket privado para la Estructura General de Bases
insert into storage.buckets (id, name, public) values ('referencias', 'referencias', false)
on conflict (id) do nothing;
