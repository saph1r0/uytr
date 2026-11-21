-- NEXO OBRA (NO) · esquema de Supabase con roles: propietario, editor, lector.
-- Revit sigue siendo la fuente de los metrados; la nube guarda una copia para ver y editar desde otras computadoras.
-- Los editores registran metrados en la nube y el panel conectado a Revit los aplica al modelo (aplicado_revit).

create table if not exists proyectos (
  id uuid primary key default gen_random_uuid(),
  nombre text not null,
  modelo text not null,
  creado_por uuid not null default auth.uid() references auth.users(id),
  creado timestamptz not null default now()
);

create table if not exists miembros (
  proyecto_id uuid not null references proyectos(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  email text not null,
  rol text not null check (rol in ('propietario', 'editor', 'lector')),
  primary key (proyecto_id, user_id)
);

-- Datos del panel (presupuesto, elementos, cronograma). La geometría 3D va al bucket privado "geometria".
create table if not exists datos_base (
  proyecto_id uuid primary key references proyectos(id) on delete cascade,
  datos jsonb not null,
  cronograma jsonb,
  version timestamptz not null default now(),
  autor uuid default auth.uid()
);

create table if not exists registros (
  id bigint generated always as identity primary key,
  proyecto_id uuid not null references proyectos(id) on delete cascade,
  fecha date not null,
  element_id bigint not null default 0,          -- 0 = partida sin elementos en el modelo
  item text not null,
  cantidad numeric not null check (cantidad <> 0),
  origen text not null check (origen in ('revit', 'nube')),
  aplicado_revit boolean not null default false,
  autor uuid default auth.uid(),
  creado timestamptz not null default now()
);
create index if not exists registros_proyecto_fecha on registros (proyecto_id, fecha);

create table if not exists valorizaciones (
  proyecto_id uuid not null references proyectos(id) on delete cascade,
  n int not null,
  documento jsonb not null,                       -- cabecera, deducciones, metrados congelados
  estado text not null default 'Borrador' check (estado in ('Borrador', 'Presentada', 'Aprobada', 'Pagada')),
  actualizado timestamptz not null default now(),
  autor uuid default auth.uid(),
  primary key (proyecto_id, n)
);

create table if not exists sugerencias (
  id bigint generated always as identity primary key,
  proyecto_id uuid not null references proyectos(id) on delete cascade,
  pagina text, referencia text, texto text not null,
  estado text not null default 'Abierta' check (estado in ('Abierta', 'Atendida')),
  autor uuid default auth.uid(),
  creado timestamptz not null default now()
);

-- rol del usuario actual en un proyecto (null = no es miembro)
create or replace function rol_en(p uuid) returns text
language sql stable security definer set search_path = public as $$
  select rol from miembros where proyecto_id = p and user_id = auth.uid()
$$;

-- quien crea un proyecto queda como propietario
create or replace function nuevo_proyecto() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into miembros (proyecto_id, user_id, email, rol)
  values (new.id, new.creado_por, (select email from auth.users where id = new.creado_por), 'propietario');
  return new;
end $$;
drop trigger if exists al_crear_proyecto on proyectos;
create trigger al_crear_proyecto after insert on proyectos for each row execute function nuevo_proyecto();

-- el propietario invita por correo (la persona debe haber creado su cuenta antes)
create or replace function invitar(p uuid, correo text, nuevo_rol text) returns text
language plpgsql security definer set search_path = public as $$
declare u uuid;
begin
  if rol_en(p) is distinct from 'propietario' then raise exception 'Solo el propietario puede invitar.'; end if;
  if nuevo_rol not in ('editor', 'lector') then raise exception 'Rol no válido.'; end if;
  select id into u from auth.users where lower(email) = lower(correo);
  if u is null then return 'Esa persona aún no tiene cuenta: pídele que se registre y vuelve a invitarla.'; end if;
  insert into miembros (proyecto_id, user_id, email, rol) values (p, u, correo, nuevo_rol)
  on conflict (proyecto_id, user_id) do update set rol = excluded.rol;
  return 'Listo: ' || correo || ' es ' || nuevo_rol || '.';
end $$;

alter table proyectos enable row level security;
alter table miembros enable row level security;
alter table datos_base enable row level security;
alter table registros enable row level security;
alter table valorizaciones enable row level security;
alter table sugerencias enable row level security;

-- proyectos
create policy "ver proyecto" on proyectos for select using (rol_en(id) is not null);
create policy "crear proyecto" on proyectos for insert with check (creado_por = auth.uid());
-- miembros
create policy "ver miembros" on miembros for select using (rol_en(proyecto_id) is not null);
create policy "propietario quita miembros" on miembros for delete using (rol_en(proyecto_id) = 'propietario' and rol <> 'propietario');
-- datos del panel: todos leen; propietario y editores publican
create policy "ver datos" on datos_base for select using (rol_en(proyecto_id) is not null);
create policy "publicar datos" on datos_base for insert with check (rol_en(proyecto_id) in ('propietario', 'editor'));
create policy "actualizar datos" on datos_base for update using (rol_en(proyecto_id) in ('propietario', 'editor'));
-- registros de metrado
create policy "ver registros" on registros for select using (rol_en(proyecto_id) is not null);
create policy "registrar metrado" on registros for insert with check (rol_en(proyecto_id) in ('propietario', 'editor'));
create policy "marcar aplicado" on registros for update using (rol_en(proyecto_id) in ('propietario', 'editor'));
-- valorizaciones
create policy "ver valorizaciones" on valorizaciones for select using (rol_en(proyecto_id) is not null);
create policy "crear valorizacion" on valorizaciones for insert with check (rol_en(proyecto_id) in ('propietario', 'editor'));
create policy "editar valorizacion" on valorizaciones for update using (rol_en(proyecto_id) in ('propietario', 'editor'));
create policy "borrar valorizacion" on valorizaciones for delete using (rol_en(proyecto_id) in ('propietario', 'editor'));
-- sugerencias: cualquier miembro (también lectores) puede dejar una
create policy "ver sugerencias" on sugerencias for select using (rol_en(proyecto_id) is not null);
create policy "dejar sugerencia" on sugerencias for insert with check (rol_en(proyecto_id) is not null);
create policy "atender sugerencia" on sugerencias for update using (rol_en(proyecto_id) in ('propietario', 'editor'));

-- geometría 3D: bucket privado; leen los miembros, suben propietario y editores (ruta: <proyecto_id>/geometria.json)
insert into storage.buckets (id, name, public) values ('geometria', 'geometria', false) on conflict (id) do nothing;
create policy "leer geometria" on storage.objects for select using (bucket_id = 'geometria' and rol_en((storage.foldername(name))[1]::uuid) is not null);
create policy "subir geometria" on storage.objects for insert with check (bucket_id = 'geometria' and rol_en((storage.foldername(name))[1]::uuid) in ('propietario', 'editor'));
create policy "reemplazar geometria" on storage.objects for update using (bucket_id = 'geometria' and rol_en((storage.foldername(name))[1]::uuid) in ('propietario', 'editor'));
