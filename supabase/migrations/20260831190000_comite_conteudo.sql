-- Conteúdo do comitê editável direto pelo hub (comentários, exposições, manejo,
-- fotos) — os slides que NÃO saem de planilha nenhuma (ver comentário em
-- tools/build_comite.py, seção "Conteúdo escrito à mão"). Até 31/08/2026 isso
-- vivia só em _docs/comite_conteudo.json, editado à mão por quem tinha o repo
-- aberto. Agora o Supabase é a fonte principal; o JSON vira só seed/backup.
--
-- DDL apenas. Nenhum e-mail real aqui — quem pode editar vai pra
-- sql/seed_allowlist.local.sql (gitignored), mesmo padrão da allowlist.

-- ---------------------------------------------------------------------
-- 1) Quem edita: allowlist PRÓPRIA, separada de quem só vê a aba comitê.
--    hub_can('comite') já cobre leitura; editar é mais restrito que ver.
-- ---------------------------------------------------------------------
create table if not exists comite_editores (
  email       text primary key references allowed_users(email) on update cascade on delete cascade,
  created_at  timestamptz not null default now()
);

alter table comite_editores enable row level security;

-- usuário vê a própria linha (pra saber se é editor); admin vê e mexe em tudo
drop policy if exists ce_self_select on comite_editores;
create policy ce_self_select on comite_editores
  for select to authenticated
  using ( email = public.hub_email() or public.hub_is_admin() );

drop policy if exists ce_admin_write on comite_editores;
create policy ce_admin_write on comite_editores
  for all to authenticated
  using ( public.hub_is_admin() ) with check ( public.hub_is_admin() );

-- admin edita por definição (é quem administra); demais precisam de linha aqui
create or replace function public.hub_comite_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_is_admin() or exists (
    select 1 from comite_editores where email = public.hub_email()
  )
$$;

-- ---------------------------------------------------------------------
-- 2) Conteúdo em si. 1 linha por mês, mesmo formato jsonb que já existe
--    no JSON local hoje (comentarios: lista de {cat,txt,delta}; exposicoes:
--    {programacao,resultados}; manejo: lista; fotos: lista de {tema,arquivos}
--    onde cada arquivo é o path no bucket comite-fotos, não mais caminho local).
-- ---------------------------------------------------------------------
create table if not exists comite_conteudo (
  mes          text primary key check (mes ~ '^\d{4}-\d{2}$'),
  comentarios  jsonb not null default '[]'::jsonb,
  exposicoes   jsonb not null default '{}'::jsonb,
  manejo       jsonb not null default '[]'::jsonb,
  fotos        jsonb not null default '[]'::jsonb,
  updated_by   text,
  updated_at   timestamptz not null default now()
);

-- updated_by/updated_at automáticos — ninguém escreve esses dois campos na mão,
-- então não dá pra mentir quem editou por último
create or replace function public.comite_conteudo_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.updated_at := now();
  new.updated_by := public.hub_email();
  return new;
end $$;

drop trigger if exists comite_conteudo_touch on comite_conteudo;
create trigger comite_conteudo_touch before insert or update on comite_conteudo
  for each row execute function public.comite_conteudo_touch();

alter table comite_conteudo enable row level security;

drop policy if exists cc_read on comite_conteudo;
create policy cc_read on comite_conteudo
  for select to authenticated
  using ( public.hub_can('comite') );

drop policy if exists cc_write on comite_conteudo;
create policy cc_write on comite_conteudo
  for all to authenticated
  using ( public.hub_comite_editor() ) with check ( public.hub_comite_editor() );

-- ---------------------------------------------------------------------
-- 3) Storage: bucket PRIVADO pras fotos do deck. Leitura = mesma regra da
--    aba comitê; escrita (upload/trocar/apagar) só editor.
-- ---------------------------------------------------------------------
insert into storage.buckets (id, name, public)
values ('comite-fotos', 'comite-fotos', false)
on conflict (id) do update set public = false;

drop policy if exists comite_fotos_read on storage.objects;
create policy comite_fotos_read on storage.objects
  for select to authenticated
  using ( bucket_id = 'comite-fotos' and public.hub_can('comite') );

drop policy if exists comite_fotos_insert on storage.objects;
create policy comite_fotos_insert on storage.objects
  for insert to authenticated
  with check ( bucket_id = 'comite-fotos' and public.hub_comite_editor() );

drop policy if exists comite_fotos_update on storage.objects;
create policy comite_fotos_update on storage.objects
  for update to authenticated
  using ( bucket_id = 'comite-fotos' and public.hub_comite_editor() )
  with check ( bucket_id = 'comite-fotos' and public.hub_comite_editor() );

drop policy if exists comite_fotos_delete on storage.objects;
create policy comite_fotos_delete on storage.objects
  for delete to authenticated
  using ( bucket_id = 'comite-fotos' and public.hub_comite_editor() );
