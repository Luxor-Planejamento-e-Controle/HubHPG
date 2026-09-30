-- Atualização semanal — dado que nenhuma planilha tem (doadoras ciclando).
--
-- Até 30/09/2026 a edição do dashboard ia para o localStorage de um navegador,
-- chaveada por semana: a semana seguinte abria vazia, ninguém mais via o número
-- e o pipeline nunca o recebia (lia só _cache/semanal_manual.json, preenchido à
-- mão). Agora o valor mora aqui e vale da semana em que foi salvo em diante, até
-- alguém salvar outro: o último salvo é o que vale.
--
-- Quem grava: admin e semanal_editores (mesma receita do plantel_editores,
-- migration 20260928150000). Quem tem a aba lê. E-mail real NÃO entra aqui —
-- vai no sql/seed_allowlist.local.sql (gitignored).

-- ---------------------------------------------------------------------
-- 1) Quem edita
-- ---------------------------------------------------------------------
create table if not exists semanal_editores (
  email       text primary key references allowed_users(email) on update cascade on delete cascade,
  created_at  timestamptz not null default now()
);

alter table semanal_editores enable row level security;

drop policy if exists se_self_select on semanal_editores;
create policy se_self_select on semanal_editores
  for select to authenticated
  using ( email = public.hub_email() or public.hub_is_admin() );

drop policy if exists se_admin_write on semanal_editores;
create policy se_admin_write on semanal_editores
  for all to authenticated
  using ( public.hub_is_admin() ) with check ( public.hub_is_admin() );

-- passa pelo hub_can, como o hub_plantel_editor: editor sem a aba não escreve
create or replace function public.hub_semanal_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_can('semanal') and (
    public.hub_is_admin() or exists (
      select 1 from semanal_editores where email = public.hub_email()
    )
  )
$$;

-- ---------------------------------------------------------------------
-- 2) O valor
-- ---------------------------------------------------------------------
create table if not exists semanal_manual (
  semana     date not null,              -- a sexta de referência da semana
  campo      text not null check (campo in ('doadoras_ciclando')),
  valor      numeric,                    -- null = vazio desta semana em diante
  salvo_por  text,
  salvo_em   timestamptz not null default now(),
  primary key (semana, campo)
);

-- autoria posta pelo banco, como no hub_job: quem grava não escolhe
create or replace function public.semanal_manual_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.salvo_por := public.hub_email();
  new.salvo_em := now();
  return new;
end $$;

drop trigger if exists semanal_manual_touch on semanal_manual;
create trigger semanal_manual_touch before insert or update on semanal_manual
  for each row execute function public.semanal_manual_touch();

alter table semanal_manual enable row level security;

-- O pipeline (agente local) lê com a service_role, por fora da RLS.
drop policy if exists sm_read on semanal_manual;
create policy sm_read on semanal_manual
  for select to authenticated
  using ( public.hub_can('semanal') );

drop policy if exists sm_insert on semanal_manual;
create policy sm_insert on semanal_manual
  for insert to authenticated
  with check ( public.hub_semanal_editor() );

drop policy if exists sm_update on semanal_manual;
create policy sm_update on semanal_manual
  for update to authenticated
  using ( public.hub_semanal_editor() ) with check ( public.hub_semanal_editor() );

drop policy if exists sm_delete on semanal_manual;
create policy sm_delete on semanal_manual
  for delete to authenticated
  using ( public.hub_semanal_editor() );
