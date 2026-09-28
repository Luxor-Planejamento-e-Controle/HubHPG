-- Plantel — quem MEXE no fechamento é mais restrito que quem VÊ a aba.
--
-- Até aqui toda escrita da aba (importar o mês, classificar, lançar à mão,
-- limpar, fechar/reabrir) exigia só hub_can('plantel'): quem tinha a aba podia
-- tudo. Decisão de 28/09/2026: movimentação só o admin e quem estiver em
-- plantel_editores; os demais com a aba são leitura.
--
-- Mesma receita do comite_editores (migration 20260831190000): allowlist
-- própria, admin edita por definição. E-mail real NÃO entra aqui — vai no
-- sql/seed_allowlist.local.sql (gitignored).
--
-- Até 28/09/2026 só o admin tinha gravado na aba, então fechar a porta não
-- deixa linha órfã de ninguém.

-- ---------------------------------------------------------------------
-- 1) Quem edita
-- ---------------------------------------------------------------------
create table if not exists plantel_editores (
  email       text primary key references allowed_users(email) on update cascade on delete cascade,
  created_at  timestamptz not null default now()
);

alter table plantel_editores enable row level security;

-- usuário vê a própria linha (pra tela saber se mostra os controles); admin vê e mexe em tudo
drop policy if exists pe_self_select on plantel_editores;
create policy pe_self_select on plantel_editores
  for select to authenticated
  using ( email = public.hub_email() or public.hub_is_admin() );

drop policy if exists pe_admin_write on plantel_editores;
create policy pe_admin_write on plantel_editores
  for all to authenticated
  using ( public.hub_is_admin() ) with check ( public.hub_is_admin() );

-- Diferente do hub_comite_editor, passa pelo hub_can: editor sem a aba (ou com
-- o acesso desativado em allowed_users.ativo) não escreve. hub_can já inclui o
-- admin.
create or replace function public.hub_plantel_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_can('plantel') and (
    public.hub_is_admin() or exists (
      select 1 from plantel_editores where email = public.hub_email()
    )
  )
$$;

-- ---------------------------------------------------------------------
-- 2) As policies de escrita trocam hub_can('plantel') por
--    hub_plantel_editor(). Leitura (pmc_read, ps_read, pms_read,
--    hpg_data_read) não muda: quem tem a aba continua vendo tudo.
-- ---------------------------------------------------------------------

-- classificação e lançamento manual
drop policy if exists pmc_write on plantel_mov_classificacao;
create policy pmc_write on plantel_mov_classificacao
  for insert to authenticated
  with check ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) );

drop policy if exists pmc_update on plantel_mov_classificacao;
create policy pmc_update on plantel_mov_classificacao
  for update to authenticated
  using ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) )
  with check ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) );

drop policy if exists pmc_delete on plantel_mov_classificacao;
create policy pmc_delete on plantel_mov_classificacao
  for delete to authenticated
  using ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) );

-- snapshot do mês (tabela)
drop policy if exists ps_write on plantel_snapshot;
create policy ps_write on plantel_snapshot
  for insert to authenticated
  with check ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) );

drop policy if exists ps_update on plantel_snapshot;
create policy ps_update on plantel_snapshot
  for update to authenticated
  using ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) )
  with check ( public.hub_plantel_editor() and not public.plantel_mes_fechado(mes) );

-- snapshot do mês (arquivo no bucket)
drop policy if exists hpg_plantel_write on storage.objects;
create policy hpg_plantel_write on storage.objects
  for insert to authenticated
  with check (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_plantel_editor()
    and split_part(name, '.', 2) ~ '^\d{4}-\d{2}$'
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  );

drop policy if exists hpg_plantel_update on storage.objects;
create policy hpg_plantel_update on storage.objects
  for update to authenticated
  using (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_plantel_editor()
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  )
  with check (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_plantel_editor()
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  );

-- fechar e reabrir mês
drop policy if exists pms_write on plantel_mes_status;
create policy pms_write on plantel_mes_status
  for insert to authenticated with check ( public.hub_plantel_editor() );

drop policy if exists pms_update on plantel_mes_status;
create policy pms_update on plantel_mes_status
  for update to authenticated
  using ( public.hub_plantel_editor() ) with check ( public.hub_plantel_editor() );
