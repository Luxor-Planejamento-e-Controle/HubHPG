-- Cadastro de naturezas do painel de Gastos da CASA (Fazenda Pao Grande), editável de
-- dentro da aba Gastos Casa ("Editar cadastro de naturezas"), como o do Haras.
--
-- O MESMO desenho da migration 20261014150000_gastos_naturezas, numa segunda linha da
-- mesma tabela: id 'naturezas_casa'. A rotina `publicar_gastos_casa` do Function App lê
-- esta linha (service_role) para montar o painel; enquanto ela estiver vazia ('{}'), a
-- rotina usa o cadastro versionado no repo Automacoes, como era antes.
--
-- Quem vê: hub_can('gastos_casa'). Quem grava: admin e os MESMOS editores do Haras
-- (gastos_editores — decisão da Controladoria, 16/10/2026), desde que vejam a aba da
-- Casa. As policies do Haras (gc_read / gc_update, id 'naturezas') ficam como estão: as
-- desta migration são outras, e o Postgres soma policies permissivas com OU.
--
-- O seed (as 66 naturezas e 24 códigos fora do escopo) NÃO entra aqui: o repo é público
-- e o conteúdo vai por sql/*.local.sql, aplicado à parte.

-- 1) a tabela passa a aceitar o segundo documento (a check nasceu `id = 'naturezas'`)
alter table gastos_cadastro drop constraint if exists gastos_cadastro_id_check;
alter table gastos_cadastro add constraint gastos_cadastro_id_check
  check (id in ('naturezas', 'naturezas_casa'));

-- 2) quem edita a Casa: os editores do Haras, passando pelo hub_can da aba da Casa (como o
--    hub_gastos_editor passa pelo de 'gastos'). hub_can já inclui o admin.
create or replace function public.hub_gastos_casa_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_can('gastos_casa') and (
    public.hub_is_admin() or exists (
      select 1 from gastos_editores where email = public.hub_email()
    )
  )
$$;

-- 3) policies próprias da linha da Casa
drop policy if exists gcc_read on gastos_cadastro;
create policy gcc_read on gastos_cadastro
  for select to authenticated
  using ( id = 'naturezas_casa' and public.hub_can('gastos_casa') );

drop policy if exists gcc_update on gastos_cadastro;
create policy gcc_update on gastos_cadastro
  for update to authenticated
  using      ( id = 'naturezas_casa' and public.hub_gastos_casa_editor() )
  with check ( id = 'naturezas_casa' and public.hub_gastos_casa_editor() );

-- 4) a linha nasce vazia; o seed a sobrescreve pela service_role / SQL editor
insert into gastos_cadastro (id, data) values ('naturezas_casa', '{}'::jsonb)
on conflict (id) do nothing;
