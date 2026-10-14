-- Gastos — cadastro de naturezas do painel.
--
-- O painel de Gastos do Haras (publicado toda segunda pela rotina `publicar_gastos_haras`)
-- tem o realizado vindo dos relatórios do Bimer, mas "o que entra no Controle, em qual
-- grupo, de quem é e em que ordem" morava na planilha Controle Gastos — e natureza nova era
-- cirurgia no Excel. Quando uma natureza foi criada no Bimer (Material de Reprodução,
-- 09/2026) e ninguém a pôs na planilha, o dinheiro sumiu do painel em silêncio; só o check
-- contra o DRE denunciou (R$ 1.181,58).
--
-- Agora o cadastro mora aqui e é editado pela aba "Naturezas". A rotina do Function App o
-- lê com a service_role (por fora da RLS) para montar o painel.
--
-- Quem vê: hub_can('gastos'). Quem grava: admin e gastos_editores (mesma receita do
-- plantel_editores, migration 20260928150000). Não há aba nova na lista do
-- user_dashboard_access: a aba Naturezas segue a permissão de 'gastos'. E-mail real NÃO
-- entra aqui — vai no sql/seed_allowlist.local.sql (gitignored).
--
-- O conteúdo inicial (152 naturezas) NÃO está neste arquivo: o repo é público e a lista de
-- naturezas e responsáveis é dado interno. Vai por fora, num seed local — ver o PR.

-- ---------------------------------------------------------------------
-- 1) Quem edita
-- ---------------------------------------------------------------------
create table if not exists gastos_editores (
  email       text primary key references allowed_users(email) on update cascade on delete cascade,
  created_at  timestamptz not null default now()
);

alter table gastos_editores enable row level security;

drop policy if exists ge_self_select on gastos_editores;
create policy ge_self_select on gastos_editores
  for select to authenticated
  using ( email = public.hub_email() or public.hub_is_admin() );

drop policy if exists ge_admin_write on gastos_editores;
create policy ge_admin_write on gastos_editores
  for all to authenticated
  using ( public.hub_is_admin() ) with check ( public.hub_is_admin() );

-- passa pelo hub_can, como o hub_plantel_editor: editor sem a aba (ou com o acesso
-- desativado) não escreve. hub_can já inclui o admin.
create or replace function public.hub_gastos_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_can('gastos') and (
    public.hub_is_admin() or exists (
      select 1 from gastos_editores where email = public.hub_email()
    )
  )
$$;

-- ---------------------------------------------------------------------
-- 2) O cadastro: UM documento, e não uma linha por natureza.
--
--    { versao, naturezas: [{ordem, cod, alt, nome, resp, bimer:[int], path:[grupo..]}],
--      ignoradas: [{bimer, nome}] }
--
--    `naturezas` é o que entra no Controle; `ignoradas` é o que o Bimer traz e está FORA do
--    escopo por decisão. Código que não está em nenhuma das duas é PENDENTE — a tela lista
--    para alguém decidir. A validação (código, ordem e código do Bimer únicos; código não
--    pode estar no Controle E nas ignoradas) mora na tela e em cadastro_naturezas.py, do
--    repo Automacoes, que o Azure roda antes de publicar o painel.
--
--    Um documento só porque a ordem é o cadastro: trocar duas naturezas de lugar é uma
--    operação sobre o conjunto, que em linhas separadas exigiria unicidade adiada.
--    Em troca, a tela grava com concorrência otimista (UPDATE casa `updated_at`).
-- ---------------------------------------------------------------------
create table if not exists gastos_cadastro (
  id          text primary key check (id = 'naturezas'),
  data        jsonb not null default '{}'::jsonb,
  updated_at  timestamptz not null default now(),
  updated_por text
);

-- autoria posta pelo banco, como no semanal_manual: quem grava não escolhe. O `updated_at`
-- também: o cliente manda o dele, e o valor que vale é o que a resposta devolve.
create or replace function public.gastos_cadastro_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.updated_at := now();
  new.updated_por := public.hub_email();
  return new;
end $$;

drop trigger if exists gastos_cadastro_touch on gastos_cadastro;
create trigger gastos_cadastro_touch before update on gastos_cadastro
  for each row execute function public.gastos_cadastro_touch();

alter table gastos_cadastro enable row level security;

drop policy if exists gc_read on gastos_cadastro;
create policy gc_read on gastos_cadastro
  for select to authenticated
  using ( id = 'naturezas' and public.hub_can('gastos') );

drop policy if exists gc_update on gastos_cadastro;
create policy gc_update on gastos_cadastro
  for update to authenticated
  using      ( id = 'naturezas' and public.hub_gastos_editor() )
  with check ( id = 'naturezas' and public.hub_gastos_editor() );

-- Sem policy de insert nem de delete: a linha nasce aqui e só é alterada. O seed a
-- sobrescreve pela service_role / SQL editor.
insert into gastos_cadastro (id, data) values ('naturezas', '{}'::jsonb)
on conflict (id) do nothing;
