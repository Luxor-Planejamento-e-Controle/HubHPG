-- SGPG — Sistema de Gestão Pao Grande. Primeiro módulo: exposições e premiações.
--
-- O SGPG é o sistema da Pao Grande como um todo (haras, fazenda, casa), montado aos
-- poucos, módulo a módulo, dentro do hub. Este é o primeiro: o registro das exposições
-- e das premiações, com o histórico consultável. O comitê mensal lê daqui os slides de
-- programação e resultados — até 05/10/2026 isso era texto digitado dentro do conteúdo
-- de cada mês do comitê, copiado de um mês para o outro e sem histórico próprio.
--
-- Quem vê: hub_can('sgpg'). Quem grava: admin e sgpg_editores (mesma receita do
-- plantel_editores, migration 20260928150000). E-mail real NÃO entra aqui — vai no
-- sql/seed_allowlist.local.sql (gitignored).

-- ---------------------------------------------------------------------
-- 1) A aba. A check só vale na criação da tabela: aba nova recria a lista inteira
--    (a do sql/hub_schema.sql já tinha 'gastos', que as migrations não tinham).
-- ---------------------------------------------------------------------
alter table user_dashboard_access drop constraint if exists user_dashboard_access_dashboard_check;
alter table user_dashboard_access add constraint user_dashboard_access_dashboard_check
  check (dashboard in ('semanal','comite','auditoria','plantel','gastos','sgpg'));

-- ---------------------------------------------------------------------
-- 2) Quem edita
-- ---------------------------------------------------------------------
create table if not exists sgpg_editores (
  email       text primary key references allowed_users(email) on update cascade on delete cascade,
  created_at  timestamptz not null default now()
);

alter table sgpg_editores enable row level security;

drop policy if exists sge_self_select on sgpg_editores;
create policy sge_self_select on sgpg_editores
  for select to authenticated
  using ( email = public.hub_email() or public.hub_is_admin() );

drop policy if exists sge_admin_write on sgpg_editores;
create policy sge_admin_write on sgpg_editores
  for all to authenticated
  using ( public.hub_is_admin() ) with check ( public.hub_is_admin() );

create or replace function public.hub_sgpg_editor() returns boolean
language sql stable security definer set search_path = public as $$
  select public.hub_can('sgpg') and (
    public.hub_is_admin() or exists (
      select 1 from sgpg_editores where email = public.hub_email()
    )
  )
$$;

-- ---------------------------------------------------------------------
-- 3) Exposições e premiações
-- ---------------------------------------------------------------------
create table if not exists sgpg_exposicao (
  id            bigint generated always as identity primary key,
  nome          text not null,
  inicio        date not null,
  fim           date not null,
  -- só o mês é conhecido ("Junho/2026"): início e fim cobrem o mês e a tela mostra o mês
  so_mes        boolean not null default false,
  local         text,
  -- manual, como no relatório do haras (decisão do Arthur, 05/10/2026)
  status        text not null default 'Próxima'
                check (status in ('Próxima','Aguardando','Realizada','Cancelada')),
  observacao    text,
  criado_por    text,
  criado_em     timestamptz not null default now(),
  alterado_por  text,
  alterado_em   timestamptz not null default now(),
  check (fim >= inicio)
);
create index if not exists sgpg_exposicao_inicio on sgpg_exposicao (inicio);

create table if not exists sgpg_premiacao (
  id            bigint generated always as identity primary key,
  exposicao_id  bigint not null references sgpg_exposicao(id) on delete cascade,
  -- nome como o haras registra; o cadastro de animais do SGPG vem depois
  animal        text not null,
  premio        text not null,
  ordem         int not null default 0,
  criado_por    text,
  criado_em     timestamptz not null default now(),
  alterado_por  text,
  alterado_em   timestamptz not null default now()
);
create index if not exists sgpg_premiacao_expo on sgpg_premiacao (exposicao_id, ordem);
create index if not exists sgpg_premiacao_animal on sgpg_premiacao (lower(animal));

-- autoria posta pelo banco, como no hub_job: quem grava não escolhe
create or replace function public.sgpg_autoria() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    new.criado_por := public.hub_email();
    new.criado_em := now();
  end if;
  new.alterado_por := public.hub_email();
  new.alterado_em := now();
  return new;
end $$;

drop trigger if exists sgpg_exposicao_autoria on sgpg_exposicao;
create trigger sgpg_exposicao_autoria before insert or update on sgpg_exposicao
  for each row execute function public.sgpg_autoria();
drop trigger if exists sgpg_premiacao_autoria on sgpg_premiacao;
create trigger sgpg_premiacao_autoria before insert or update on sgpg_premiacao
  for each row execute function public.sgpg_autoria();

alter table sgpg_exposicao enable row level security;
alter table sgpg_premiacao enable row level security;

-- O build do comitê lê com a service_role, por fora da RLS.
drop policy if exists sgx_read on sgpg_exposicao;
create policy sgx_read on sgpg_exposicao for select to authenticated
  using ( public.hub_can('sgpg') );
drop policy if exists sgx_write on sgpg_exposicao;
create policy sgx_write on sgpg_exposicao for all to authenticated
  using ( public.hub_sgpg_editor() ) with check ( public.hub_sgpg_editor() );

drop policy if exists sgp_read on sgpg_premiacao;
create policy sgp_read on sgpg_premiacao for select to authenticated
  using ( public.hub_can('sgpg') );
drop policy if exists sgp_write on sgpg_premiacao;
create policy sgp_write on sgpg_premiacao for all to authenticated
  using ( public.hub_sgpg_editor() ) with check ( public.hub_sgpg_editor() );
