-- Snapshot mensal do plantel, como veio do arquivo importado na aba.
--
-- O arquivo NÃO é puxado do Drive: a versão de trabalho do haras muda até o
-- fechamento sair, então quem fecha o mês importa o arquivo na tela. O que fica
-- guardado aqui é o estado daquele mês (todas as colunas, o log de ocorrências e
-- a atribuição Carla/Eduardo), que é o que o mês seguinte compara para descobrir
-- o que se movimentou.
--
-- `atribuicao` guarda de quem é cada animal (hpg = Carla, eduardo, nenhum). Isso
-- NÃO é fórmula: no mapa do Luxor a cota inteira vai para um dos dois e há
-- animais em nenhum — conferido em jul/2026, onde HPG soma R$ 15.970.552,71
-- (o saldo do Resumo Contábil liberado) e Eduardo R$ 588.750,00. Por isso a
-- atribuição é herdada do mês anterior e animal novo vai para a conciliação.

create table if not exists plantel_snapshot (
  mes            text primary key check (mes ~ '^\d{4}-\d{2}$'),
  arquivo        text,
  linhas         jsonb not null default '[]'::jsonb,   -- aba PLANTEL, todas as colunas
  log            jsonb not null default '[]'::jsonb,   -- aba MOVIMENTAÇÕES do haras
  atribuicao     jsonb not null default '{}'::jsonb,   -- chave -> hpg | eduardo | nenhum
  importado_por  text,
  importado_em   timestamptz not null default now()
);

comment on table plantel_snapshot is
  'Estado do plantel no fechamento de cada mês, do arquivo importado na aba Plantel.';

create or replace function public.plantel_snapshot_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.importado_em := now();
  new.importado_por := public.hub_email();
  return new;
end $$;

drop trigger if exists plantel_snapshot_touch on plantel_snapshot;
create trigger plantel_snapshot_touch before insert or update on plantel_snapshot
  for each row execute function public.plantel_snapshot_touch();

alter table plantel_snapshot enable row level security;

drop policy if exists ps_read on plantel_snapshot;
create policy ps_read on plantel_snapshot
  for select to authenticated using ( public.hub_can('plantel') );

-- Importar/reimportar é permitido enquanto o mês estiver aberto. Mês fechado
-- (ver plantel_mes_fechado, na migration 20260902190000) não se reimporta.
drop policy if exists ps_write on plantel_snapshot;
create policy ps_write on plantel_snapshot
  for insert to authenticated
  with check ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) );

drop policy if exists ps_update on plantel_snapshot;
create policy ps_update on plantel_snapshot
  for update to authenticated
  using ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) )
  with check ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) );

-- A semente (o mapa de julho, que traz a atribuição) é mês fechado, então entra
-- uma vez pelo SQL Editor e não pela tela:
--   update public.plantel_mes_fechado ... -- não: use insert direto como service_role,
--   ou libere temporariamente o mês para importar pela aba e feche de novo.
