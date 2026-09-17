-- Plantel — (1) abrir e fechar mês pela tela e (2) movimentação lançada à mão.
--
-- ---------------------------------------------------------------------
-- 1) O mês fechado deixa de ser constante e vira REGISTRO.
--
-- Até aqui o corte morava em dois lugares que ninguém sincroniza: a constante
-- `FECHADO_ATE` no plantel.js e a função `plantel_mes_fechado`, com a data
-- escrita dentro. Fechar agosto exigia editar código, publicar e aplicar
-- migration — para uma decisão que é operacional, não de software. E reabrir um
-- mês para um ajuste exigia o mesmo caminho de volta.
--
-- Agora é uma linha por mês. O fallback para `m <= '2026-07'` fica: os meses
-- anteriores à existência desta tela foram fechados fora dela, e inventar linha
-- para eles seria afirmar um autor e uma data que não existiram.
-- ---------------------------------------------------------------------
create table if not exists plantel_mes_status (
  mes       text primary key check (mes ~ '^\d{4}-\d{2}$'),
  fechado   boolean     not null default false,
  autor     text,
  em        timestamptz not null default now()
);

comment on table plantel_mes_status is
  'Mês aberto ou fechado para registro de movimentação. Sem linha = vale o corte histórico.';

-- autor/data carimbados pelo banco: a mesma receita do plantel_mov_touch
create or replace function public.plantel_mes_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.em := now();
  new.autor := public.hub_email();
  return new;
end $$;

drop trigger if exists plantel_mes_touch on plantel_mes_status;
create trigger plantel_mes_touch before insert or update on plantel_mes_status
  for each row execute function public.plantel_mes_touch();

alter table plantel_mes_status enable row level security;

drop policy if exists pms_read on plantel_mes_status;
create policy pms_read on plantel_mes_status
  for select to authenticated using ( public.hub_can('plantel') );

-- Fechar e reabrir são a MESMA permissão: quem fecha tem de poder desfazer, ou
-- um clique errado vira chamado.
drop policy if exists pms_write on plantel_mes_status;
create policy pms_write on plantel_mes_status
  for insert to authenticated with check ( public.hub_can('plantel') );

drop policy if exists pms_update on plantel_mes_status;
create policy pms_update on plantel_mes_status
  for update to authenticated
  using ( public.hub_can('plantel') ) with check ( public.hub_can('plantel') );

-- A função passa a consultar a tabela. Era `immutable` (constante dentro do
-- corpo); lendo tabela tem de ser `stable`, senão o planner pode congelar o
-- resultado dentro da mesma transação e uma reabertura não valeria na hora.
create or replace function public.plantel_mes_fechado(m text) returns boolean
language sql stable as $$
  select coalesce( (select fechado from public.plantel_mes_status where mes = m),
                   m <= '2026-07' )
$$;

-- ---------------------------------------------------------------------
-- 2) Movimentação lançada à mão.
--
-- A tabela só sabia CLASSIFICAR um movimento que o cálculo já tinha achado no
-- diff entre dois meses — por isso não tem valor: o valor vinha do próprio
-- animal. Um lançamento manual não tem esse lastro (o animal comprado ainda nem
-- está no arquivo do haras), então carrega o valor consigo.
--
-- `valor` fica NULO no caso normal. Preenchido, a linha é um lançamento manual,
-- e a chave dela começa com 'MANUAL:' — prefixo impossível de colidir com a
-- chave de animal, que é NOME|LETRA.
-- ---------------------------------------------------------------------
alter table plantel_mov_classificacao add column if not exists valor numeric;

comment on column plantel_mov_classificacao.valor is
  'Só em lançamento manual (chave MANUAL:*). Nulo quando a classificação é de um movimento apurado.';
