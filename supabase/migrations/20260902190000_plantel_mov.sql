-- Plantel / Movimentação — decisão humana do fechamento mensal.
--
-- O cálculo (tools/build_plantel_mov.py) explica sozinho a maior parte dos
-- movimentos, cruzando o diff de patrimônio mês a mês com o log de
-- MOVIMENTAÇÕES da própria planilha do haras. O que ele NÃO explica não pode
-- ser adivinhado: vai pra fila da aba e uma pessoa decide. Esta tabela é onde
-- essa decisão mora.
--
-- Por que banco e não arquivo: o protótipo (LuxorMonthlyP-CRoutines/PlantelHPG)
-- guardava isso num mov_classificacao.json local e perguntava no terminal. Em
-- 02/09/2026 o arquivo estava VAZIO e 676 movimentos (R$ 43,8M) seguiam sem
-- classificação — não havia onde fazer, e o que fosse feito não teria autor
-- nem data.
--
-- DDL apenas. Nenhum e-mail real aqui.

-- ---------------------------------------------------------------------
-- 1) Mês fechado não se reclassifica. A trava é no banco, não só na tela:
--    fechado até jul/2026 por decisão do controle (02/09/2026). Ao fechar
--    agosto, é esta função que muda.
-- ---------------------------------------------------------------------
create or replace function public.plantel_mes_fechado(m text) returns boolean
language sql immutable as $$ select m <= '2026-07' $$;

-- ---------------------------------------------------------------------
-- 2) A decisão. Uma linha por (mês, movimento). `chave` é a chave do animal
--    no cálculo (nome|letra, ou PSEUDO:RECEPTORAS para a linha agregada de
--    receptoras) ou LOG:<produto>|<data> quando o item da fila é uma linha do
--    log que não achou efeito no patrimônio.
-- ---------------------------------------------------------------------
create table if not exists plantel_mov_classificacao (
  mes        text        not null check (mes ~ '^\d{4}-\d{2}$'),
  chave      text        not null,
  classe     text        not null check (classe in (
                 'venda','morte','doacao','compra','producao','transferencia',
                 'manter','baixar','sem_efeito','corrigir_planilha')),
  nota       text,
  nome       text,                    -- nome no momento da decisão (o animal é renomeado)
  autor      text,
  criado_em  timestamptz not null default now(),
  primary key (mes, chave)
);

comment on table plantel_mov_classificacao is
  'Decisão humana sobre movimento do plantel que o cálculo não explicou.';

-- autor/criado_em automáticos: ninguém escreve na mão, então não dá pra mentir
-- quem classificou o quê (mesma receita do comite_conteudo_touch)
create or replace function public.plantel_mov_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.criado_em := now();
  new.autor := public.hub_email();
  return new;
end $$;

drop trigger if exists plantel_mov_touch on plantel_mov_classificacao;
create trigger plantel_mov_touch before insert or update on plantel_mov_classificacao
  for each row execute function public.plantel_mov_touch();

alter table plantel_mov_classificacao enable row level security;

drop policy if exists pmc_read on plantel_mov_classificacao;
create policy pmc_read on plantel_mov_classificacao
  for select to authenticated
  using ( public.hub_can('plantel') );

-- Escrever só quem tem a aba, e só em mês aberto. Apagar não é operação de
-- tela (sumiria o rastro de quem decidiu): corrigir = gravar outra classe por
-- cima, que é o que o upsert da aba faz.
drop policy if exists pmc_write on plantel_mov_classificacao;
create policy pmc_write on plantel_mov_classificacao
  for insert to authenticated
  with check ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) );

drop policy if exists pmc_update on plantel_mov_classificacao;
create policy pmc_update on plantel_mov_classificacao
  for update to authenticated
  using ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) )
  with check ( public.hub_can('plantel') and not public.plantel_mes_fechado(mes) );

-- ---------------------------------------------------------------------
-- 3) Liberar a aba pra alguém (o insert exige ser admin — ver hub_schema.sql):
--      insert into user_dashboard_access (email, dashboard)
--      values ('pessoa@dominio.com', 'plantel') on conflict do nothing;
-- ---------------------------------------------------------------------
