-- Duas coisas que o hub passou a precisar quando o conteúdo deixou de ser
-- escrito só pelo dono do repo:
--
--   1) HISTÓRICO do conteúdo do comitê. `comite_conteudo` tem uma linha por mês
--      e o editor grava por upsert: com Ana, Aline e Arthur editando o mesmo
--      mês, quem salva por último sobrescreve o anterior e não há de onde
--      voltar. O histórico é feito por TRIGGER, não pelo cliente — versão que
--      depende de o navegador lembrar de gravar é versão que um dia falta.
--
--   2) FILA de atualização. O pipeline (semanal e comitê) lê planilhas do Google
--      Drive montado em G:, então não roda em lugar nenhum além da máquina que
--      tem o Drive. O botão no hub não executa nada: ele enfileira um pedido
--      aqui, e o agente local (tools/agente_hub.py, tarefa agendada) executa e
--      devolve status e log. É o mesmo arranjo do ETL de Indicadores do P&C.

-- ---------------------------------------------------------------------
-- 1) Histórico do conteúdo do comitê
-- ---------------------------------------------------------------------
create table if not exists comite_conteudo_versao (
  id          bigint generated always as identity primary key,
  mes         text not null check (mes ~ '^\d{4}-\d{2}$'),
  versao      int  not null,
  conteudo    jsonb not null,          -- {comentarios, exposicoes, manejo, fotos}
  -- `cravada` = alguém marcou esta como a boa, com rótulo ("levada ao comitê").
  -- As demais são o rastro automático de cada salvar.
  cravada     boolean not null default false,
  rotulo      text,
  criado_por  text,
  criado_em   timestamptz not null default now(),
  unique (mes, versao)
);

create index if not exists cc_versao_mes on comite_conteudo_versao (mes, versao desc);

-- Guarda o estado ANTERIOR a cada gravação. AFTER UPDATE e não BEFORE: só
-- interessa versionar mudança que de fato entrou. Update que não altera nada
-- (salvar sem editar) não gera versão — senão o histórico vira ruído.
create or replace function public.comite_conteudo_versiona() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  anterior jsonb := jsonb_build_object(
    'comentarios', old.comentarios, 'exposicoes', old.exposicoes,
    'manejo', old.manejo, 'fotos', old.fotos);
  atual jsonb := jsonb_build_object(
    'comentarios', new.comentarios, 'exposicoes', new.exposicoes,
    'manejo', new.manejo, 'fotos', new.fotos);
begin
  if anterior = atual then
    return null;
  end if;
  insert into comite_conteudo_versao (mes, versao, conteudo, criado_por, criado_em)
  select old.mes,
         coalesce((select max(versao) from comite_conteudo_versao where mes = old.mes), 0) + 1,
         anterior, old.updated_by, old.updated_at;
  return null;
end $$;

drop trigger if exists comite_conteudo_versiona on comite_conteudo;
create trigger comite_conteudo_versiona after update on comite_conteudo
  for each row execute function public.comite_conteudo_versiona();

-- Cravar = congelar o estado ATUAL como versão nomeada. Diferente do rastro
-- automático, que guarda o estado de antes de cada edição.
create or replace function public.comite_cravar_versao(p_mes text, p_rotulo text default null)
returns int language plpgsql security definer set search_path = public as $$
declare
  v int;
  c comite_conteudo%rowtype;
begin
  if not public.hub_comite_editor() then
    raise exception 'sem permissão para cravar versão';
  end if;
  select * into c from comite_conteudo where mes = p_mes;
  if not found then
    raise exception 'mês % não tem conteúdo', p_mes;
  end if;
  select coalesce(max(versao), 0) + 1 into v from comite_conteudo_versao where mes = p_mes;
  insert into comite_conteudo_versao (mes, versao, conteudo, cravada, rotulo, criado_por)
  values (p_mes, v,
          jsonb_build_object('comentarios', c.comentarios, 'exposicoes', c.exposicoes,
                             'manejo', c.manejo, 'fotos', c.fotos),
          true, nullif(btrim(coalesce(p_rotulo, '')), ''), public.hub_email());
  return v;
end $$;

-- Restaurar devolve o conteúdo de uma versão pro mês. O que está lá hoje não se
-- perde: o trigger acima versiona o estado atual antes de ser substituído.
create or replace function public.comite_restaurar_versao(p_mes text, p_versao int)
returns void language plpgsql security definer set search_path = public as $$
declare j jsonb;
begin
  if not public.hub_comite_editor() then
    raise exception 'sem permissão para restaurar versão';
  end if;
  select conteudo into j from comite_conteudo_versao where mes = p_mes and versao = p_versao;
  if j is null then
    raise exception 'versão % do mês % não existe', p_versao, p_mes;
  end if;
  update comite_conteudo set
    comentarios = coalesce(j->'comentarios', '[]'::jsonb),
    exposicoes  = coalesce(j->'exposicoes',  '{}'::jsonb),
    manejo      = coalesce(j->'manejo',      '[]'::jsonb),
    fotos       = coalesce(j->'fotos',       '[]'::jsonb)
  where mes = p_mes;
end $$;

alter table comite_conteudo_versao enable row level security;

drop policy if exists ccv_read on comite_conteudo_versao;
create policy ccv_read on comite_conteudo_versao
  for select to authenticated
  using ( public.hub_can('comite') );

-- Escrita só pelo trigger e pelas funções (security definer). Ninguém insere
-- versão à mão: histórico que aceita escrita direta não serve de histórico.
drop policy if exists ccv_no_write on comite_conteudo_versao;
create policy ccv_no_write on comite_conteudo_versao
  for all to authenticated using ( false ) with check ( false );

-- ---------------------------------------------------------------------
-- 2) Fila de atualização
-- ---------------------------------------------------------------------
create table if not exists hub_job (
  id            bigint generated always as identity primary key,
  tipo          text not null check (tipo in ('semanal', 'comite')),
  status        text not null default 'fila' check (status in ('fila', 'rodando', 'ok', 'erro')),
  pedido_por    text,
  pedido_em     timestamptz not null default now(),
  iniciado_em   timestamptz,
  terminado_em  timestamptz,
  log           text,
  detalhe       jsonb
);

create index if not exists hub_job_fila on hub_job (status, pedido_em);

-- Um pedido ativo por tipo. Sem isto, clicar duas vezes põe dois runs do mesmo
-- pipeline pra rodar em paralelo sobre as mesmas planilhas.
create unique index if not exists hub_job_um_ativo
  on hub_job (tipo) where status in ('fila', 'rodando');

create or replace function public.hub_job_touch() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    new.pedido_por := public.hub_email();
    new.pedido_em := now();
    new.status := 'fila';           -- quem pede não escolhe o status
  end if;
  return new;
end $$;

drop trigger if exists hub_job_touch on hub_job;
create trigger hub_job_touch before insert on hub_job
  for each row execute function public.hub_job_touch();

alter table hub_job enable row level security;

-- Vê a fila quem vê o painel correspondente; pede atualização quem tem acesso a
-- ele. O agente local usa a service_role e passa por fora da RLS.
drop policy if exists hub_job_read on hub_job;
create policy hub_job_read on hub_job
  for select to authenticated
  using ( public.hub_can(tipo) );

drop policy if exists hub_job_insert on hub_job;
create policy hub_job_insert on hub_job
  for insert to authenticated
  with check ( public.hub_can(tipo) );
