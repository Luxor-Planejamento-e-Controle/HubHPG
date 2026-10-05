-- SGPG: os animais do plantel. O prêmio passa a apontar para um animal do plantel em
-- vez de um nome digitado — decisão do Arthur em 05/10/2026, depois de o mesmo animal
-- aparecer com três grafias no histórico ("da PG" e "da Pao Grande", com e sem acento).
--
-- A lista vem do roster do controle mensal do haras (CONTROLE_DE_PLANTEL, aba
-- PLANTEL), pela mesma regra da contagem do fechamento semanal
-- (PGSemanalReport._plantel_por_status): sem embrião e receptora, uma linha por animal
-- mesmo quando o controle repete por cotista. Quem grava é só a sincronização
-- (tools/sync_sgpg_animais.py, no fim de cada fechamento semanal), pela função abaixo.
-- Nome de animal NÃO entra neste arquivo: o repo é público.

create table if not exists sgpg_animal (
  id              bigint generated always as identity primary key,
  -- nome do controle sem cotista, maiúsculo e sem acento: é a identidade
  chave           text not null unique,
  nome_plantel    text not null,
  -- como o SGPG mostra; nasce do nome do plantel e não é sobrescrito pela sincronização
  nome            text not null,
  categoria       text,
  mae             text,
  pai             text,
  local           text,
  status_plantel  text,
  -- estava no roster da última sincronização; quem saiu continua aqui para o histórico
  no_plantel      boolean not null default true,
  fonte           text,
  atualizado_em   timestamptz not null default now()
);
create index if not exists sgpg_animal_no_plantel on sgpg_animal (no_plantel);

alter table sgpg_premiacao add column if not exists animal_id bigint references sgpg_animal(id);
create index if not exists sgpg_premiacao_animal_id on sgpg_premiacao (animal_id);

alter table sgpg_animal enable row level security;
drop policy if exists sga_read on sgpg_animal;
create policy sga_read on sgpg_animal for select to authenticated
  using ( public.hub_can('sgpg') or public.hub_can('comite') );
-- sem policy de escrita: só a função de sincronização (service_role) grava

create or replace function public.sgpg_sincroniza_animais(linhas jsonb, origem text)
returns integer
language plpgsql security definer set search_path = public as $$
declare n integer;
begin
  insert into sgpg_animal as a (chave, nome_plantel, nome, categoria, mae, pai, local,
                                status_plantel, no_plantel, fonte, atualizado_em)
  select x.chave, x.nome_plantel, x.nome, x.categoria, x.mae, x.pai, x.local,
         x.status_plantel, true, origem, now()
    from jsonb_to_recordset(linhas) as x(chave text, nome_plantel text, nome text,
         categoria text, mae text, pai text, local text, status_plantel text)
  on conflict (chave) do update
     set nome_plantel = excluded.nome_plantel, categoria = excluded.categoria,
         mae = excluded.mae, pai = excluded.pai, local = excluded.local,
         status_plantel = excluded.status_plantel, no_plantel = true,
         fonte = excluded.fonte, atualizado_em = now();
  get diagnostics n = row_count;
  -- quem não veio neste roster saiu do plantel (vendido, óbito, doado)
  update sgpg_animal set no_plantel = false, atualizado_em = now()
   where no_plantel
     and chave not in (select x->>'chave' from jsonb_array_elements(linhas) x);
  return n;
end $$;

revoke execute on function public.sgpg_sincroniza_animais(jsonb, text) from public, anon, authenticated;
grant execute on function public.sgpg_sincroniza_animais(jsonb, text) to service_role;
