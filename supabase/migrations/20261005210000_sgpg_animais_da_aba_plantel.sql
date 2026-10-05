-- SGPG: a lista de animais sai da base do próprio hub — a aba Plantel —, não do Drive.
--
-- A migration 20261005200000 enchia sgpg_animal por um script que lia o controle
-- mensal direto do Drive (tools/sync_sgpg_animais.py, no fim do fechamento semanal).
-- Errado para um sistema que tem base própria: o plantel do hub é o que quem fecha o
-- mês importa na aba Plantel (plantel_snapshot, uma linha por mês com a aba PLANTEL
-- inteira). Agora a lista é derivada dela, aqui no banco, e se refaz sozinha sempre
-- que a aba Plantel importa um mês.
--
-- Regra: do mês mais recente importado, os animais com STATUS PLANTEL = 'PLANTEL',
-- sem embrião e receptora, uma linha por animal (o controle repete o animal por
-- cotista: "(CARLA)" / "(EDUARDO)"). Quem sai do plantel fica na tabela com
-- no_plantel = false — o prêmio antigo continua apontando para ele.

-- nome do controle sem o cotista: 'ORACAO DA PAO GRANDE (EDUARDO)' -> 'ORACAO DA PAO GRANDE'
create or replace function public.sgpg_sem_cotista(t text) returns text
language sql immutable as $$
  select regexp_replace(trim(regexp_replace(
           regexp_replace(upper(coalesce(t, '')), '\s*\((CARLA|EDUARDO)\)\s*', ' ', 'g'),
           '\s*\([^)]*\)\s*$', '')), '\s+', ' ', 'g')
$$;

-- identidade: sem cotista, maiúsculo, sem acento
create or replace function public.sgpg_chave_animal(t text) returns text
language sql immutable as $$
  select translate(public.sgpg_sem_cotista(t),
                   'ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ', 'AAAAAEEEEIIIIOOOOOUUUUCN')
$$;

-- 'NINA DO ATLANTICO SUL DA PAO GRANDE' -> 'Nina do Atlantico Sul da Pao Grande'
create or replace function public.sgpg_nome_exibicao(t text) returns text
language plpgsql immutable as $$
declare
  w text; partes text[] := '{}'; i int := 0;
begin
  foreach w in array regexp_split_to_array(trim(t), '\s+') loop
    i := i + 1;
    if w ~ '[0-9]' or upper(w) !~ '[AEIOU]' then
      partes := partes || upper(w);                         -- MH2, N19, PG, X
    elsif i > 1 and upper(w) in ('DA', 'DE', 'DO', 'DAS', 'DOS', 'E', 'DI') then
      partes := partes || lower(w);
    else
      partes := partes || initcap(lower(w));
    end if;
  end loop;
  return array_to_string(partes, ' ');
end $$;

create or replace function public.sgpg_animais_do_plantel() returns integer
language plpgsql security definer set search_path = public as $$
declare
  s record; n integer;
begin
  select mes, linhas into s from plantel_snapshot order by mes desc limit 1;
  if s is null then return 0; end if;

  -- todo mundo sai e quem está no plantel do mês volta: uma transação só
  update sgpg_animal set no_plantel = false where no_plantel;

  with ix as (
    select (s.linhas->'ix'->>'nome')::int      as nome,
           (s.linhas->'ix'->>'status')::int    as status,
           (s.linhas->'ix'->>'categoria')::int as categoria,
           (s.linhas->'ix'->>'mae')::int       as mae,
           (s.linhas->'ix'->>'pai')::int       as pai,
           (s.linhas->'ix'->>'local')::int     as local
  ), linhas as (
    select public.sgpg_sem_cotista(r->>ix.nome)       as nome_plantel,
           upper(trim(coalesce(r->>ix.status, '')))    as status,
           upper(trim(coalesce(r->>ix.categoria, ''))) as categoria,
           nullif(trim(r->>ix.mae), '')                as mae,
           nullif(trim(r->>ix.pai), '')                as pai,
           nullif(trim(r->>ix.local), '')              as local,
           ord
      from ix, jsonb_array_elements(s.linhas->'rows') with ordinality as t(r, ord)
  ), plantel as (
    select distinct on (chave) *
      from (select l.*, public.sgpg_chave_animal(l.nome_plantel) as chave
              from linhas l
             where l.status = 'PLANTEL' and l.categoria not in ('EMBRIAO', 'RECEPTORA')
               and l.nome_plantel <> '') x
     order by chave, ord
  )
  insert into sgpg_animal (chave, nome_plantel, nome, categoria, mae, pai, local,
                           status_plantel, no_plantel, fonte, atualizado_em)
  select chave, nome_plantel, public.sgpg_nome_exibicao(nome_plantel), categoria, mae, pai,
         local, status, true, 'aba Plantel · ' || s.mes, now()
    from plantel
  on conflict (chave) do update
     set nome_plantel = excluded.nome_plantel, categoria = excluded.categoria,
         mae = excluded.mae, pai = excluded.pai, local = excluded.local,
         status_plantel = excluded.status_plantel, no_plantel = true,
         fonte = excluded.fonte, atualizado_em = now();
  get diagnostics n = row_count;
  return n;
end $$;
revoke execute on function public.sgpg_animais_do_plantel() from public, anon, authenticated;

create or replace function public.sgpg_plantel_importado() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  perform public.sgpg_animais_do_plantel();
  return null;
end $$;

drop trigger if exists sgpg_plantel_importado on plantel_snapshot;
create trigger sgpg_plantel_importado after insert or update on plantel_snapshot
  for each statement execute function public.sgpg_plantel_importado();

-- a sincronização pelo Drive sai
drop function if exists public.sgpg_sincroniza_animais(jsonb, text);

-- o que veio do Drive e não tem prêmio sai; o resto se refaz pela aba Plantel
delete from sgpg_animal a
 where not exists (select 1 from sgpg_premiacao p where p.animal_id = a.id);
select public.sgpg_animais_do_plantel();
