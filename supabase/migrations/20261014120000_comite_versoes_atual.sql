-- Versões do conteúdo do comitê, regra nova (pedido do Arthur, 14/10/2026).
--
-- O que estava errado:
--   * o rastro automático guardava o estado de ANTES de cada edição, e o cravar o
--     de agora. A lista misturava os dois e nenhuma linha dizia qual era a que o
--     deck estava usando.
--   * todo pedido de atualização do comitê cravava TODOS os meses com conteúdo.
--     Agosto ganhou v9 e v10, idênticas à v8, quando setembro foi gerado (09/10) e
--     atualizado (13/10).
--   * toda versão automática nascia cravada, então "cravada" não distinguia nada.
--
-- Regra daqui pra frente:
--   1) versão = o conteúdo DEPOIS de cada mudança. A mais nova é a que o deck usa.
--   2) conteúdo igual ao da última versão não cria versão (salvar sem mudar,
--      cravar duas vezes, pedido de atualização).
--   3) uma cravada por mês: a última criada. Criar ou remover versão passa a marca.
--   4) remover versão, pelo editor do comitê. Cravada não sai; a versão que o deck
--      usa só sai se houver outra igual a ela.

-- ---------------------------------------------------------------------
-- núcleo: grava uma versão nova, ou não grava nada se for repetida
-- ---------------------------------------------------------------------
create or replace function public.comite_versao_nova(
  p_mes text, p_conteudo jsonb, p_por text, p_rotulo text default null, p_rotula_igual boolean default false)
returns int language plpgsql security definer set search_path = public as $$
declare
  ult record;
  v int;
  rot text := nullif(btrim(coalesce(p_rotulo, '')), '');
begin
  select versao, conteudo into ult from comite_conteudo_versao
   where mes = p_mes order by versao desc limit 1;
  if found and ult.conteudo = p_conteudo then
    -- nada mudou: quem cravou com nome dá o nome à versão que já existe
    if rot is not null and p_rotula_igual then
      update comite_conteudo_versao set rotulo = rot where mes = p_mes and versao = ult.versao;
    end if;
    return ult.versao;
  end if;
  v := coalesce(ult.versao, 0) + 1;
  insert into comite_conteudo_versao (mes, versao, conteudo, cravada, rotulo, criado_por)
  values (p_mes, v, p_conteudo, true, rot, p_por);
  update comite_conteudo_versao set cravada = (versao = v)
   where mes = p_mes and cravada is distinct from (versao = v);
  return v;
end $$;

create or replace function public.comite_conteudo_json(p_mes text) returns jsonb
language sql stable security definer set search_path = public as $$
  select jsonb_build_object('comentarios', comentarios, 'exposicoes', exposicoes,
                            'manejo', manejo, 'fotos', fotos, 'pendencias', pendencias)
    from comite_conteudo where mes = p_mes
$$;

-- ---------------------------------------------------------------------
-- 1) o trigger guarda o estado NOVO (e o primeiro conteúdo do mês também)
-- ---------------------------------------------------------------------
create or replace function public.comite_conteudo_versiona() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  perform public.comite_versao_nova(new.mes, public.comite_conteudo_json(new.mes),
                                    coalesce(new.updated_by, 'atualização automática'));
  return null;
end $$;

drop trigger if exists comite_conteudo_versiona on comite_conteudo;
create trigger comite_conteudo_versiona after insert or update on comite_conteudo
  for each row execute function public.comite_conteudo_versiona();

-- ---------------------------------------------------------------------
-- cravar = dar nome à versão em uso (ou criar, se o conteúdo ainda não tiver)
-- ---------------------------------------------------------------------
create or replace function public.comite_cravar_versao(p_mes text, p_rotulo text default null)
returns int language plpgsql security definer set search_path = public as $$
declare j jsonb;
begin
  if not (public.hub_comite_editor() or auth.role() = 'service_role') then
    raise exception 'sem permissão para cravar versão';
  end if;
  j := public.comite_conteudo_json(p_mes);
  if j is null then
    raise exception 'mês % não tem conteúdo', p_mes;
  end if;
  -- a rotina automática não renomeia versão de ninguém
  return public.comite_versao_nova(p_mes, j, coalesce(public.hub_email(), 'atualização automática'),
                                   p_rotulo, public.hub_email() is not null);
end $$;

-- restaurar: o conteúdo volta e vira a versão nova, com o nome dizendo de onde veio
create or replace function public.comite_restaurar_versao(p_mes text, p_versao int)
returns void language plpgsql security definer set search_path = public as $$
declare j jsonb; v int;
begin
  if not (public.hub_comite_editor() or auth.role() = 'service_role') then
    raise exception 'sem permissão para restaurar versão';
  end if;
  select conteudo into j from comite_conteudo_versao where mes = p_mes and versao = p_versao;
  if j is null then
    raise exception 'versão % do mês % não existe', p_versao, p_mes;
  end if;
  select max(versao) into v from comite_conteudo_versao where mes = p_mes;
  update comite_conteudo set
    comentarios = coalesce(j->'comentarios', '[]'::jsonb),
    exposicoes  = coalesce(j->'exposicoes',  '{}'::jsonb),
    manejo      = coalesce(j->'manejo',      '[]'::jsonb),
    fotos       = coalesce(j->'fotos',       '[]'::jsonb),
    pendencias  = coalesce(j->'pendencias',  pendencias)
  where mes = p_mes;
  update comite_conteudo_versao set rotulo = 'restaurada da v' || p_versao
   where mes = p_mes and versao > v and rotulo is null;
end $$;

-- ---------------------------------------------------------------------
-- 4) remover
-- ---------------------------------------------------------------------
create or replace function public.comite_remover_versao(p_mes text, p_versao int)
returns void language plpgsql security definer set search_path = public as $$
declare j jsonb; agora jsonb; crav boolean;
begin
  if not (public.hub_comite_editor() or auth.role() = 'service_role') then
    raise exception 'sem permissão para remover versão';
  end if;
  select conteudo, cravada into j, crav from comite_conteudo_versao where mes = p_mes and versao = p_versao;
  if j is null then
    raise exception 'versão % do mês % não existe', p_versao, p_mes;
  end if;
  if crav then
    raise exception 'a v% está cravada e não pode ser removida', p_versao;
  end if;
  agora := public.comite_conteudo_json(p_mes);
  if j = agora and not exists (select 1 from comite_conteudo_versao
                                where mes = p_mes and versao <> p_versao and conteudo = agora) then
    raise exception 'a v% é a versão em uso no deck: restaure outra antes de removê-la', p_versao;
  end if;
  delete from comite_conteudo_versao where mes = p_mes and versao = p_versao;
  update comite_conteudo_versao
     set cravada = (versao = (select max(versao) from comite_conteudo_versao where mes = p_mes))
   where mes = p_mes;
end $$;

-- ---------------------------------------------------------------------
-- a lista da janela, com a versão em uso marcada
-- ---------------------------------------------------------------------
create or replace function public.comite_versoes(p_mes text)
returns table (versao int, cravada boolean, rotulo text, criado_por text, criado_em timestamptz, atual boolean)
language plpgsql stable security definer set search_path = public as $$
declare agora jsonb; v_atual int;
begin
  if not (public.hub_can('comite') or auth.role() = 'service_role') then
    raise exception 'sem acesso ao comitê';
  end if;
  agora := public.comite_conteudo_json(p_mes);
  select max(x.versao) into v_atual from comite_conteudo_versao x
   where x.mes = p_mes and x.conteudo = agora;
  return query
    select x.versao, x.cravada, x.rotulo, x.criado_por, x.criado_em, x.versao = v_atual
      from comite_conteudo_versao x
     where x.mes = p_mes
     order by x.versao desc;
end $$;

grant execute on function public.comite_versoes(text) to authenticated;
grant execute on function public.comite_remover_versao(text, int) to authenticated;
-- as duas internas não checam permissão: só o trigger e as funções acima chamam
revoke execute on function public.comite_versao_nova(text, jsonb, text, text, boolean) from public, anon, authenticated;
revoke execute on function public.comite_conteudo_json(text) from public, anon, authenticated;

-- ---------------------------------------------------------------------
-- o que já existe passa a obedecer a regra
-- ---------------------------------------------------------------------
-- mês cujo conteúdo em uso não é a última versão ganha a versão dele
do $$
declare r record;
begin
  for r in select mes, updated_by from comite_conteudo loop
    perform public.comite_versao_nova(r.mes, public.comite_conteudo_json(r.mes),
                                      coalesce(r.updated_by, 'atualização automática'));
  end loop;
end $$;

-- uma cravada por mês: a última
update comite_conteudo_versao x
   set cravada = (x.versao = m.ult)
  from (select mes, max(versao) ult from comite_conteudo_versao group by mes) m
 where x.mes = m.mes and x.cravada is distinct from (x.versao = m.ult);
