-- Cravar e restaurar precisam funcionar TAMBÉM pela service_role, não só pelo
-- navegador de um editor logado: quem crava a versão antes de uma atualização
-- de dados é o agente local (tools/agente_hub.py), que não tem sessão de
-- usuário nenhuma — usa a service_role, como todo o resto do pipeline.
--
-- Sem isto o agente levava "sem permissão para cravar versão" e a proteção que
-- justifica o botão de atualizar não existiria justamente na hora de atualizar.

create or replace function public.comite_cravar_versao(p_mes text, p_rotulo text default null)
returns int language plpgsql security definer set search_path = public as $$
declare
  v int;
  c comite_conteudo%rowtype;
begin
  if not (public.hub_comite_editor() or auth.role() = 'service_role') then
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
          true, nullif(btrim(coalesce(p_rotulo, '')), ''),
          coalesce(public.hub_email(), 'atualização automática'));
  return v;
end $$;

create or replace function public.comite_restaurar_versao(p_mes text, p_versao int)
returns void language plpgsql security definer set search_path = public as $$
declare j jsonb;
begin
  if not (public.hub_comite_editor() or auth.role() = 'service_role') then
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
