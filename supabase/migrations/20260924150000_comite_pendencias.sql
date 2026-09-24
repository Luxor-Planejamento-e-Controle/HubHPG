-- Pendências da apresentação anterior: o terceiro slide do relatório da Ana
-- ("PENDÊNCIAS DA APRESENTAÇÃO DE JUNHO", com a lista do que ficou combinado).
-- É conteúdo escrito à mão, como comentários e manejo — então mora na mesma
-- linha do mês em `comite_conteudo`, editável pelo hub, e entra no histórico de
-- versões junto com o resto. Coluna nova com default: nada do que existe muda.

alter table comite_conteudo
  add column if not exists pendencias jsonb not null default '[]'::jsonb;

-- O histórico, o cravar e o restaurar passam a levar o campo. Sem isso, restaurar
-- uma versão apagaria as pendências (a versão não teria a chave) e o trigger não
-- versionaria uma edição que só mexeu nelas.
create or replace function public.comite_conteudo_versiona() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  anterior jsonb := jsonb_build_object(
    'comentarios', old.comentarios, 'exposicoes', old.exposicoes,
    'manejo', old.manejo, 'fotos', old.fotos, 'pendencias', old.pendencias);
  atual jsonb := jsonb_build_object(
    'comentarios', new.comentarios, 'exposicoes', new.exposicoes,
    'manejo', new.manejo, 'fotos', new.fotos, 'pendencias', new.pendencias);
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
                             'manejo', c.manejo, 'fotos', c.fotos, 'pendencias', c.pendencias),
          true, nullif(btrim(coalesce(p_rotulo, '')), ''),
          coalesce(public.hub_email(), 'atualização automática'));
  return v;
end $$;

-- Versão antiga (de antes desta coluna) não tem a chave: restaurar uma delas
-- mantém as pendências de hoje em vez de zerá-las.
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
    fotos       = coalesce(j->'fotos',       '[]'::jsonb),
    pendencias  = coalesce(j->'pendencias',  pendencias)
  where mes = p_mes;
end $$;
