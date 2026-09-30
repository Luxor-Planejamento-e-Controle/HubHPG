-- Fila do hub → Azure. A cada pedido novo em hub_job, o banco avisa a rota HTTP do
-- Function App (luxor-planejamento-functions/hpg_hub_job), que enfileira e responde na
-- hora; quem roda o pipeline é o consumidor da fila de lá. Até 30/09/2026 só o agente
-- no notebook do Arthur executava: com ele desligado, o pedido ficava parado.
--
-- O aviso é assíncrono (pg_net manda depois do commit) e só um empurrão: se ele se
-- perder, a varredura da Azure (a cada 10 min) pega o pedido parado na fila.

create extension if not exists pg_net;

-- Endereço e chave da rota ficam no Vault, nunca aqui: o repo é público, e a chave
-- enfileira pedidos. Quem grava é hub_job_define_rota(), só pela service_role.
create or replace function public.hub_job_define_rota(p_url text, p_chave text)
returns void
language plpgsql security definer set search_path = public, vault as $$
begin
  delete from vault.secrets where name in ('hpg_hub_job_url', 'hpg_hub_job_chave');
  perform vault.create_secret(p_url, 'hpg_hub_job_url');
  perform vault.create_secret(p_chave, 'hpg_hub_job_chave');
end $$;

revoke all on function public.hub_job_define_rota(text, text) from public, anon, authenticated;
grant execute on function public.hub_job_define_rota(text, text) to service_role;

create or replace function public.hub_job_avisa_azure() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  u text;
  k text;
begin
  select decrypted_secret into u from vault.decrypted_secrets where name = 'hpg_hub_job_url';
  select decrypted_secret into k from vault.decrypted_secrets where name = 'hpg_hub_job_chave';
  if u is null or k is null then
    return new;       -- rota não configurada: o pedido espera na fila
  end if;
  perform net.http_post(
    url := u,
    body := jsonb_build_object('id', new.id, 'tipo', new.tipo),
    headers := jsonb_build_object('Content-Type', 'application/json',
                                  'x-functions-key', k),
    timeout_milliseconds := 5000);
  return new;
exception when others then
  -- o aviso nunca impede o pedido de ser gravado: a varredura cobre
  raise warning 'hub_job_avisa_azure: %', sqlerrm;
  return new;
end $$;

drop trigger if exists hub_job_avisa_azure on hub_job;
create trigger hub_job_avisa_azure after insert on hub_job
  for each row execute function public.hub_job_avisa_azure();
