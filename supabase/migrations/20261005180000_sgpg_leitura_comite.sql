-- O deck do comitê lê as exposições ao vivo do SGPG (assets/comite/deck.js,
-- buscaSgpg), com a sessão de quem está vendo. Quem tem a aba Comitê e não tem a do
-- SGPG veria a lista vazia pela RLS, e o deck trocaria o slide publicado por
-- "registre as exposições". É o mesmo conteúdo que o deck já mostra a essa pessoa:
-- ela só lê; gravar segue com hub_sgpg_editor().

drop policy if exists sgx_read on sgpg_exposicao;
create policy sgx_read on sgpg_exposicao for select to authenticated
  using ( public.hub_can('sgpg') or public.hub_can('comite') );

drop policy if exists sgp_read on sgpg_premiacao;
create policy sgp_read on sgpg_premiacao for select to authenticated
  using ( public.hub_can('sgpg') or public.hub_can('comite') );
