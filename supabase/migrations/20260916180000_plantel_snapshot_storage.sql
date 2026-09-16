-- Plantel — o snapshot do mês passa a ser ARQUIVO no bucket, não só linha na tabela.
--
-- Por quê: a aba lia `plantel_snapshot` pelo PostgREST e trazia TODOS os meses de
-- uma vez. Isso é o Postgres lendo JSONB, convertendo pra texto e streamando pela
-- API — sem CDN e com trabalho de banco por byte. As outras abas do hub (semanal,
-- comitê, indicadores no P&C) são rápidas porque fazem outra coisa: URL assinada +
-- GET de objeto estático servido pela borda (ver assets/auth.js, loadData).
--
-- Medido em jul/2026: um mês são 391 linhas × 31 colunas = 131 KB de JSON, 24 KB
-- comprimido. O ano inteiro, 290 KB — o problema nunca foi volume, foi o caminho.
--
-- A tabela CONTINUA sendo gravada: ela é o registro durável (e o que já existe lá
-- segue valendo como fallback). O arquivo é a cópia rápida de leitura.
--
-- Nome: `plantel.<AAAA-MM>.json`. A policy de leitura usa
-- `split_part(name, '.', 1)`, então o prefixo antes do PRIMEIRO ponto continua
-- sendo `plantel` e o `hub_can('plantel')` de sempre decide quem baixa. Por isso
-- o separador é ponto e não hífen.

-- ---------------------------------------------------------------------
-- Escrita pela TELA. Até aqui o bucket só recebia do publish_hub.py, que usa a
-- service_role e ignora RLS — daí não existir policy de insert. Agora quem
-- importa o mês grava o arquivo do próprio navegador, então precisa de policy.
--
-- Mesma trava do registro de movimentação: mês fechado não recebe escrita. O mês
-- sai do nome do arquivo (`plantel.2026-08.json` -> `2026-08`), e é o banco que
-- confere — não a tela.
-- ---------------------------------------------------------------------
drop policy if exists hpg_plantel_write on storage.objects;
create policy hpg_plantel_write on storage.objects
  for insert to authenticated
  with check (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_can('plantel')
    and split_part(name, '.', 2) ~ '^\d{4}-\d{2}$'
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  );

-- Reimportar o mesmo mês sobrescreve: o upsert do arquivo é update, não insert.
drop policy if exists hpg_plantel_update on storage.objects;
create policy hpg_plantel_update on storage.objects
  for update to authenticated
  using (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_can('plantel')
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  )
  with check (
    bucket_id = 'hpg-data'
    and split_part(name, '.', 1) = 'plantel'
    and public.hub_can('plantel')
    and not public.plantel_mes_fechado( split_part(name, '.', 2) )
  );

-- Apagar não é operação de tela: sumir com o snapshot de um mês apaga o rastro do
-- fechamento. Quem precisar remove pelo painel, com a service_role.
