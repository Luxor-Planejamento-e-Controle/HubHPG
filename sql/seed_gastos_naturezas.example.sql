-- =====================================================================
-- Cadastro de naturezas do painel de Gastos — EXEMPLO.
--
-- Os valores abaixo são FICTÍCIOS. A lista real de naturezas, responsáveis e códigos do
-- Bimer é informação interna, e este repositório é público: o cadastro real vem de
-- `python cadastro_naturezas.py` (repositório Automacoes, pasta "Controle de Gasto do
-- HARAS"), que gera o JSON fora do Git, e o SQL resultante fica em
-- sql/seed_gastos_naturezas.local.sql (ignorado: *.local.sql).
--
-- Rodar UMA vez, depois da migration 20261014150000_gastos_naturezas. Depois disso o
-- documento é mantido pela aba "Naturezas" — rodar de novo APAGA o que foi editado nela.
-- =====================================================================

-- naturezas  o que entra no Controle. `bimer` = códigos numéricos do Bimer que contam para
--            ela; `path` = grupos da raiz ao pai; `ordem` de 10 em 10.
-- ignoradas  o que o Bimer traz e está FORA do escopo por decisão. Código que não está em
--            nenhuma das duas listas é PENDENTE e aparece na aba para alguém decidir.
update gastos_cadastro
   set data = $json$
{"versao":1,
 "naturezas":[
   {"ordem":10,"cod":"10.1.01","alt":"","nome":"Natureza de exemplo","resp":"Fulano",
    "bimer":[1001],"path":["Custos e Despesas","Custos","Grupo de exemplo"]}
 ],
 "ignoradas":[
   {"bimer":9001,"nome":"NATUREZA FORA DO ESCOPO (EXEMPLO)"}
 ]}
$json$::jsonb
 where id = 'naturezas';

-- Quem vai editar precisa da aba 'gastos' e de uma linha em gastos_editores (admin dispensa
-- as duas):
--   insert into gastos_editores (email) values ('fulano@exemplo.com');
--   insert into user_dashboard_access (email, dashboard) values ('fulano@exemplo.com', 'gastos');
