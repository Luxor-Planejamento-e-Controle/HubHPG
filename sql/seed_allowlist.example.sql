-- Modelo do seed da allowlist. NÃO versionar a versão preenchida:
-- salve como sql/seed_allowlist.local.sql, que é gitignored.
--
-- Rodar DEPOIS de sql/hub_schema.sql. O primeiro registro tem de ser um admin,
-- senão não existe ninguém pra liberar os demais.

-- 1) quem entra no hub
insert into allowed_users (email, nome, role) values
  ('voce@luxor.com.br',   'Nome do Admin', 'admin'),
  ('fulano@luxor.com.br', 'Nome da Pessoa', 'user')
on conflict (email) do update
  set nome = excluded.nome, role = excluded.role, ativo = true;

-- 2) quem vê o quê. Admin NÃO precisa de linha aqui — vê tudo.
insert into user_dashboard_access (email, dashboard) values
  ('fulano@luxor.com.br', 'semanal'),
  ('fulano@luxor.com.br', 'comite')
on conflict do nothing;

-- Revogar acesso sem apagar histórico de auditoria:
--   update allowed_users set ativo = false where email = 'saiu@luxor.com.br';
