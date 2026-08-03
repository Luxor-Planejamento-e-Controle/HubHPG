/* Configuração do Hub HPG.
   Fase atual: DEMO OFFLINE — abre o index.html direto (file://), sem login e
   sem rede. Os dados vêm dos assets/data/*.js gerados por tools/build_data.py.

   Quando for a gold, é aqui que entram SUPABASE_URL + anon key (pública por
   design) e o hub passa a montar só depois do porteiro (assets/auth.js),
   igual ao LuxorP&CHub. A allowlist do HPG é PRÓPRIA — outro projeto Supabase,
   outras pessoas; nada compartilhado com o hub do P&C. */
window.HUB = {
  offline: true,
  dashboards: ['semanal', 'comite', 'plantel'],   // offline: tudo liberado
};
