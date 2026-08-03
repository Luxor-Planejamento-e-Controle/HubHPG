/* Configuração do Hub HPG.
   Fase atual: DEMO OFFLINE — abre o index.html direto (file://), sem login e
   sem rede. O painel da única aba vem de tools/build_semanal.py.

   Quando for a gold, é aqui que entram SUPABASE_URL + anon key (pública por
   design) e o hub passa a montar só depois do porteiro (assets/auth.js),
   igual ao LuxorP&CHub. A allowlist do HPG é PRÓPRIA — outro projeto Supabase,
   outras pessoas; nada compartilhado com o hub do P&C. */
window.HUB = {
  offline: true,
  dashboards: ['semanal'],   // offline: tudo liberado
};
