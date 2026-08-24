/* Configuração PÚBLICA do Hub HPG.

   A anon key é pública por design — quem protege é a RLS + a allowlist do
   Supabase (sql/hub_schema.sql). NUNCA colocar a service_role / secret aqui:
   ela ignora RLS e daria acesso total ao banco. Ela vive só no .env local, que
   é gitignored, e é usada pelo tools/publish_hub.py.

   O projeto Supabase do HPG é PRÓPRIO — outra allowlist, outras pessoas, outro
   bucket. Nada compartilhado com o hub do P&C: erro de policy num hub não deve
   alcançar o outro.

   Enquanto a ANON KEY estiver vazia, o hub roda em DEMO OFFLINE: abre o
   index.html por file://, sem login e sem rede, com os assets que os
   tools/build_*.py geraram na máquina. */

window.SUPABASE_URL = "https://opiihihhfucmzbnpkfmu.supabase.co";
window.SUPABASE_ANON_KEY = "";   // Supabase > Settings > API > "anon public"
                                 // NUNCA a service_role: ela ignora RLS, e este
                                 // arquivo vai pro repo publico.

// Bucket PRIVADO com o que o pipeline publica (tools/publish_hub.py).
window.HUB_BUCKET = "hpg-data";
