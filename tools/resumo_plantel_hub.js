/* Resumo contábil do plantel COMO A ABA PLANTEL DO HUB O APURA — para o slide
   "Resumo da movimentação do plantel" do comitê.

   O comitê lia o resumo da planilha da Controladoria (aba `Resumo Contabil` do
   mapa). O fechamento agora é feito no hub: cada mês importado vira um arquivo
   `plantel.<AAAA-MM>.json` no bucket, e quem fecha classifica as movimentações
   (tabela `plantel_mov_classificacao`, com os lançamentos manuais). Este script
   roda o MESMO motor da aba (assets/plantel/plantel.js), com um DOM de mentira,
   sobre esses dados — então o número do slide é o número da tela.

   Lê do Supabase com a service role do .env (a mesma do publish_hub.py); só lê.

   Uso: node tools/resumo_plantel_hub.js 2026-08
   Saída (stdout): {"meses": {"2026-01": {ini, fim, causas, classificado, total,
                   fechado, liberado}, ...}} — YTD até o mês pedido.
*/
'use strict';
const fs = require('fs');
const path = require('path');

const RAIZ = path.resolve(__dirname, '..');
const ALVO = process.argv[2];
if (!/^\d{4}-\d{2}$/.test(ALVO || '')) {
  console.error('uso: node tools/resumo_plantel_hub.js <AAAA-MM>');
  process.exit(2);
}

/* ---- credenciais: .env do repo, nunca no Git ---- */
function env(){
  const out = {};
  const p = path.join(RAIZ, '.env');
  if (!fs.existsSync(p)) return out;
  for (const l of fs.readFileSync(p, 'utf8').split(/\r?\n/)) {
    const m = /^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/.exec(l);
    if (m) out[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
  return Object.assign(out, process.env);
}
const E = env();
const URL = (E.SUPABASE_URL || '').replace(/\/$/, '');
const KEY = E.SUPABASE_SERVICE_ROLE_KEY;
if (!URL || !KEY) {
  console.error('sem SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env');
  process.exit(3);
}
const H = {apikey: KEY, Authorization: `Bearer ${KEY}`};

/* ---- DOM de mentira: o suficiente pro motor carregar sem tela ---- */
const elemento = () => new Proxy({}, {
  get: (o, k) => k in o ? o[k]
    : (k === 'classList' ? {add(){}, remove(){}, toggle(){}, contains: () => false}
      : k === 'style' || k === 'dataset' ? {}
      : k === 'querySelectorAll' ? (() => [])
      : k === 'querySelector' ? (() => null)
      : ['addEventListener', 'appendChild', 'focus', 'setSelectionRange'].includes(k) ? (() => {})
      : undefined),
  set: (o, k, v) => { o[k] = v; return true; },
});
global.document = {getElementById: () => elemento(), querySelector: () => null,
                   querySelectorAll: () => [], body: {addEventListener: () => {}},
                   createElement: () => elemento()};
global.window = {parent: {}, innerWidth: 1280, innerHeight: 800, location: {hash: ''}};
global.alert = msg => console.error('[alert]', msg);
global.XLSX = require(path.join(RAIZ, 'assets/vendor/xlsx.mini.min.js'));

const fonte = fs.readFileSync(path.join(RAIZ, 'assets/plantel/plantel.js'), 'utf8');
// o boot() sai sozinho sem Supabase (sb() devolve null)
const M = new Function(fonte + '\n;return {ST, resumoAno, normalizaSnapshot, mesAnterior};')();

async function json(url){
  const r = await fetch(url, {headers: H});
  if (!r.ok) throw new Error(`${r.status} em ${url.replace(URL, '')}`);
  return r.json();
}

(async () => {
  // meses do ano até o alvo, mais dezembro anterior (o saldo inicial de janeiro)
  const ano = +ALVO.slice(0, 4);
  const meses = [M.mesAnterior(`${ano}-01`)];
  for (let m = 1; m <= +ALVO.slice(5); m++) meses.push(`${ano}-${String(m).padStart(2, '0')}`);

  for (const m of meses) {
    let r = null;
    try { r = await json(`${URL}/storage/v1/object/hpg-data/plantel.${m}.json`); }
    catch (e) {
      // mês sem arquivo no bucket: a tabela é o registro durável
      const t = await json(`${URL}/rest/v1/plantel_snapshot?select=mes,arquivo,linhas,log&mes=eq.${m}`);
      r = t[0] || null;
    }
    if (r) M.ST.meses[m] = M.normalizaSnapshot(r);
  }
  for (const r of await json(`${URL}/rest/v1/plantel_mov_classificacao?select=mes,chave,classe,nota,autor,nome,valor`))
    M.ST.decisoes[`${r.mes}|${r.chave}`] = r;
  for (const r of await json(`${URL}/rest/v1/plantel_mes_status?select=mes,fechado`))
    M.ST.statusMes[r.mes] = !!r.fechado;

  M.ST.mes = ALVO;
  const res = M.resumoAno();
  const out = {};
  for (const [m, v] of Object.entries(res))
    out[m] = Object.assign({}, v, {arquivo: M.ST.meses[m].arquivo, liberado: M.ST.meses[m].liberado || null});
  process.stdout.write(JSON.stringify({meses: out}));
})().catch(e => { console.error(String(e && e.stack || e)); process.exit(1); });
