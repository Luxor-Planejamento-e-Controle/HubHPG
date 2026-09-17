/* Despeja, para um mês, TODA movimentação que o motor apurou, agrupada pela
   classe que ele sugeriu e ordenada por valor. É o instrumento para casar a
   diferença do `confere_resumo.js` (ex.: "venda -R$ 22.250, morte_doacao
   +R$ 22.250") com o animal que a causou.

   Uso: node tools/dump_mes.js 2026-03 [2026-07 ...]
*/
'use strict';
const fs = require('fs');
const path = require('path');
const {execFileSync} = require('child_process');

const RAIZ = path.resolve(__dirname, '..');
const XLSX = require(path.join(RAIZ, 'assets/vendor/xlsx.mini.min.js'));

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
global.alert = () => {};
global.XLSX = XLSX;

const fonte = fs.readFileSync(path.join(RAIZ, 'assets/plantel/plantel.js'), 'utf8');
const M = new Function(fonte + '\n;return {ST, lerArquivo, movimentacaoDoMes, mesAnterior};')();

function arquivoDoMes(mes){
  const [ano, mm] = mes.split('-');
  return execFileSync('python', ['-c',
    `import sys; sys.path.insert(0, r'${path.join(RAIZ, 'scripts')}')\n`
    + `from _pg_common import parse_mes_aaaa, find_source_file\n`
    + `print(find_source_file(parse_mes_aaaa('${mm}/${ano}')))`],
    {encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'],
     env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'})}).trim();
}
function carrega(mes){
  if (M.ST.meses[mes]) return;
  const p = arquivoDoMes(mes);
  M.ST.meses[mes] = M.lerArquivo(fs.readFileSync(p), path.basename(p));
}

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');

for (const mes of process.argv.slice(2).filter(a => /^\d{4}-\d{2}$/.test(a))) {
  carrega(mes); carrega(M.mesAnterior(mes));
  M.ST.mes = mes;
  const porClasse = {};
  for (const mo of M.movimentacaoDoMes(mes).movs) {
    if (!mo.delta_carla) continue;
    const c = mo.sugestao || '(sem sugestao)';
    (porClasse[c] = porClasse[c] || []).push(mo);
  }
  console.log(`\n================ ${mes} ================`);
  for (const c of Object.keys(porClasse).sort()) {
    const lista = porClasse[c].sort((a, b) => Math.abs(b.delta_carla) - Math.abs(a.delta_carla));
    const tot = lista.reduce((s, m) => s + m.delta_carla, 0);
    console.log(`\n  [${c}] ${rs(tot)} — ${lista.length} linha(s)`);
    for (const mo of lista) {
      console.log(`    ${rs(mo.delta_carla).padStart(14)}  ${String(mo.nome).slice(0, 52)}`
        + ` | cota ${mo.cota_ant}->${mo.cota_atual} | valor ${mo.valor_ant}->${mo.valor_atual}`
        + `${mo.entrou ? ' | ENTROU' : ''}${mo.saiu ? ' | SAIU' : ''}`);
      for (const l of (mo.log || [])) {
        console.log(`          log: ${String(l.ocorrencia).slice(0, 118)}`);
      }
    }
  }
}
