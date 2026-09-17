/* Procura, no arquivo do haras de um mês, as LINHAS e as OCORRÊNCIAS cujo nome
   casa com o termo. Serve para ver o dado cru quando o motor e o mapa discordam
   e o resumo por animal não explica.

   Uso: node tools/busca_haras.js 2026-03 "FADA DA PAO GRANDE X PALHACO"
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
const M = new Function(fonte + '\n;return {ST, lerArquivo};')();

const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();

const mes = process.argv[2];
const termos = process.argv.slice(3).map(norm);
const [ano, mm] = mes.split('-');
const p = execFileSync('python', ['-c',
  `import sys; sys.path.insert(0, r'${path.join(RAIZ, 'scripts')}')\n`
  + `from _pg_common import parse_mes_aaaa, find_source_file\n`
  + `print(find_source_file(parse_mes_aaaa('${mm}/${ano}')))`],
  {encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'],
   env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'})}).trim();
console.log('arquivo:', path.basename(p));
const d = M.lerArquivo(fs.readFileSync(p), path.basename(p));
const ix = d.ix;

console.log('\n--- linhas ---');
for (const l of d.linhas) {
  const n = norm(l[ix.nome]);
  if (!termos.some(t => n.includes(t))) continue;
  console.log(`  ${l[ix.nome]}`);
  for (const [k, j] of Object.entries(ix)) {
    if (j == null || l[j] == null || l[j] === '') continue;
    console.log(`     ${k.padEnd(12)} ${String(l[j] instanceof Date ? l[j].toISOString().slice(0, 10) : l[j]).slice(0, 90)}`);
  }
}
console.log('\n--- ocorrências ---');
for (const x of d.log || []) {
  const n = norm(x.produto);
  if (!termos.some(t => n.includes(t))) continue;
  const dt = x.data instanceof Date ? x.data : new Date(x.data);
  console.log(`  ${isNaN(dt) ? '??' : dt.toISOString().slice(0, 10)}  ${x.produto}`);
  console.log(`        ${x.ocorrencia}`);
}
