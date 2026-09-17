/* Compara a apuração do motor com o MAPA da Controladoria, animal por animal.

   O mapa (`Plantel Haras Pao Grande - Movimentação Jan a Dez 2026 (<Mês>).xlsx`,
   aba `Movimentações`) é a classificação curada: uma linha por animal com o
   rateio ACUMULADO no ano entre COMPRAS, EMBRIÕES, BAIXAS POR VENDA, BAIXAS POR
   MORTE/DOAÇÃO e REAVALIAÇÃO. Os valores já estão na fatia da Carla, então batem
   direto com `delta_carla`.

   Comparar totais de resumo só diz QUANTO diverge; isto diz QUEM. Sem isso,
   mexer em regra de classificação é chute.

   Uso: node tools/confere_mapa.js 2026-06 [--todos]
*/
'use strict';
const fs = require('fs');
const path = require('path');
const {execFileSync} = require('child_process');

const RAIZ = path.resolve(__dirname, '..');
const XLSX = require(path.join(RAIZ, 'assets/vendor/xlsx.mini.min.js'));

/* ---- DOM de mentira para carregar o motor ---- */
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

/* ---- arquivos do haras (mesma escolha que a aba faz) ---- */
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

/* ---- mapa da Controladoria ---- */
const MAPA_DIR = 'G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais/'
  + 'RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG/Posição Equinos/PLANTEL - Movimentações/2026';
const ABR = {'01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr', '05': 'Mai', '06': 'Jun',
             '07': 'Jul', '08': 'Ago'};
// coluna do mapa -> classe do motor
const COLUNA_CLASSE = {
  'COMPRAS': 'compra',
  'EMBRIOES CONFIRMADOS ACIMA DE 60 DIAS': 'embriao',
  'BAIXAS POR VENDA': 'venda',
  'BAIXAS POR MORTE/DOACAO': 'morte_doacao',
  'REAVALIACAO': 'reavaliacao',
};
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();

function leMapa(mes){
  const arq = fs.readdirSync(MAPA_DIR)
    .find(f => f.includes(`(${ABR[mes.slice(5)]} ${mes.slice(0, 4)})`) && !f.startsWith('~$'));
  if (!arq) throw new Error('mapa não encontrado para ' + mes);
  const wb = XLSX.read(fs.readFileSync(path.join(MAPA_DIR, arq)), {cellDates: true});
  const nome = wb.SheetNames.find(n => norm(n).startsWith('MOVIMENTA'));
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[nome], {header: 1, defval: null, raw: true});
  const iCab = linhas.findIndex(r => r && r.some(c => norm(c) === 'NOME'));
  const cab = linhas[iCab].map(norm);
  const iNome = cab.indexOf('NOME');
  const cols = {};
  for (const [rot, cls] of Object.entries(COLUNA_CLASSE)) {
    const j = cab.findIndex(c => c === rot);
    if (j >= 0) cols[cls] = j;
  }
  const out = {};
  for (const r of linhas.slice(iCab + 1)) {
    if (!r || !r[iNome]) continue;
    const k = norm(r[iNome]);
    out[k] = out[k] || {};
    for (const [cls, j] of Object.entries(cols)) {
      const v = Number(r[j]);
      if (isFinite(v) && v) out[k][cls] = (out[k][cls] || 0) + v;
    }
  }
  return {arquivo: arq, animais: out};
}

/* ---- apuração do motor, acumulada de jan até o mês ---- */
function apuraAte(mes){
  const meses = [];
  for (let m = 1; m <= Number(mes.slice(5)); m++) meses.push(`${mes.slice(0, 4)}-${String(m).padStart(2, '0')}`);
  const acc = {};
  for (const m of meses) {
    carrega(m); carrega(M.mesAnterior(m));
    M.ST.mes = m;
    for (const mo of M.movimentacaoDoMes(m).movs) {
      if (!mo.delta_carla) continue;
      const cls = mo.sugestao === 'morte' || mo.sugestao === 'doacao' ? 'morte_doacao' : mo.sugestao;
      const k = norm(mo.nome);
      acc[k] = acc[k] || {};
      acc[k][cls || '(sem sugestao)'] = (acc[k][cls || '(sem sugestao)'] || 0) + mo.delta_carla;
    }
  }
  return acc;
}

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');
const CLASSES = ['compra', 'embriao', 'venda', 'morte_doacao', 'reavaliacao'];

const mes = process.argv[2] || '2026-06';
const todos = process.argv.includes('--todos');
const mapa = leMapa(mes);
const nosso = apuraAte(mes);
console.log(`mapa: ${mapa.arquivo}`);
console.log(`animais no mapa com movimento: ${Object.values(mapa.animais).filter(v => Object.keys(v).length).length}`);

const nomes = [...new Set([...Object.keys(mapa.animais), ...Object.keys(nosso)])].sort();
let iguais = 0;
const divergentes = [];
for (const n of nomes) {
  const a = mapa.animais[n] || {}, b = nosso[n] || {};
  const difs = CLASSES.concat(['(sem sugestao)'])
    .map(c => [c, (b[c] || 0) - (a[c] || 0)])
    .filter(([, d]) => Math.abs(d) >= 1);
  if (!difs.length) { iguais++; continue; }
  divergentes.push([n, difs, a, b]);
}
console.log(`\n${iguais} animais batendo · ${divergentes.length} divergentes\n`);
divergentes.sort((x, y) => Math.max(...y[1].map(d => Math.abs(d[1])))
                          - Math.max(...x[1].map(d => Math.abs(d[1]))));
for (const [n, difs, a, b] of divergentes.slice(0, todos ? 999 : 25)) {
  console.log(`${n.slice(0, 44)}`);
  console.log(`   mapa : ${CLASSES.filter(c => a[c]).map(c => c + ' ' + rs(a[c])).join(' · ') || '—'}`);
  console.log(`   nosso: ${Object.keys(b).filter(c => b[c]).map(c => c + ' ' + rs(b[c])).join(' · ') || '—'}`);
}
