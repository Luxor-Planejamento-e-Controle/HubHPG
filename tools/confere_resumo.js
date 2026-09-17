/* Confere a apuração do motor contra o Resumo Contábil DIVULGADO, mês a mês.

   O divulgado é lido do mapa da Controladoria (aba `Resumo Contabil` do
   `Plantel Haras Pao Grande - Movimentação Jan a Dez 2026 (<Mês>).xlsx`), não de
   tabela digitada aqui: número copiado à mão envelhece e mente.

   Tudo num processo só, carregando cada arquivo do haras UMA vez. A versão
   anterior disparava um node por mês e cada um relia os xlsx do Drive — 14
   leituras de um mount lento para 7 meses, o que tornava cada iteração de regra
   insuportável.

   Uso: node tools/confere_resumo.js [--animais]
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

const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();

/* ---- arquivos do haras ---- */
const MESES = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07'];
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

/* ---- divulgado, direto do mapa ---- */
const MAPA_DIR = 'G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais/'
  + 'RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG/Posição Equinos/PLANTEL - Movimentações/2026';
const ABR = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago'];
const LINHA_CLASSE = [
  ['(+) COMPRAS', 'compra'],
  ['(+) PRODUCAO EMBRIOES', 'embriao'],
  ['(-) BAIXA VENDAS', 'venda'],
  ['(-) BAIXA MORTES E DOACOES', 'morte_doacao'],
  ['(+/-) REAVALIACOES', 'reavaliacao'],
];

function leDivulgado(){
  // o mapa mais recente já traz a coluna de todos os meses fechados
  const arq = fs.readdirSync(MAPA_DIR)
    .filter(f => f.startsWith('Plantel Haras Pao Grande') && !f.startsWith('~$'))
    .sort((a, b) => fs.statSync(path.join(MAPA_DIR, b)).mtimeMs
                  - fs.statSync(path.join(MAPA_DIR, a)).mtimeMs)[0];
  const wb = XLSX.read(fs.readFileSync(path.join(MAPA_DIR, arq)), {cellDates: true});
  const nome = wb.SheetNames.find(n => norm(n).startsWith('RESUMO'));
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[nome], {header: 1, defval: null, raw: true});
  const iCab = linhas.findIndex(r => r && r.some(c => norm(c) === 'TITULO'));
  const cab = linhas[iCab].map(norm);
  const colMes = {};
  cab.forEach((c, j) => {
    const k = ABR.findIndex(a => norm(a) === c);
    if (k >= 0) colMes[`2026-${String(k + 1).padStart(2, '0')}`] = j;
  });
  const out = {};
  for (const r of linhas.slice(iCab + 1)) {
    if (!r) continue;
    const rot = norm(r.find(c => typeof c === 'string' && c.trim()));
    const par = LINHA_CLASSE.find(([l]) => rot.startsWith(l));
    if (!par) continue;
    for (const [mes, j] of Object.entries(colMes)) {
      const v = Number(r[j]);
      (out[mes] = out[mes] || {})[par[1]] = isFinite(v) ? v : 0;
    }
  }
  return {arquivo: arq, meses: out};
}

/* ---- apuração ---- */
function causas(mes){
  carrega(mes); carrega(M.mesAnterior(mes));
  M.ST.mes = mes;
  const out = {}, semClasse = [];
  for (const mo of M.movimentacaoDoMes(mes).movs) {
    if (!mo.delta_carla) continue;
    if (!mo.sugestao) { semClasse.push(mo); continue; }
    const c = (mo.sugestao === 'morte' || mo.sugestao === 'doacao') ? 'morte_doacao' : mo.sugestao;
    out[c] = (out[c] || 0) + mo.delta_carla;
  }
  return {out, semClasse};
}

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');
const CLASSES = ['compra', 'embriao', 'venda', 'morte_doacao', 'reavaliacao'];

const div = leDivulgado();
console.log('divulgado lido de:', div.arquivo);
const detalhe = process.argv.includes('--animais');
let ok = 0, tot = 0;
for (const mes of MESES) {
  const d = div.meses[mes];
  if (!d) continue;
  const {out, semClasse} = causas(mes);
  console.log(`\n${mes}`);
  for (const c of CLASSES) {
    const a = out[c] || 0, b = d[c] || 0, dif = a - b;
    tot++;
    const bate = Math.abs(dif) < 1;
    if (bate) ok++;
    console.log(`    ${c.padEnd(13)} apurado ${rs(a).padStart(15)}   divulgado ${rs(b).padStart(15)}`
      + `   dif ${rs(dif).padStart(13)}  ${bate ? 'ok' : (b === 0 && a !== 0 ? 'INVENTADO' : 'DIVERGE')}`);
  }
  if (semClasse.length) {
    const s = semClasse.reduce((x, m) => x + m.delta_carla, 0);
    console.log(`    ${'(sem sugestao)'.padEnd(13)} ${rs(s).padStart(23)}   ${semClasse.length} animal(is)`);
    if (detalhe) for (const m of semClasse) {
      console.log(`        ${rs(m.delta_carla).padStart(13)}  ${m.nome.slice(0, 44)}`);
      for (const l of (m.log || [])) console.log(`            ${String(l.ocorrencia).slice(0, 96)}`);
    }
  }
}
console.log(`\n${ok} de ${tot} linhas batendo.`);
