/* Para cada animal em que o motor discorda do mapa, mostra TUDO que o motor
   tinha em mãos: status, cota, valor e o log do mês — ao lado da classe que a
   Controladoria deu.

   O mapa é acumulado no ano, então a classificação DAQUELE mês sai da diferença
   entre o mapa do mês e o do mês anterior. É isso que permite comparar mês a
   mês em vez de só no acumulado.

   Serve para derivar regra a partir de evidência: em vez de perguntar "o que
   aconteceu com esse animal", a resposta já está no mapa; o que falta é achar o
   que no arquivo do haras prediz aquela resposta.

   Uso: node tools/diagnostica_motor.js [2026-02] [--todos]
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

const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();
function chaveNome(s){
  return norm(s)
    .replace(/^(MACHO|FEMEA|POTRA|POTRO|EMBRIAO)\s+/, '')
    .replace(/\s*-?\s*\d{2}\/\d{2}\/\d{2,4}.*$/, '')
    .replace(/\s+RECEP\s+\d+.*$/, '')
    .replace(/[()]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 28);
}

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

/* ---- mapas ---- */
const MAPA_DIR = 'G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais/'
  + 'RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG/Posição Equinos/PLANTEL - Movimentações/2026';
const ABR = {'01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr', '05': 'Mai', '06': 'Jun',
             '07': 'Jul', '08': 'Ago'};
const COLUNA_CLASSE = {
  'COMPRAS': 'compra',
  'EMBRIOES CONFIRMADOS ACIMA DE 60 DIAS': 'embriao',
  'BAIXAS POR VENDA': 'venda',
  'BAIXAS POR MORTE/DOACAO': 'morte_doacao',
  'REAVALIACAO': 'reavaliacao',
};
const cacheMapa = {};
function leMapa(mes){
  if (cacheMapa[mes]) return cacheMapa[mes];
  const arq = fs.readdirSync(MAPA_DIR)
    .find(f => f.includes(`(${ABR[mes.slice(5)]} ${mes.slice(0, 4)})`) && !f.startsWith('~$'));
  if (!arq) return (cacheMapa[mes] = {});
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
    if (!r || !r[iNome] || /TOTA/.test(norm(r[iNome]))) continue;
    const k = chaveNome(r[iNome]);
    out[k] = out[k] || {};
    for (const [cls, j] of Object.entries(cols)) {
      const v = Number(r[j]);
      if (isFinite(v) && v) out[k][cls] = (out[k][cls] || 0) + v;
    }
  }
  return (cacheMapa[mes] = out);
}

/* classificação DAQUELE mês = acumulado do mês menos o do mês anterior */
function mapaDoMes(mes){
  const hoje = leMapa(mes);
  const ant = mes === '2026-01' ? {} : leMapa(M.mesAnterior(mes));
  const out = {};
  for (const k of new Set([...Object.keys(hoje), ...Object.keys(ant)])) {
    const a = ant[k] || {}, b = hoje[k] || {};
    for (const c of Object.values(COLUNA_CLASSE)) {
      const d = (b[c] || 0) - (a[c] || 0);
      if (Math.abs(d) >= 1) (out[k] = out[k] || {})[c] = d;
    }
  }
  return out;
}

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');
const fmtCls = o => Object.entries(o).map(([c, v]) => `${c} ${rs(v)}`).join(' · ') || '—';

const MESES = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07'];
const alvoMes = process.argv.find(a => /^\d{4}-\d{2}$/.test(a));
const todos = process.argv.includes('--todos');

for (const mes of (alvoMes ? [alvoMes] : MESES)) {
  carrega(mes); carrega(M.mesAnterior(mes));
  M.ST.mes = mes;
  const nossos = {};
  for (const mo of M.movimentacaoDoMes(mes).movs) {
    if (!mo.delta_carla) continue;
    const c = (mo.sugestao === 'morte' || mo.sugestao === 'doacao') ? 'morte_doacao'
            : (mo.sugestao || '(sem sugestao)');
    const k = chaveNome(mo.nome);
    nossos[k] = nossos[k] || {classes: {}, mo};
    nossos[k].classes[c] = (nossos[k].classes[c] || 0) + mo.delta_carla;
  }
  const doMapa = mapaDoMes(mes);
  const chaves = [...new Set([...Object.keys(doMapa), ...Object.keys(nossos)])];
  const difs = [];
  for (const k of chaves) {
    const a = doMapa[k] || {}, b = (nossos[k] || {}).classes || {};
    const classes = new Set([...Object.keys(a), ...Object.keys(b)]);
    let igual = true;
    for (const c of classes) if (Math.abs((a[c] || 0) - (b[c] || 0)) >= 1) igual = false;
    if (!igual) difs.push([k, a, b, (nossos[k] || {}).mo]);
  }
  console.log(`\n${'='.repeat(78)}\n${mes} — ${chaves.length - difs.length} batendo · ${difs.length} divergentes`);
  difs.sort((x, y) => Math.max(...Object.values(y[1]).map(Math.abs), 0)
                    - Math.max(...Object.values(x[1]).map(Math.abs), 0));
  for (const [k, a, b, mo] of difs.slice(0, todos ? 999 : 8)) {
    console.log(`\n  ${k}`);
    console.log(`     mapa : ${fmtCls(a)}`);
    console.log(`     motor: ${fmtCls(b)}`);
    if (mo) {
      console.log(`     dados: status ${mo.mudou_status ? mo.mudou_status.join(' -> ') : (mo.status || '')}`
        + ` | cota ${mo.cota_ant} -> ${mo.cota_atual} | valor ${mo.valor_ant} -> ${mo.valor_atual}`
        + `${mo.entrou ? ' | ENTROU' : ''}${mo.saiu ? ' | SAIU' : ''}`);
      for (const l of (mo.log || [])) console.log(`     log  : ${String(l.ocorrencia).slice(0, 118)}`);
    } else {
      console.log('     dados: (o motor não viu movimento neste animal)');
    }
  }
}
