/* Para cada animal em que o motor discorda do mapa, mostra TUDO que o motor
   tinha em mãos — status, cota, valor e o log — ao lado da classe que a
   Controladoria deu. Serve para derivar regra a partir de evidência.

   ATENÇÃO ao que o mapa é e ao que ele não é:

   A aba `Movimentações` é uma planilha VIVA, não um fechamento mensal. O mapa
   de janeiro já traz PATRONO, MADONA e RAIZ com baixa por venda — vendas
   datadas de FEVEREIRO no log do haras — e os arquivos de Jan, Fev e Mar têm
   valores idênticos para esses animais. Cada arquivo mensal é um retrato da
   mesma planilha; diferenciar um contra o outro para achar "o mês" dá zero.

   A atribuição POR MÊS mora só na aba `Resumo Contabil` (colunas Jan..Ago), e é
   o que tools/confere_resumo.js compara. Aqui a comparação é do ACUMULADO do
   período: que classe cada animal recebeu no total, contra a que o motor deu
   somando os meses. É assim que se acha regra errada sem depender de o mapa
   saber dizer quando.

   Uso: node tools/diagnostica_motor.js [2026-07] [--todos]
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

const MAPA_DIR = 'G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais/'
  + 'RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG/Posição Equinos/PLANTEL - Movimentações/2026';
const COLUNA_CLASSE = {
  'COMPRAS': 'compra',
  'EMBRIOES CONFIRMADOS ACIMA DE 60 DIAS': 'embriao',
  'BAIXAS POR VENDA': 'venda',
  'BAIXAS POR MORTE/DOACAO': 'morte_doacao',
  'REAVALIACAO': 'reavaliacao',
};

function mapaAcumulado(){
  // o mapa mais novo é o retrato mais completo da planilha viva
  const arq = fs.readdirSync(MAPA_DIR)
    .filter(f => f.startsWith('Plantel Haras Pao Grande') && !f.startsWith('~$'))
    .sort((a, b) => fs.statSync(path.join(MAPA_DIR, b)).mtimeMs
                  - fs.statSync(path.join(MAPA_DIR, a)).mtimeMs)[0];
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
  return {arq, animais: out};
}

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');
const fmtCls = o => Object.entries(o).filter(([, v]) => Math.abs(v) >= 1)
  .map(([c, v]) => `${c} ${rs(v)}`).join(' · ') || '—';

const ate = process.argv.find(a => /^\d{4}-\d{2}$/.test(a)) || '2026-07';
const todos = process.argv.includes('--todos');

/* nosso acumulado de jan até `ate`, por animal, guardando o contexto do mês em
   que cada pedaço aconteceu */
const nossos = {};
for (let n = 1; n <= Number(ate.slice(5)); n++) {
  const mes = `${ate.slice(0, 4)}-${String(n).padStart(2, '0')}`;
  carrega(mes); carrega(M.mesAnterior(mes));
  M.ST.mes = mes;
  for (const mo of M.movimentacaoDoMes(mes).movs) {
    if (!mo.delta_carla) continue;
    const c = (mo.sugestao === 'morte' || mo.sugestao === 'doacao') ? 'morte_doacao'
            : (mo.sugestao || '(sem sugestao)');
    const k = chaveNome(mo.nome);
    nossos[k] = nossos[k] || {classes: {}, ctx: []};
    nossos[k].classes[c] = (nossos[k].classes[c] || 0) + mo.delta_carla;
    nossos[k].ctx.push({mes, mo});
  }
}

const mapa = mapaAcumulado();
console.log('mapa:', mapa.arq);
const chaves = [...new Set([...Object.keys(mapa.animais), ...Object.keys(nossos)])];
const difs = [];
for (const k of chaves) {
  const a = mapa.animais[k] || {}, b = (nossos[k] || {}).classes || {};
  const classes = new Set([...Object.keys(a), ...Object.keys(b)]);
  let igual = true;
  for (const c of classes) if (Math.abs((a[c] || 0) - (b[c] || 0)) >= 1) igual = false;
  if (!igual) difs.push([k, a, b, (nossos[k] || {}).ctx || []]);
}
console.log(`\njan..${ate} — ${chaves.length - difs.length} animais batendo · ${difs.length} divergentes\n`);
difs.sort((x, y) => Math.max(...Object.values(y[1]).map(Math.abs), 0)
                  - Math.max(...Object.values(x[1]).map(Math.abs), 0));
for (const [k, a, b, ctx] of difs.slice(0, todos ? 999 : 12)) {
  console.log(`\n  ${k}`);
  console.log(`     mapa : ${fmtCls(a)}`);
  console.log(`     motor: ${fmtCls(b)}`);
  for (const {mes, mo} of ctx) {
    console.log(`     ${mes}: ${mo.sugestao || '(sem sugestao)'} ${rs(mo.delta_carla)}`
      + ` | status ${mo.mudou_status ? mo.mudou_status.join('->') : (mo.status || '')}`
      + ` | cota ${mo.cota_ant}->${mo.cota_atual} | valor ${mo.valor_ant}->${mo.valor_atual}`
      + `${mo.entrou ? ' | ENTROU' : ''}${mo.saiu ? ' | SAIU' : ''}`);
    for (const l of (mo.log || [])) console.log(`          log: ${String(l.ocorrencia).slice(0, 112)}`);
  }
  if (!ctx.length) console.log('     (o motor não viu movimento nenhum neste animal)');
}
