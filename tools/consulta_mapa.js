/* Mostra as linhas da aba `Movimentações` do mapa da Controladoria para os
   animais cujo nome casa com os termos passados. É a evidência usada para
   decidir em que classe um evento entra quando o log do haras é ambíguo.

   Uso: node tools/consulta_mapa.js MORENA OASIS "PLATAO"
*/
'use strict';
const fs = require('fs');
const path = require('path');
const RAIZ = path.resolve(__dirname, '..');
const XLSX = require(path.join(RAIZ, 'assets/vendor/xlsx.mini.min.js'));

const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();

const MAPA_DIR = 'G:/Drives compartilhados/Luxor Controladoria/Relatórios Gerenciais/'
  + 'RELATORIOS - OPERAÇÃO HARAS E FAZENDA PG/Posição Equinos/PLANTEL - Movimentações/2026';
const arq = fs.readdirSync(MAPA_DIR)
  .filter(f => f.startsWith('Plantel Haras Pao Grande') && !f.startsWith('~$'))
  .sort((a, b) => fs.statSync(path.join(MAPA_DIR, b)).mtimeMs
                - fs.statSync(path.join(MAPA_DIR, a)).mtimeMs)[0];
console.log('mapa:', arq);
const wb = XLSX.read(fs.readFileSync(path.join(MAPA_DIR, arq)), {cellDates: true});
const nomeAba = wb.SheetNames.find(n => norm(n).startsWith('MOVIMENTA'));
const linhas = XLSX.utils.sheet_to_json(wb.Sheets[nomeAba], {header: 1, defval: null, raw: true});
const iCab = linhas.findIndex(r => r && r.some(c => norm(c) === 'NOME'));
const cab = linhas[iCab].map(norm);
const iNome = cab.indexOf('NOME');

const termos = process.argv.slice(2).map(norm);
const fmt = v => v instanceof Date ? v.toISOString().slice(0, 10)
  : (typeof v === 'number' ? v.toLocaleString('pt-BR') : String(v));

for (const r of linhas.slice(iCab + 1)) {
  if (!r || !r[iNome]) continue;
  const n = norm(r[iNome]);
  if (!termos.some(t => n.includes(t))) continue;
  console.log(`\n  ${r[iNome]}`);
  r.forEach((v, j) => {
    if (j === iNome || v == null || v === '' || v === 0) return;
    console.log(`     ${(cab[j] || `col${j}`).slice(0, 44).padEnd(46)} ${fmt(v)}`);
  });
}

/* --coluna "BAIXAS POR MORTE/DOACAO": lista toda linha com valor nessa coluna */
const iCol = process.argv.indexOf('--coluna');
if (iCol > 0) {
  const alvo = norm(process.argv[iCol + 1]);
  const j = cab.findIndex(c => c === alvo);
  console.log(`\n--- linhas com ${alvo} (coluna ${j}) ---`);
  let soma = 0;
  for (const r of linhas.slice(iCab + 1)) {
    if (!r || !r[iNome] || /TOTA/.test(norm(r[iNome]))) continue;
    const v = Number(r[j]);
    if (!isFinite(v) || !v) continue;
    soma += v;
    console.log(`  ${fmt(v).padStart(14)}  ${String(r[iNome]).slice(0, 62)}`);
  }
  console.log(`  soma: ${soma.toLocaleString('pt-BR')}`);
}
