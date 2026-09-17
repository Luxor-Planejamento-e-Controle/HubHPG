/* Compara, mês a mês, as causas que o motor apura com o Resumo Contábil que foi
   DIVULGADO (o número que valeu, digitado à mão na planilha do haras).

   Serve para achar em qual causa e em qual mês a apuração diverge, que é o que
   nenhuma comparação de saldo mostra: o saldo pode fechar com a repartição
   errada. Os valores divulgados abaixo vieram da planilha original, conferidos
   com o Arthur em 17/09/2026.

   Uso: node tools/confere_resumo.js [--animais]
*/
'use strict';
const {execFileSync} = require('child_process');
const path = require('path');

// Resumo Contábil divulgado, jan a ago/2026 (escopo Carla)
const DIVULGADO = {
  '2026-01': {compra: 3000, embriao: 83500, venda: -176000, morte_doacao: -150000, reavaliacao: -18000},
  '2026-02': {compra: 7500, embriao: 246000, venda: -284696.56, morte_doacao: -120000, reavaliacao: 0},
  '2026-03': {compra: 0, embriao: 142500, venda: -1090811.98, morte_doacao: -283250, reavaliacao: 0},
  '2026-04': {compra: 64300, embriao: 0, venda: -133154.94, morte_doacao: 0, reavaliacao: -27500},
  '2026-05': {compra: 20000, embriao: 242500, venda: -282340, morte_doacao: -28773.50, reavaliacao: 0},
  '2026-06': {compra: 0, embriao: 139000, venda: -1100000, morte_doacao: -6222.76, reavaliacao: 0},
  '2026-07': {compra: 0, embriao: 82500, venda: -39574.25, morte_doacao: -22500, reavaliacao: 0},
};

const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');
const linha = (rot, a, b) => {
  const d = a - b;
  const marca = Math.abs(d) < 1 ? 'ok' : (b === 0 && a !== 0 ? 'INVENTADO' : 'DIVERGE');
  return `    ${rot.padEnd(14)} apurado ${rs(a).padStart(15)}   divulgado ${rs(b).padStart(15)}`
    + `   dif ${rs(d).padStart(14)}  ${marca}`;
};

const detalhe = process.argv.includes('--animais');
let totalOk = 0, totalLinhas = 0;
for (const mes of Object.keys(DIVULGADO)) {
  const saida = execFileSync('node', [path.join(__dirname, 'motor_plantel.js'), mes,
                                      ...(detalhe ? ['--animais'] : [])],
                             {encoding: 'utf8', maxBuffer: 1 << 24});
  const apurado = {};
  for (const l of saida.split('\n')) {
    const m = /^\s{2}(\w+)\s+(-?R\$ [\d.]+)/.exec(l);
    if (m) apurado[m[1]] = Number(m[2].replace(/[^\d-]/g, '')) * (m[2].startsWith('-') ? 1 : 1);
  }
  const d = DIVULGADO[mes];
  const pares = [
    ['compra', apurado.compra || 0, d.compra],
    ['embriao', apurado.embriao || 0, d.embriao],
    ['venda', apurado.venda || 0, d.venda],
    ['morte+doacao', (apurado.morte || 0) + (apurado.doacao || 0), d.morte_doacao],
    ['reavaliacao', apurado.reavaliacao || 0, d.reavaliacao],
  ];
  console.log(`\n${mes}`);
  for (const [rot, a, b] of pares) {
    console.log(linha(rot, a, b));
    totalLinhas++;
    if (Math.abs(a - b) < 1) totalOk++;
  }
  if (detalhe) console.log(saida.split('\n').filter(l => /^\s{6,}/.test(l)).join('\n'));
}
console.log(`\n${totalOk} de ${totalLinhas} linhas batendo.`);
