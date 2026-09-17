/* Roda o MOTOR da aba Plantel (assets/plantel/plantel.js) fora do navegador,
   contra os arquivos reais do haras.

   Existe porque comparar o resumo contábil com o original no olho não diz QUAL
   animal caiu na causa errada — e mexer na regra sem isso é chute. Aqui o mesmo
   código que roda no hub é carregado com um DOM de mentira, alimentado com os
   xlsx do Drive, e devolve a apuração animal por animal.

   Uso:
     node tools/motor_plantel.js 2026-06            # causas do mês + quem pesa
     node tools/motor_plantel.js 2026-06 --animais  # linha por animal
     node tools/motor_plantel.js --resumo           # todos os meses, só as causas
*/
'use strict';
const fs = require('fs');
const path = require('path');

const RAIZ = path.resolve(__dirname, '..');
const XLSX = require(path.join(RAIZ, 'assets/vendor/xlsx.mini.min.js'));

/* ---- DOM de mentira: o suficiente pro arquivo carregar sem tela ---- */
const elemento = () => new Proxy({}, {
  get: (o, k) => k in o ? o[k]
    : (k === 'classList' ? {add(){}, remove(){}, toggle(){}, contains: () => false}
      : k === 'style' ? {}
      : k === 'dataset' ? {}
      : k === 'querySelectorAll' ? (() => [])
      : k === 'querySelector' ? (() => null)
      : k === 'addEventListener' ? (() => {})
      : k === 'appendChild' ? (() => {})
      : k === 'focus' || k === 'setSelectionRange' ? (() => {})
      : undefined),
  set: (o, k, v) => { o[k] = v; return true; },
});
global.document = {
  getElementById: () => elemento(),
  querySelector: () => null,
  querySelectorAll: () => [],
  body: {addEventListener: () => {}},
  createElement: () => elemento(),
};
global.window = {parent: {}, innerWidth: 1280, innerHeight: 800, location: {hash: ''}};
global.alert = msg => console.error('[alert]', msg);
global.XLSX = XLSX;

/* ---- carrega o motor ---- */
const fonte = fs.readFileSync(path.join(RAIZ, 'assets/plantel/plantel.js'), 'utf8');
// o boot() sai sozinho sem Supabase (sb() devolve null), mas é async: o resto
// do módulo já está definido quando ele começa, então basta não esperar por ele
const modulo = new Function(fonte + '\n;return {ST, lerArquivo, movimentacaoDoMes, resumoAno,'
  + ' patrMes, cotasMes, linhasEfetivas, mesAnterior, rotMes, donoDaLinha, sugere};');
const M = modulo();

/* ---- fontes: os mesmos arquivos que a aba importa ---- */
const _pg = require('child_process');
function arquivoDoMes(mes){
  const [ano, mm] = mes.split('-');
  const saida = _pg.execFileSync('python', ['-c', `
import sys; sys.path.insert(0, r'${path.join(RAIZ, 'scripts').replace(/\\/g, '\\\\')}')
from _pg_common import parse_mes_aaaa, find_source_file
print(find_source_file(parse_mes_aaaa('${mm}/${ano}')))
`], {encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'],
       // sem isto o Python imprime na codepage do console e 'Estação' volta
       // como mojibake, virando caminho que não existe
       env: Object.assign({}, process.env, {PYTHONIOENCODING: 'utf-8'})});
  return saida.trim();
}

function carrega(mes){
  if (M.ST.meses[mes]) return;
  const p = arquivoDoMes(mes);
  M.ST.meses[mes] = M.lerArquivo(fs.readFileSync(p), path.basename(p));
}

/* ---- relatórios ---- */
const rs = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(Math.round(v)).toLocaleString('pt-BR');

function causasDoMes(mes, detalhe){
  carrega(mes); carrega(M.mesAnterior(mes));
  M.ST.mes = mes;
  const mv = M.movimentacaoDoMes(mes);
  const porClasse = {};
  for (const m of mv.movs) {
    if (!m.delta_carla) continue;
    (porClasse[m.sugestao] = porClasse[m.sugestao] || []).push(m);
  }
  console.log(`\n=== ${mes} — causas apuradas (escopo Carla) ===`);
  for (const [cls, itens] of Object.entries(porClasse).sort()) {
    const tot = itens.reduce((s, m) => s + m.delta_carla, 0);
    console.log(`  ${cls.padEnd(13)} ${rs(tot).padStart(16)}   ${itens.length} animal(is)`);
    if (detalhe) {
      for (const m of itens.sort((a, b) => Math.abs(b.delta_carla) - Math.abs(a.delta_carla))) {
        const oc = (m.log || []).map(x => x.ocorrencia).join(' | ').slice(0, 96);
        console.log(`      ${rs(m.delta_carla).padStart(14)}  ${m.nome.slice(0, 38).padEnd(38)}`
          + ` ${m.status || ''}`);
        if (oc) console.log(`                      log: ${oc}`);
      }
    }
  }
}

const args = process.argv.slice(2);
if (args[0] === '--resumo') {
  const meses = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07'];
  for (const m of meses) causasDoMes(m, false);
} else if (args[0]) {
  causasDoMes(args[0], args.includes('--animais'));
} else {
  console.log('uso: node tools/motor_plantel.js <AAAA-MM> [--animais] | --resumo');
}
