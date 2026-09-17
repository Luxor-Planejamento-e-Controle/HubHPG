/* Extrai do MOTOR da aba Plantel tudo que o mapa da Controladoria precisa, em
   JSON, para tools/exporta_mapa.py escrever dentro do template.

   Por que passar por JSON em vez de reimplementar a apuração em Python: o motor
   é um só. As regras de classificação, a decomposição em eventos, a herança de
   comissão e as linhas efetivas moram em assets/plantel/plantel.js e foram
   conferidas contra o Resumo Contábil divulgado (35 de 35 linhas de jan a
   jul/2026). Uma segunda implementação em Python começaria a divergir no
   primeiro mês em que alguém mexesse numa regra e esquecesse da outra.

   A aba `Movimentações` do mapa é CUMULATIVA — jan a dez na mesma linha, com uma
   coluna por causa — então o que sai aqui é o acumulado do ano até o mês pedido,
   animal por animal, e não o movimento do mês.

   Uso: node tools/motor_mapa.js 2026-08 > mapa.json
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
const M = new Function(fonte + `
  ;return {ST, lerArquivo, movimentacaoDoMes, mesAnterior, linhasEfetivasIx, patr,
           comissaoDaLinha, donoDaLinha, chaveCom, ehLinhaReceptoras, resumoAno, num, norm,
           LINHAS_RESUMO};`)();

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
  if (M.ST.meses[mes]) return true;
  try {
    const p = arquivoDoMes(mes);
    if (!p) return false;
    M.ST.meses[mes] = M.lerArquivo(fs.readFileSync(p), path.basename(p));
    return true;
  } catch (e) { return false; }
}

const alvo = process.argv.find(a => /^\d{4}-\d{2}$/.test(a)) || '2026-08';
const ano = alvo.slice(0, 4);
const ateN = +alvo.slice(5);
const meses = [];
for (let n = 1; n <= ateN; n++) meses.push(`${ano}-${String(n).padStart(2, '0')}`);
const base = M.mesAnterior(meses[0]);            // dez do ano anterior
if (!carrega(base)) throw new Error(`sem o arquivo de ${base}, que é a base do ano`);
for (const m of meses) if (!carrega(m)) throw new Error(`sem o arquivo de ${m}`);

/* ---- classe do motor -> coluna do mapa ---- */
const COLUNA = {compra: 'compras', embriao: 'embrioes', venda: 'venda',
                morte: 'morte_doacao', doacao: 'morte_doacao', reavaliacao: 'reavaliacao'};

/* ---- identidade do animal ao longo do ANO ----

   A chave do motor carrega o NOME, e o nome muda: o embrião "LIA DA PAO GRANDE
   X DAMASCO ... RECEP 517" nasce e vira "PRIMAZIA DA PAO GRANDE". Agregando por
   chave de cada mês, o mesmo animal virava duas linhas no mapa — uma com o valor
   de dezembro e nenhum movimento, outra com o movimento e sem base. Foi o que
   fez o teste de julho sair com 398 linhas contra as 351 do mapa da
   Controladoria.

   `movimentacaoDoMes` já resolve o renome DENTRO do mês e devolve os pares em
   `renomes`; aqui eles são encadeados de janeiro até o mês pedido, para que todo
   nome antigo caia no nome que o animal tem hoje. */
const apelido = {};
for (const mes of meses) {
  M.ST.mes = mes;
  for (const r of (M.movimentacaoDoMes(mes).renomes || [])) {
    apelido[M.norm(r.de)] = M.norm(r.para);
  }
}
function canon(nome){
  /* As receptoras são UMA linha agregada cujo nome é a contagem do rebanho
     ("RECEPTORAS 175" vira "RECEPTORAS 124"), então o nome muda todo mês sem que
     o animal mude — o motor já lhes dá identidade própria e aqui vale a mesma,
     senão a base de dezembro fica numa linha e as vendas do ano em outra. */
  if (M.ehLinhaReceptoras(nome)) return 'RECEPTORAS';
  let n = M.norm(nome), voltas = 0;
  while (apelido[n] && apelido[n] !== n && voltas++ < 30) n = apelido[n];
  return n;
}

/* ---- cadastro e valor-base de cada linha, por chave ---- */
const animais = {};
function registra(mes, par){
  const {l, ix} = par;
  const k = canon(l[ix.nome]);
  const a = animais[k] || (animais[k] = {
    chave: k, compras: 0, embrioes: 0, venda: 0, morte_doacao: 0, reavaliacao: 0,
    base: 0, visto: null, aliases: new Set(),
  });
  /* TODO nome que a linha já teve vai junto. A Controladoria não renomeia as
     linhas dela — o mapa de junho ainda chama "RECEPTORAS 121" o que o arquivo
     de julho chama "RECEPTORAS 120", e "MORENA L2 X DAMASCO..." o que virou
     "POTRA MORENA L2 X DAMASCO...". Procurando a linha do template só pelo nome
     de hoje, ela não é achada e o exportador cria uma linha nova ao lado da
     antiga, com o movimento do ano numa e a base na outra. */
  a.aliases.add(M.norm(l[ix.nome]));
  // o cadastro vale o do mês MAIS RECENTE em que a linha aparece
  a.visto = mes;
  a.nome = l[ix.nome];
  a.sufixo = l[ix.sufixo];
  a.categoria = l[ix.categoria];
  a.status = l[ix.status];
  a.cota = M.num(l[ix.cota]);
  a.valor = M.num(l[ix.valor]);
  a.comissao = M.comissaoDaLinha(l, ix);
  a.dono = M.donoDaLinha(l, ix);
  a.patr = a.cota * a.valor + a.comissao;
  return a;
}
// a base do ano é dezembro: é a coluna "VALOR EM DEZ/xxxx" do mapa
for (const par of M.linhasEfetivasIx(base)) {
  const a = registra(base, par);
  a.base = a.patr;
  a.na_base = true;
}
for (const mes of meses) {
  M.ST.mes = mes;
  for (const par of M.linhasEfetivasIx(mes)) registra(mes, par);
  for (const mo of M.movimentacaoDoMes(mes).movs) {
    if (!mo.delta) continue;
    const col = COLUNA[mo.sugestao];
    if (!col) continue;                 // renome, sem efeito: não é linha do mapa
    const k = canon(mo.nome);
    const a = animais[k] || (animais[k] = {
      chave: k, compras: 0, embrioes: 0, venda: 0, morte_doacao: 0, reavaliacao: 0,
      base: 0, visto: mes, nome: mo.nome, sufixo: mo.sufixo, categoria: mo.categoria,
      status: mo.status, cota: mo.cota_atual, valor: mo.valor_atual, dono: mo.dono,
      comissao: 0, patr: mo.patr_atual, aliases: new Set(),
    });
    a.aliases.add(M.norm(mo.nome));
    if (mo.renome) a.aliases.add(M.norm(mo.renome.de));
    a[col] += mo.delta;
  }
}

/* O encadeamento de renomes é a outra ponta: quem aponta para este animal já foi
   o nome dele, mesmo que a linha antiga nunca tenha sido registrada aqui. */
for (const [de, para] of Object.entries(apelido)) {
  const a = animais[canon(para)];
  if (a) a.aliases.add(de);
}

/* ---- linhas da aba Movimentações: só quem tem base ou movimento no ano ---- */
const movimentacoes = Object.values(animais)
  .filter(a => a.base || a.compras || a.embrioes || a.venda || a.morte_doacao || a.reavaliacao)
  .map(a => ({
    chave: a.chave, nome: a.nome, aliases: [...a.aliases],
    sufixo: a.sufixo, categoria: a.categoria,
    status: a.status, cota: a.cota, dono: a.dono,
    base: +a.base.toFixed(2), compras: +a.compras.toFixed(2),
    embrioes: +a.embrioes.toFixed(2), venda: +a.venda.toFixed(2),
    morte_doacao: +a.morte_doacao.toFixed(2), reavaliacao: +a.reavaliacao.toFixed(2),
    final: +(a.base + a.compras + a.embrioes + a.venda + a.morte_doacao + a.reavaliacao).toFixed(2),
  }))
  .sort((x, y) => String(x.nome).localeCompare(String(y.nome), 'pt-BR'));

/* ---- aba PLANTEL: o arquivo do haras do mês, nas linhas efetivas ---- */
M.ST.mes = alvo;
const d = M.ST.meses[alvo];
const plantel = M.linhasEfetivasIx(alvo).map(par => {
  const o = {};
  d.cab.forEach((r, i) => {
    const rot = String(r == null ? '' : r).replace(/\s+/g, ' ').trim();
    if (!rot) return;
    const v = par.l[i];
    o[rot] = v instanceof Date ? v.toISOString().slice(0, 10) : (v == null ? '' : v);
  });
  // os nomes anteriores servem para achar a linha no template (ver casa_linhas)
  const reg = animais[canon(par.l[par.ix.nome])];
  o.__aliases = reg ? [...reg.aliases] : [];
  o.__cota = M.num(par.l[par.ix.cota]);
  o.__valor = M.num(par.l[par.ix.valor]);
  o.__comissao = M.comissaoDaLinha(par.l, par.ix);
  o.__dono = M.donoDaLinha(par.l, par.ix);
  o.__nome = par.l[par.ix.nome];
  return o;
});

/* ---- Resumo Contábil: uma coluna por mês ---- */
const r = M.resumoAno();
const resumo = {};
for (const mes of Object.keys(r).sort()) {
  const c = r[mes].causas;
  resumo[mes] = {
    saldo_ini: +r[mes].ini.toFixed(2), saldo_fim: +r[mes].fim.toFixed(2),
    compras: +(c.compra || 0).toFixed(2),
    embrioes: +(c.embriao || 0).toFixed(2),
    venda: +(c.venda || 0).toFixed(2),
    morte_doacao: +((c.morte || 0) + (c.doacao || 0)).toFixed(2),
    reavaliacao: +(c.reavaliacao || 0).toFixed(2),
  };
}

process.stdout.write(JSON.stringify({mes: alvo, base, movimentacoes, plantel, resumo}, null, 1));
