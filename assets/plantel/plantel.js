/* Plantel Haras Pao Grande — Movimentação.

   Sistema de fechamento mensal. O arquivo é IMPORTADO aqui (nada é puxado do
   Drive sozinho: versão de trabalho muda até o fechamento sair), lido no próprio
   navegador e processado contra o mês anterior. O código propõe; quem registra
   a movimentação é quem está fechando o mês.

   Dois formatos de arquivo entram:
     mapa Luxor  — tem as colunas PLANTEL HPG / PLANTEL EDUARDO. É a semente:
                   traz a atribuição Carla/Eduardo animal por animal, que não é
                   fórmula (a cota inteira vai pra um dos dois, e alguns não vão
                   pra nenhum) e por isso não pode ser calculada.
     arquivo do haras (CONTROLE_DE_PLANTEL) — o mês novo. A atribuição é herdada
                   do mês anterior; animal novo entra na fila pra ser atribuído.

   Escopo: movimentações = Carla + Eduardo · resumo contábil = só Carla. */
'use strict';

const MESES_PT = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const CLASSES_MOV = ['compra', 'embriao', 'venda', 'morte', 'doacao', 'reavaliacao',
                     'renome', 'sem_efeito'];
/* Mesmo corte do banco (plantel_mes_fechado): mês fechado não recebe registro.
   O RLS já recusava, mas calado — a tela oferecia o select e o erro só aparecia
   depois de escolher. */
const FECHADO_ATE = '2026-07';
const mesFechado = m => !!m && m <= FECHADO_ATE;
const ATRIB = {hpg: 'Carla', eduardo: 'Eduardo', nenhum: 'nenhum'};

/* As colunas são localizadas pelo RÓTULO do cabeçalho, não por índice fixo: a
   aba MOVIMENTAÇÕES começa na coluna B e o leitor descarta a coluna vazia da
   frente, então o mesmo campo tem índice diferente em cada aba — e um índice
   fixo passa a ler a coluna errada se o haras inserir uma coluna. */
const ALVOS = {
  qtde: ['QTDE'], letra: ['LETRA'], sufixo: ['SUFIXO'], nome: ['NOME'], sexo: ['SEXO'],
  categoria: ['CATEGORIA'], status: ['STATUS PLANTEL'], local: ['LOCAL'], mae: ['MAE'],
  pai: ['PAI'], nascimento: ['NASCIMENTO'], cota: ['COTAS'], socio: ['NOME SOCIO'],
  condicao: ['CONDICAO ATUAL'], valor: ['VALOR (R$)', 'VALOR 100%', 'VALOR'],
  comissao: ['COMISSAO', 'COMISS'], hpgCota: ['PLANTEL HPG'], edCota: ['PLANTEL EDUARDO'],
};

function idxDe(cab, alvos){
  for (const alvo of alvos) {
    const i = cab.findIndex(c => norm(c).startsWith(alvo));
    if (i >= 0) return i;
  }
  return -1;
}

function mapaColunas(cab){
  const ix = {};
  for (const [campo, alvos] of Object.entries(ALVOS)) ix[campo] = idxDe(cab, alvos);
  // o valor da parceria fica na coluna seguinte à da cota (PLANTEL HPG = cota, ao lado o valor)
  ix.hpgVal = ix.hpgCota >= 0 ? ix.hpgCota + 1 : -1;
  ix.edVal = ix.edCota >= 0 ? ix.edCota + 1 : -1;
  return ix;
}

/* colunas numéricas e de data, para formatar e ordenar certo */
const EH_NUM = r => /QTDE|IDADE|SAFRA|COTAS|PERIODO|VALOR|COMISS|PLANTEL HPG|PLANTEL EDUARDO/.test(norm(r));
const EH_DATA = r => /^DATA|NASCIMENTO/.test(norm(r));
const EH_PCT = r => /^COTAS|PLANTEL HPG|PLANTEL EDUARDO/.test(norm(r));
const EH_MOEDA = r => /VALOR|COMISS/.test(norm(r));

const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
const rs = v => (v == null || v === '' || isNaN(v)) ? '—'
  : (v < 0 ? '−' : '') + 'R$ ' + Math.abs(v).toLocaleString('pt-BR', {maximumFractionDigits: 0});
const pct = v => (v == null || v === '' || isNaN(v)) ? '—'
  : (Number(v) * 100).toLocaleString('pt-BR', {maximumFractionDigits: 2}) + '%';
const clsN = v => !v ? 'zero' : v > 0 ? 'pos' : 'neg';
const rotMes = m => m ? `${MESES_PT[+m.split('-')[1] - 1]}/${m.split('-')[0].slice(2)}` : '';
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/\s+/g, ' ').trim().toUpperCase();
const num = v => { const n = Number(v); return isFinite(n) ? n : 0; };
/* identidade do animal: nome + letra (a coorte). Nome sozinho junta dois potros
   chamados "MACHO ..."; a letra separa. */
const chaveCom = (l, ix) => norm(l[ix.nome]) + '|' + norm(l[ix.letra]);
const chave = l => chaveCom(l, ST.meses[ST.mes] ? ST.meses[ST.mes].ix : {nome: 3, letra: 1});
const dataBR = v => {
  if (v == null || v === '') return '';
  if (v instanceof Date) return v.toLocaleDateString('pt-BR');
  const d = new Date(v);
  return isNaN(d) ? String(v) : d.toLocaleDateString('pt-BR');
};

/* ---- estado ---- */
const ST = {
  meses: {},        // 'AAAA-MM' -> {arquivo, linhas, log, ix}
  /* De quem é cada animal EM CADA MÊS: {mes: {chave: hpg|eduardo|nenhum}}.
     Medido: tratar como atributo fixo do animal erra feio — o dono muda no ano
     (animal da Carla em janeiro vira 'nenhum' depois de repassado), e com uma
     atribuição só o saldo de janeiro saiu R$ 2,49 milhões abaixo do liberado.
     Mês sem atribuição própria herda do mês anterior mais próximo. */
  atrib: {},
  /* Inferência por SUFIXO fica separada e é por ANIMAL, não por mês. Misturada
     com os mapas ela fabricava troca de dono: animal que o mapa de junho não
     cobria (inferido Carla) e que o mapa de julho marca 'nenhum' virava R$ 449
     mil de "mudança de dono" em julho, que nunca aconteceu. Mapa manda;
     inferência é o último recurso. */
  inferido: {},
  sugeridos: {},    // chave -> true enquanto ninguém confirmou
  mes: null,
  decisoes: {},     // 'mes|chave' -> {classe, nota, autor}
  aba: 'plantel',
  sub: 'movimentacoes',
  ordem: {plantel: {col: null, dir: 1}, mov: {col: null, dir: 1}},
  /* filtro por coluna = LISTA de valores marcados (seleção múltipla), não texto */
  filtros: {plantel: {}, mov: {}},
  pop: null,        // {qual, col, busca, foco} — menu de filtro aberto
  ctx: {},          // por aba: base sem filtro + como ler o texto de cada coluna
};

/* ---- sessão do hub ---- */
function hub(){ try { return window.parent.HUB || null; } catch (e) { return null; } }
function sb(){ const h = hub(); return h && h.sb; }
function eu(){ const h = hub(); return (h && h.email) || null; }

/* ================= leitura do arquivo ================= */
function lerArquivo(buf, nomeArquivo){
  const wb = XLSX.read(buf, {cellDates: true});
  const nomePl = wb.SheetNames.find(n => norm(n).startsWith('PLANTEL'));
  if (!nomePl) throw new Error('não achei a aba PLANTEL neste arquivo');

  // acha a linha de cabeçalho pelo próprio conteúdo (a planilha tem título e
  // linha de grupo antes dele, e isso já mudou de posição entre versões)
  const tudo = XLSX.utils.sheet_to_json(wb.Sheets[nomePl], {header: 1, defval: null, raw: true});
  const iCab = tudo.findIndex(r => r && r.some(c => norm(c) === 'NOME') && r.some(c => norm(c).startsWith('CATEGORIA')));
  if (iCab < 0) throw new Error('não achei o cabeçalho da aba PLANTEL (linha com NOME e CATEGORIA)');
  const cab = tudo[iCab];
  const ix = mapaColunas(cab);
  const linhas = tudo.slice(iCab + 1).filter(r => r && r[ix.nome]);

  // log de ocorrências (só o arquivo do haras tem): PRODUTO / DATA / OCORRENCIA
  let log = [];
  const nomeMv = wb.SheetNames.find(n => norm(n).startsWith('MOVIMENTA'));
  if (nomeMv) {
    const m = XLSX.utils.sheet_to_json(wb.Sheets[nomeMv], {header: 1, defval: null, raw: true});
    const iC = m.findIndex(r => r && r.some(c => norm(c).startsWith('OCORRENCIA')));
    if (iC >= 0) {
      const c = m[iC];
      const iProd = idxDe(c, ['PRODUTO']), iData = idxDe(c, ['DATA']), iOc = idxDe(c, ['OCORRENCIA']);
      if (iProd >= 0 && iData >= 0 && iOc >= 0) {
        log = m.slice(iC + 1)
          .filter(r => r && r[iProd] && r[iData] instanceof Date)
          .map(r => ({produto: String(r[iProd]).trim(), data: r[iData], ocorrencia: String(r[iOc] || '').trim()}));
      }
    }
  }
  return {arquivo: nomeArquivo, cab, ix, linhas, log, temSplit: ix.hpgCota >= 0};
}

/* mês de referência: pelo nome do arquivo (MES_26 ou "(Jul 2026)"), senão pergunta */
function mesDoArquivo(nome){
  const n = norm(nome);
  const abr = {JAN: 1, FEV: 2, MAR: 3, ABR: 4, MAI: 5, JUN: 6, JUL: 7, AGO: 8, SET: 9, OUT: 10, NOV: 11, DEZ: 12};
  let m = /(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[ _](\d{2,4})/.exec(n);
  if (m) {
    const ano = m[2].length === 2 ? '20' + m[2] : m[2];
    return `${ano}-${String(abr[m[1]]).padStart(2, '0')}`;
  }
  return null;
}

/* De quem é cada animal. Do arquivo quando ele traz o split (mapa do Luxor),
   senão herdada do mês anterior. Não achou nem um nem outro: SUGERE pelo sufixo
   e marca como sugestão — animal novo aparece na conciliação pra ser confirmado,
   porque a atribuição não é fórmula (no mapa a cota inteira vai pra um dos dois
   e há animal em nenhum dos dois). */
/* O mapa do Luxor é a fonte da atribuição: lê PLANTEL HPG / PLANTEL EDUARDO e
   grava por animal. Ele NÃO vira snapshot de mês — o snapshot é sempre o arquivo
   do haras, senão se compara a população curada do mapa (362 linhas em jul/26)
   com a população cheia do haras (391) e nascem movimentações que não existiram. */
function aplicaAtribuicaoDoMapa(d, mes){
  const alvo = ST.atrib[mes] = ST.atrib[mes] || {};
  let n = 0;
  for (const l of d.linhas) {
    const k = chaveCom(l, d.ix);
    const temHpg = num(l[d.ix.hpgCota]) || num(l[d.ix.hpgVal]);
    const temEd = num(l[d.ix.edCota]) || num(l[d.ix.edVal]);
    alvo[k] = temHpg ? 'hpg' : temEd ? 'eduardo' : 'nenhum';
    delete ST.sugeridos[`${mes}|${k}`];
    n++;
  }
  return n;
}

/* Dono do animal NAQUELE mês: o próprio mês, senão herda do mês anterior mais
   próximo que tenha atribuição. Não havendo nenhum mapa ANTES, vale o mapa mais
   antigo que conheça o animal — dez/25 é base e não tem mapa, e adivinhar por
   sufixo ali fabricava troca de dono em janeiro: R$ 220 mil saíam do saldo da
   Carla em jan/26 como se o animal tivesse sido transferido, quando o mapa de
   janeiro só estava dizendo o que já era verdade em dezembro. Mapa manda em
   qualquer direção; sufixo é o último recurso. */
function donoDe(mes, k){
  const meses = Object.keys(ST.atrib).sort();
  for (const m of meses.filter(m => m <= mes).reverse()) if (k in ST.atrib[m]) return ST.atrib[m][k];
  for (const m of meses) if (k in ST.atrib[m]) return ST.atrib[m][k];
  return k in ST.inferido ? ST.inferido[k] : null;
}

/* Animal que nenhum mapa atribuiu. A regra foi MEDIDA contra os 7 mapas de 2026:

     sufixo simples (DA PAO GRANDE / OUTRO) e com valor -> Carla em 1248 de 1248
     sufixo de parceria (- E xx%) e com valor            -> nenhum 200, Eduardo 30, Carla 1

   Ou seja: sufixo simples com valor pode entrar direto; parceria NÃO se adivinha
   (a cota inteira vai pra um dos dois ou pra nenhum, animal por animal) e vai pra
   conciliação. Sem valor entra como 'nenhum': não mexe no saldo, e se um dia
   ganhar valor a conciliação cobra. */
function completaAtribuicao(mes){
  const d = ST.meses[mes];
  for (const l of d.linhas) {
    const k = chaveCom(l, d.ix);
    if (donoDe(mes, k) != null) continue;
    const suf = norm(l[d.ix.sufixo]);
    const simples = suf === 'DA PAO GRANDE' || suf === 'OUTRO';
    const valor = num(l[d.ix.cota]) * num(l[d.ix.valor]) + num(l[d.ix.comissao]);
    if (!valor) { ST.inferido[k] = 'nenhum'; continue; }
    if (simples) { ST.inferido[k] = 'hpg'; continue; }
    ST.inferido[k] = 'nenhum';
    ST.sugeridos[k] = true;   // parceria com valor: quem fecha decide
  }
}

const mesAnterior = m => {
  if (!m) return null;
  let [a, mm] = m.split('-').map(Number);
  mm -= 1;
  if (!mm) { mm = 12; a -= 1; }
  return `${a}-${String(mm).padStart(2, '0')}`;
};

/* Patrimônio do animal no escopo pedido.
   cota × valor + comissão. A comissão entra porque é assim que o mapa do Luxor
   fecha: PRIMOROSA tem cota 50% de R$120.000 (=60.000) e comissão 5.100, e a
   coluna PLANTEL HPG dela é 65.100. Na soma de jul/2026 isso é a diferença entre
   15.942.120,21 (só cota×valor) e 15.970.552,71 — o saldo do Resumo Contábil
   liberado. */
function patr(l, escopo, ix, mes){
  ix = ix || (ST.meses[ST.mes] && ST.meses[ST.mes].ix) || {};
  const a = donoDe(mes || ST.mes, chaveCom(l, ix));
  if (escopo === 'hpg' && a !== 'hpg') return 0;
  if (escopo === 'carla_eduardo' && a !== 'hpg' && a !== 'eduardo') return 0;
  return num(l[ix.cota]) * num(l[ix.valor]) + num(l[ix.comissao]);
}
/* Linhas do mês já descontando o que foi editado DEPOIS do fim do mês. O arquivo
   do haras é editado durante o mês seguinte, então a mesma planilha contém o
   fechamento e o começo do mês novo; para o animal tocado depois do dia 31 vale
   a linha do mês anterior. Estoque e movimentação usam esta mesma base, senão o
   check "inicial + movimentações = final" não fecha. */
function posterioresDoMes(mes){
  const d = ST.meses[mes];
  if (!d) return {};
  const fim = new Date(+mes.slice(0, 4), +mes.slice(5, 7), 0, 23, 59, 59);
  const o = {};
  for (const x of d.log || []) {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    if (dt > fim) (o[norm(x.produto)] = o[norm(x.produto)] || []).push({...x, data: dt});
  }
  return o;
}

function linhasEfetivas(mes){
  const d = ST.meses[mes];
  if (!d) return [];
  const ant = ST.meses[mesAnterior(mes)];
  const pos = posterioresDoMes(mes);
  if (!ant || !Object.keys(pos).length) return d.linhas;
  const porChave = {};
  for (const l of ant.linhas) porChave[chaveCom(l, ant.ix)] = l;
  return d.linhas.map(l => {
    if (!pos[norm(l[d.ix.nome])]) return l;
    const velha = porChave[chaveCom(l, d.ix)];
    return velha || l;   // sem linha anterior, fica a do arquivo
  });
}

/* Cada linha efetiva pode vir do mês anterior, e o layout de colunas muda de
   arquivo pra arquivo — então a linha anda junto com o seu ix. */
function linhasEfetivasIx(mes){
  const d = ST.meses[mes];
  if (!d) return [];
  const ant = ST.meses[mesAnterior(mes)];
  return linhasEfetivas(mes).map((l, i) =>
    ({l, ix: l === d.linhas[i] ? d.ix : (ant ? ant.ix : d.ix)}));
}

const patrMes = (mes, escopo) =>
  linhasEfetivasIx(mes).reduce((s, par) => s + patr(par.l, escopo, par.ix, mes), 0);

/* ================= movimentação do mês ================= */
const RX_ESTAVA = /ESTAVA (?:COMO )?"?(.+?)"?(?=\s+-\s+(?:MUDOU|TINHA|FOI|PASSOU|ADICAO|ERA|E\s)|\s+PASSOU\s|$)/;

function tipoLog(oc){
  const o = norm(oc);
  if (/MUDOU DE NOME|MUDOU O NOME|ALTEROU O NOME|TROCOU O NOME|ADICAO DE SUFIXO/.test(o)) return 'nome';
  if (/NASCEU/.test(o) && /ESTAVA/.test(o)) return 'nome';
  if (/MUDOU O LOCAL|MUDOU DE LOCAL|FOI PARA O CENTRO DE TREINAMENTO/.test(o)) return 'local';
  return 'financeira';
}

function movimentacaoDoMes(mes){
  const d = ST.meses[mes], ant = ST.meses[mesAnterior(mes)];
  if (!d) return null;
  const ix = d.ix, ixA = ant ? ant.ix : ix;
  /* O arquivo do haras segue sendo editado durante o mês seguinte: a cópia de
     trabalho de setembro/2026 já trazia 15 ocorrências de 01/09 (13 doações pro
     Mato Grosso), que zeram cota e valor. Sem separar isso, o fechamento de
     agosto nasce com R$ 241 mil de baixa que aconteceu em setembro. Animal
     tocado por ocorrência POSTERIOR ao mês fica com o valor do mês anterior e
     aparece na conciliação — não some calado. */
  const fimMes = new Date(+mes.slice(0, 4), +mes.slice(5, 7), 0, 23, 59, 59);
  const posteriores = {};
  for (const x of d.log || []) {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    if (dt > fimMes) (posteriores[norm(x.produto)] = posteriores[norm(x.produto)] || []).push({...x, data: dt});
  }
  const noEscopo = l => ['hpg', 'eduardo'].includes(donoDe(mes, chaveCom(l, ix)));

  const log = (d.log || []).filter(x => {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}` === mes;
  }).map(x => ({...x, tipo: tipoLog(x.ocorrencia)}));

  const logPorNome = {};
  for (const x of log) (logPorNome[norm(x.produto)] = logPorNome[norm(x.produto)] || []).push(x);

  // pontes de identidade: nome antigo -> chave nova (só quando a antiga sumiu)
  const idxAnt = {}, idxAtual = {};
  const efetAnt = ant ? linhasEfetivasIx(mesAnterior(mes)) : [];
  const efet = linhasEfetivasIx(mes);
  for (const par of efetAnt) idxAnt[norm(par.l[par.ix.nome])] = chaveCom(par.l, par.ix);
  for (const par of efet) idxAtual[norm(par.l[par.ix.nome])] = chaveCom(par.l, par.ix);
  const achaNome = (nome, idx) => {
    if (idx[nome]) return idx[nome];
    if (nome.length < 25) return null;
    const c = [...new Set(Object.entries(idx)
      .filter(([n]) => n.startsWith(nome.slice(0, 25)) || nome.startsWith(n.slice(0, 25)))
      .map(([, k]) => k))];
    return c.length === 1 ? c[0] : null;
  };
  const ponte = {}, renomes = [];
  for (const x of log) {
    if (x.tipo !== 'nome') continue;
    const m = RX_ESTAVA.exec(norm(x.ocorrencia));
    if (!m) continue;
    const antigo = norm(m[1]);
    const kOld = achaNome(antigo, idxAnt), kNew = achaNome(norm(x.produto), idxAtual);
    if (!kOld || !kNew || kOld === kNew) continue;
    if (idxAtual[antigo] || (ant && idxAnt[norm(x.produto)])) continue;  // convivem: não é renome
    ponte[kOld] = kNew;
    renomes.push({de: antigo, para: x.produto, data: x.data, chave: kNew});
  }

  /* chave nova -> chave antiga, pra achar o dono do mês anterior de quem foi
     renomeado. Sem isto o animal renomeado entra como se tivesse aparecido do
     nada no escopo da Carla: RECEPTORAS 121->120 sozinho inflava o fluxo de
     julho em R$ 248.910, e o total em R$ 448.984. */
  const ponteInv = {};
  for (const [kOld, kNew] of Object.entries(ponte)) ponteInv[kNew] = kOld;

  /* A abertura sai das linhas EFETIVAS do mês anterior, não do arquivo cru: o
     fechamento de fevereiro já trazia ocorrência de março (OASIS DA PAO GRANDE,
     cota 37,5% -> 42,5%), então o estoque de fevereiro valia a linha de janeiro
     enquanto a abertura de março valia a de fevereiro — R$ 1.500 de fluxo que
     nunca aparecia em movimentação nenhuma. */
  const mapa = (pares) => {
    const o = {};
    for (const par of pares) { const k = chaveCom(par.l, par.ix); o[ponte[k] || k] = par; }
    return o;
  };
  const A = ant ? mapa(efetAnt) : {};
  const B = mapa(efet);

  const movs = [];
  for (const k of new Set([...Object.keys(A), ...Object.keys(B)])) {
    const pa = A[k], pb = B[k];
    const a = pa ? pa.l : null, b = pb ? pb.l : null;
    const ixa = pa ? pa.ix : ixA, ixb = pb ? pb.ix : ix;
    const q0 = a ? num(a[ixa.cota]) : 0, q1 = b ? num(b[ixb.cota]) : 0;
    const v0 = a ? num(a[ixa.valor]) : 0, v1 = b ? num(b[ixb.valor]) : 0;
    // patrimônio na MESMA régua do resumo: cota × valor + comissão
    const p0 = a ? q0 * v0 + num(a[ixa.comissao]) : 0;
    const p1 = b ? q1 * v1 + num(b[ixb.comissao]) : 0;
    const nomeAtual = b ? norm(b[ixb.nome]) : (a ? norm(a[ixa.nome]) : '');
    const posterior = posteriores[nomeAtual] || null;
    const ren = renomes.find(r => r.chave === k);
    const stA = a ? norm(a[ixa.status]) : '', stB = b ? norm(b[ixb.status]) : '';
    const loA = a ? norm(a[ixa.local]) : '', loB = b ? norm(b[ixb.local]) : '';
    const mudouStatus = a && b && stA !== stB;
    const mudouLocal = a && b && loA !== loB;
    const trocouDono = ant && donoDe(mesAnterior(mes), ponteInv[k] || k) !== donoDe(mes, k);
    if (Math.abs(p1 - p0) < 0.01 && !ren && !mudouStatus && !mudouLocal && !trocouDono && a && b) continue;

    const linha = b || a, ixL = b ? ixb : ixa;
    const dono = donoDe(mes, k);
    const kAnt = ponteInv[k] || k;
    const donoAnt = ant ? donoDe(mesAnterior(mes), kAnt) : dono;
    /* Trocar de dono MOVE patrimônio: animal que era da Carla e virou 'nenhum'
       sai do saldo dela mesmo sem venda registrada. Sem isto o check de julho
       acusava R$ 60 mil de diferença entre estoque e fluxo. */
    const noEsc = (dn, v) => dn === 'hpg' ? v : 0;
    const emCE = dn => ['hpg', 'eduardo'].includes(dn);
    const deltaCarla = +(noEsc(dono, p1) - noEsc(donoAnt, p0)).toFixed(2);
    const deltaCE = +((emCE(dono) ? p1 : 0) - (emCE(donoAnt) ? p0 : 0)).toFixed(2);
    const itensLog = logPorNome[norm(linha[ixL.nome])] || (ren ? logPorNome[norm(ren.para)] : []) || [];
    movs.push({
      chave: k, linha, nome: linha[ixL.nome], sufixo: linha[ixL.sufixo],
      categoria: linha[ixL.categoria], status: linha[ixL.status], dono,
      cota_ant: q0, cota_atual: q1, valor_ant: v0, valor_atual: v1,
      patr_ant: p0, patr_atual: p1, delta: +(p1 - p0).toFixed(2),
      entrou: !a, saiu: !b, renome: ren || null,
      dono_ant: donoAnt, delta_carla: deltaCarla, delta_ce: deltaCE,
      mudou_dono: donoAnt !== dono ? [donoAnt, dono] : null,
      mudou_status: mudouStatus ? [a[ixa.status], b[ixb.status]] : null,
      mudou_local: mudouLocal ? [a[ixa.local], b[ixb.local]] : null,
      log: itensLog, posterior,
      sugestao: sugere({q0, q1, v0, v1, p0, p1, ren, entrou: !a, saiu: !b, status: stB || stA,
                        categoria: norm(linha[ixL.categoria]), itensLog}),
      no_escopo: noEscopo(linha),
    });
  }
  movs.sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta));
  return {movs, log, renomes};
}

/* Cancelamento DESFAZ um evento anterior, e a mesma palavra serve pros dois
   sentidos: "VENDA CANCELADA" devolve cota pra Pao Grande, "COMPRA CANCELADA"
   zera a cota de animal que nunca foi nosso. Sem ler isso, o aumento de cota de
   uma venda cancelada era sugerido como 'compra' — evento que não houve.
   As regras são as mesmas de scripts/_pg_cancelamentos.py (58 ocorrências reais
   do controle de plantel); mexer em uma, mexer na outra. */
const RX_CANCEL_TERCEIRO = /CANCEL\w*\s+A\s+COMPRA|COMPRA\s+DE\s+[\d,.]+%?\s+CANCEL/;
const RX_CANCEL_COMPRA = /COMPRA\s+CANCELAD/;
const RX_CANCEL_VENDA = /VENDA[^.]{0,80}?CANCELAD|CANCELAD\w*\s+A\s+VENDA|CANCELAMENTO\s+DE/;
const RX_CANCEL_REVENDA = /CANCELAD\w*[^.]{0,40}?VENDID/;

function cancelamento(oc){
  const s = norm(oc);
  if (!/CANCEL/.test(s)) return null;
  if (RX_CANCEL_REVENDA.test(s)) return {tipo: 'venda', sentido: 'sai', revenda: true};
  if (RX_CANCEL_TERCEIRO.test(s)) return {tipo: 'venda', sentido: 'volta'};
  if (RX_CANCEL_COMPRA.test(s)) return {tipo: 'compra', sentido: 'sai'};
  if (RX_CANCEL_VENDA.test(s)) return {tipo: 'venda', sentido: 'volta'};
  return {tipo: null, sentido: null};
}

/* sugestão pelo padrão — só sugestão; o registro é o input de quem fecha.
   Dinheiro manda na ordem: renome é tipo 1-Nome e não move saldo, então só
   ganha a sugestão quando o patrimônio ficou igual. E patrimônio é
   cota × valor + comissão — olhar só cota e valor deixava mudança de comissão
   como 'sem_efeito'. Juntas, as duas falhas jogavam R$ 241 mil de janeiro/26
   em causas que o resumo não tem linha pra mostrar. */
function sugere({q0, q1, v0, v1, p0, p1, ren, entrou, saiu, status, categoria, itensLog}){
  const st = status || '', cat = categoria || '';
  const oc = norm((itensLog || []).map(x => x.ocorrencia).join(' '));
  const mexeu = Math.abs((p1 || 0) - (p0 || 0)) >= 0.01;
  /* Animal que troca de dono entre um mapa e outro SEM mudar cota nem valor
     não é evento contábil: o fechamento nunca teve essa causa. O que existe é
     diferença de atribuição entre os dois mapas (jan/26: 9 animais que o mapa
     de dez/25 conta como da Carla e o de jan/26 não), e isso se resolve no
     mapa, não virando linha no resumo. Cai em 'sem_efeito' e aparece na coluna
     'O que foi feito' como troca de dono. */
  if (ren && !mexeu) return 'renome';
  if (entrou) return /EMBRI/.test(cat) ? 'embriao' : /NASCEU/.test(oc) ? 'embriao' : 'compra';
  if (saiu || (q0 && !q1)) {
    if (/MORREU|OBITO|ABORTOU/.test(oc) || /OBITO/.test(st)) return 'morte';
    if (/DOAD|DOACAO/.test(oc)) return 'doacao';
    return 'venda';
  }
  /* Cancelamento é ESTORNO: entra na MESMA causa do evento desfeito, com o sinal
     invertido. Venda cancelada devolve cota, então é 'venda' com valor positivo —
     as duas linhas se anulam no resumo, que é o que de fato aconteceu. Sugerir
     'compra' criava uma compra que nunca existiu, e inflava tanto compras quanto
     vendas do mês. */
  const canc = cancelamento(oc);
  if (canc && canc.sentido === 'volta') return 'venda';
  if (canc && canc.tipo === 'compra' && canc.sentido === 'sai') return 'compra';
  if (q1 > q0) return 'compra';
  if (q1 < q0) return /MORREU|OBITO/.test(oc) ? 'morte' : /DOAD/.test(oc) ? 'doacao' : 'venda';
  if (mexeu || v1 !== v0) return 'reavaliacao';
  return 'sem_efeito';
}

/* ================= resumo contábil (YTD, só Carla) =================
   As linhas são AS DA PLANILHA, na mesma ordem e com os mesmos rótulos da aba
   `Resumo Contabil` do mapa (ver tools/seed_plantel_hub.py). O hub reproduz o
   fechamento que já era feito; linha que a planilha não tem, aqui não entra. */
const LINHAS_RESUMO = [
  ['Saldo inicial', null],
  ['(+) Compras', ['compra']],
  ['(+) Produção embriões', ['embriao']],
  ['(−) Baixa vendas', ['venda']],
  ['(−) Baixa mortes e doações', ['morte', 'doacao']],
  ['(+/−) Reavaliações', ['reavaliacao']],
  ['Saldo final', null],
];

function resumoAno(){
  const meses = Object.keys(ST.meses).sort();
  const out = {};
  for (const m of meses) {
    const ant = mesAnterior(m);
    const d = ST.meses[m], da = ST.meses[ant];
    if (!da) continue;
    const ini = patrMes(ant, 'hpg'), fim = patrMes(m, 'hpg');
    const mv = movimentacaoDoMes(m);
    const causas = {};
    for (const mo of (mv ? mv.movs : [])) {
      if (!mo.delta_carla) continue;
      const dec = ST.decisoes[`${m}|${mo.chave}`];
      const classe = dec ? dec.classe : mo.sugestao;
      causas[classe] = +((causas[classe] || 0) + mo.delta_carla).toFixed(2);
    }
    out[m] = {ini: +ini.toFixed(2), fim: +fim.toFixed(2), causas,
              registrado: (mv ? mv.movs : []).filter(mo => ST.decisoes[`${m}|${mo.chave}`]).length,
              total: (mv ? mv.movs : []).length};
  }
  return out;
}

/* ================= UI ================= */
function topo(){
  const meses = Object.keys(ST.meses).sort();
  // sem <h1> aqui: o cabeçalho do hub já mostra o nome da aba, e repetir era
  // exatamente o tipo de poluição que o painel não precisa
  document.getElementById('topo').innerHTML = `
    <div class="ferramentas">
      <label class="mes-sel">Mês
        <select id="selMes">${meses.map(m => `<option value="${m}"${m === ST.mes ? ' selected' : ''}>${rotMes(m)}</option>`).join('')}</select>
      </label>
      <label class="botao-arquivo">Importar arquivo
        <input type="file" id="arq" accept=".xlsx,.xlsm" hidden>
      </label>
      <span id="statusImp"></span>
    </div>`;
  document.getElementById('selMes').onchange = e => {
    ST.mes = e.target.value;
    /* Filtro e ordenação são por coluna DAQUELE mês, então trocar de mês limpa
       os dois — mas mantendo a forma {plantel, mov}. Zerando pra {} e {col,dir}
       todo render seguinte morria em Object.entries(undefined), e a tela ficava
       congelada no mês anterior. */
    ST.ordem = {plantel: {col: null, dir: 1}, mov: {col: null, dir: 1}};
    ST.filtros = {plantel: {}, mov: {}};
    ST.pop = null;
    pinta();
  };
  document.getElementById('arq').onchange = importa;
}

/* O mês NÃO é decidido pelo nome do arquivo: a cópia de trabalho do haras se
   chama ..._JUL_26.xlsx enquanto já contém agosto. O nome só sugere; quem
   confirma é quem está importando. */
async function importa(e){
  const f = e.target.files[0];
  if (!f) return;
  const st = document.getElementById('statusImp');
  st.textContent = 'lendo ' + f.name + '…';
  try {
    const d = lerArquivo(await f.arrayBuffer(), f.name);
    st.textContent = '';
    confirmaImport(d, f.name, mesDoArquivo(f.name));
  } catch (err) {
    st.textContent = 'erro: ' + (err.message || err);
  }
  e.target.value = '';
}

function confirmaImport(d, arquivo, mesSugerido){
  const ano = (mesSugerido || ST.mes || '2026').slice(0, 4);
  const ultimo = (d.log || []).reduce((mx, x) => {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    return !mx || dt > mx ? dt : mx;
  }, null);
  const caixa = document.getElementById('import');
  caixa.hidden = false;
  caixa.innerHTML = `
    <div><b>${esc(arquivo)}</b> — ${d.linhas.length} animais${d.log.length ? ` · ${d.log.length} ocorrências` : ''}
      ${ultimo ? `<br><span class="log">última ocorrência no arquivo: ${ultimo.toLocaleDateString('pt-BR')}</span>` : ''}</div>
    ${d.temSplit
      ? `<div>Tem as colunas <b>PLANTEL HPG</b> e <b>PLANTEL EDUARDO</b>: entra como
           <b>atribuição de ${rotMes(mesSugerido || ST.mes) || 'mês a definir'}</b> — de quem é cada
           animal naquele mês —, sem virar fechamento.</div>
         <button type="button" id="impAtrib"${mesSugerido || ST.mes ? '' : ' disabled'}>Usar a atribuição</button>`
      : `<label>Mês de fechamento
           <select id="impMes">${MESES_PT.map((r, i) => {
             const m = `${ano}-${String(i + 1).padStart(2, '0')}`;
             return `<option value="${m}"${m === mesSugerido ? ' selected' : ''}>${r}/${ano.slice(2)}</option>`;
           }).join('')}</select></label>
         <button type="button" id="impOk">Processar</button>`}
    <button type="button" id="impNao" class="secundario">Cancelar</button>`;
  const fecha = () => { caixa.hidden = true; caixa.innerHTML = ''; };
  document.getElementById('impNao').onclick = fecha;
  if (d.temSplit) {
    document.getElementById('impAtrib').onclick = async () => {
      const n = aplicaAtribuicaoDoMapa(d, mesSugerido || ST.mes);
      fecha();
      document.getElementById('statusImp').textContent = `atribuição de ${n} animais carregada`;
      await salvaAtribuicao();
      pinta();
    };
    return;
  }
  document.getElementById('impOk').onclick = async () => {
    const mes = document.getElementById('impMes').value;
    ST.meses[mes] = d;
    completaAtribuicao(mes);
    ST.mes = mes;
    fecha();
    await salvaSnapshot(mes, d, arquivo);
    topo(); pinta();
  };
}

async function salvaAtribuicao(){
  const c = sb();
  if (!c || !ST.mes) return;
  try { await c.from('plantel_snapshot').update({atribuicao: ST.atrib[ST.mes] || {}}).eq('mes', ST.mes); }
  catch (e) { console.warn('[plantel] atribuição não gravada', e.message || e); }
}

async function salvaSnapshot(mes, d, arquivo){
  const c = sb();
  if (!c) return;
  try {
    await c.from('plantel_snapshot').upsert({
      mes, arquivo,
      linhas: {cab: d.cab, ix: d.ix, rows: d.linhas},
      log: d.log.map(x => ({...x, data: (x.data instanceof Date ? x.data : new Date(x.data)).toISOString()})),
      atribuicao: ST.atrib[mes] || {},
    }, {onConflict: 'mes'});
  } catch (err) { console.warn('[plantel] snapshot não gravado', err.message || err); }
}

/* `linhas` no banco é autocontido: {cab, ix, rows}. O cabeçalho e o mapa de
   colunas viajam junto com as linhas porque a planilha do haras muda de layout
   entre versões — guardar só as linhas obrigaria a tela a adivinhar em que
   coluna está o quê. Formato antigo (array puro) ainda é aceito. */
function normalizaSnapshot(r){
  const bruto = r.linhas;
  const pacote = Array.isArray(bruto) ? {rows: bruto, cab: [], ix: {}} : (bruto || {});
  return {
    arquivo: r.arquivo,
    cab: pacote.cab || [],
    ix: pacote.ix || {},
    linhas: pacote.rows || [],
    liberado: pacote.liberado || null,   // Resumo Contábil que foi divulgado
    log: (r.log || []).map(x => ({...x, data: new Date(x.data)})),
    temSplit: (pacote.ix || {}).hpgCota >= 0,
  };
}

async function carregaSnapshots(){
  const c = sb();
  if (!c) return;
  try {
    const { data, error } = await c.from('plantel_snapshot')
      .select('mes,arquivo,linhas,log,atribuicao').order('mes');
    if (error) throw error;
    for (const r of data || []) {
      ST.meses[r.mes] = normalizaSnapshot(r);
      ST.atrib[r.mes] = r.atribuicao || {};
    }
    // O mapa do mês não cobre todo animal do fechamento: em jan/26 faltavam
    // R$ 612 mil e em abr/26 R$ 261 mil de animais que existem no arquivo do
    // haras e não no mapa. A mesma regra do import preenche o resto, em ordem
    // de mês — a atribuição de um mês herda do anterior.
    for (const m of Object.keys(ST.meses).sort()) completaAtribuicao(m);
    const { data: dec } = await c.from('plantel_mov_classificacao').select('mes,chave,classe,nota,autor');
    for (const r of dec || []) ST.decisoes[`${r.mes}|${r.chave}`] = r;
  } catch (err) { console.warn('[plantel] sem dados salvos', err.message || err); }
}

async function registra(mes, mov, classe, nota){
  const k = `${mes}|${mov.chave}`, antes = ST.decisoes[k];
  ST.decisoes[k] = {classe, nota, autor: eu() || '(local)'};
  const c = sb();
  if (c) {
    const { error } = await c.from('plantel_mov_classificacao')
      .upsert({mes, chave: mov.chave, classe, nota: nota || null, nome: mov.nome}, {onConflict: 'mes,chave'});
    /* banco recusou: desfazer o estado local. Sem isto a linha ficava verde de
       'registrada' na tela e vazia no banco — o alerta aparecia, a pessoa
       seguia, e o registro só sumia no próximo carregamento da aba. */
    if (error) {
      if (antes) ST.decisoes[k] = antes; else delete ST.decisoes[k];
      alert('não gravou: ' + error.message);
      return false;
    }
  }
  return true;
}

async function atribui(mes, k, dono){
  // decisão humana vale daquele mês em diante (entra como se fosse mapa)
  (ST.atrib[mes] = ST.atrib[mes] || {})[k] = dono;
  delete ST.sugeridos[k];
  const c = sb();
  if (c) await c.from('plantel_snapshot').update({atribuicao: ST.atrib[mes]}).eq('mes', mes);
}

/* ---------- filtro por seleção (estilo planilha) ----------
   Era uma caixa de texto por coluna: pra ver DOIS valores ao mesmo tempo
   (garanhão e matriz, venda e morte) não havia jeito — texto casa um padrão só,
   e ninguém adivinha a grafia exata do que está na célula. Agora cada coluna
   abre a lista dos valores que existem nela, com caixa de seleção.
   O menu mora FORA de #painel: o painel inteiro é redesenhado a cada marcação,
   e um menu desenhado dentro dele morria no primeiro clique. */

/* aplica os filtros da aba. `pular` deixa uma coluna de fora — é assim que a
   lista de uma coluna continua oferecendo os valores dela mesma depois de
   marcar alguns (senão sobrava só o que já estava marcado, e não dava pra
   ampliar a seleção sem limpar tudo). */
function aplicaFiltros(itens, qual, txtDe, pular){
  for (const [ci, sel] of Object.entries(ST.filtros[qual] || {})) {
    const i = +ci;
    if (!sel || !sel.length || i === pular) continue;
    const set = new Set(sel);
    itens = itens.filter(it => set.has(txtDe(it, i)));
  }
  return itens;
}

function opcoesDe(itens, qual, txtDe, ordDe, i){
  const m = new Map();
  for (const it of aplicaFiltros(itens, qual, txtDe, i)) {
    const t = txtDe(it, i), o = m.get(t);
    if (o) o.n++; else m.set(t, {n: 1, ord: ordDe(it, i)});
  }
  // valor marcado que sumiu da base continua na lista, senão não dá pra desmarcar
  for (const v of ST.filtros[qual][i] || []) if (!m.has(v)) m.set(v, {n: 0, ord: v});
  return [...m.entries()].sort((a, b) =>
    typeof a[1].ord === 'number' && typeof b[1].ord === 'number'
      ? a[1].ord - b[1].ord
      : String(a[0]).localeCompare(String(b[0]), 'pt-BR', {numeric: true}));
}

/* O filtro mora NO cabeçalho, ao lado do nome da coluna — é onde a planilha
   põe e onde a mão procura. A linha extra de caixas embaixo do cabeçalho saía
   cara duas vezes: ocupava altura e, por ter largura mínima própria, esticava
   coluna estreita (QTDE, LETRA, SEXO) muito além do texto que mostra. O que
   está filtrado é dito ACIMA da tabela, em etiqueta, não dentro da coluna. */
function iconeFiltro(qual, i, rot){
  const sel = ST.filtros[qual][i] || [];
  return `<button type="button" class="fbtn${sel.length ? ' on' : ''}" data-fb="${qual}:${i}"`
    + ` title="filtrar ${esc(rot)}">▾</button>`;
}

function cabFiltro(qual, i, rot, ehNum, ordem){
  const seta = ordem.col === i ? (ordem.dir > 0 ? ' ▲' : ' ▼') : '';
  return `<th data-ord="${qual}:${i}" class="${ehNum ? '' : 'l'}${ordem.col === i ? ' ord' : ''}">`
    + `<span class="th-in"><span class="th-rot">${esc(rot)}${seta}</span>`
    + `${iconeFiltro(qual, i, rot)}</span></th>`;
}

/* etiquetas do que está filtrado, acima da tabela: sem elas o filtro some de
   vista assim que o menu fecha — só o ícone aceso denuncia, e ícone aceso em
   coluna que rolou pra fora da tela não denuncia nada. */
function chipsFiltro(qual){
  const ctx = ST.ctx[qual] || {rotulos: {}};
  const ativos = Object.entries(ST.filtros[qual]).filter(([, v]) => v && v.length);
  if (!ativos.length) return '';
  return `<div class="chips">${ativos.map(([ci, v]) => {
      const um = String(v[0] == null ? '' : v[0]);
      const txt = v.length === 1 ? (um.length > 22 ? um.slice(0, 21) + '…' : um)
        : v.length + ' itens';
      return `<span class="chip" title="${esc(v.join(' · '))}">`
        + `<b>${esc(ctx.rotulos[ci] || '')}</b>${esc(txt)}`
        + `<button type="button" data-fx="${qual}:${ci}" title="remover">×</button></span>`;
    }).join('')}<button type="button" class="chip-limpa" data-fzerar="${qual}">limpar tudo</button></div>`;
}

function pintaPop(){
  const el = document.getElementById('fpop');
  if (!el) return;
  const ctx = ST.pop ? ST.ctx[ST.pop.qual] : null;
  if (!ST.pop || !ctx) { el.hidden = true; el.innerHTML = ''; return; }
  const {qual, col, busca} = ST.pop;
  const opts = opcoesDe(ctx.base, qual, ctx.txtDe, ctx.ordDe, col);
  const sel = new Set(ST.filtros[qual][col] || []);
  const b = norm(busca || '');
  const vis = b ? opts.filter(([t]) => norm(t).includes(b)) : opts;
  el.hidden = false;
  el.innerHTML = `
    <div class="fpop-cab">${esc(ctx.rotulos[col] || '')}</div>
    <input class="fpop-busca" id="fpopBusca" placeholder="buscar…" value="${esc(busca || '')}">
    <div class="fpop-acoes">
      <button type="button" data-fpop="todos">marcar ${vis.length}</button>
      <button type="button" data-fpop="nenhum">desmarcar</button>
      <button type="button" data-fpop="fecha">fechar</button>
    </div>
    <div class="fpop-lista">${vis.map(([t, o]) =>
      `<label><input type="checkbox" data-fv="${esc(t)}"${sel.has(t) ? ' checked' : ''}>`
      + `<span class="fpop-txt">${esc(t) || '<i>(vazio)</i>'}</span>`
      + `<span class="fpop-n">${o.n}</span></label>`).join('')
      || '<div class="fpop-vazio">nada encontrado</div>'}</div>`;
  posPop();
  if (ST.pop.foco) {
    const bs = document.getElementById('fpopBusca');
    if (bs) { bs.focus(); bs.setSelectionRange(bs.value.length, bs.value.length); }
  }
}

/* o menu é fixed e se pendura no botão da coluna; como o painel é redesenhado
   inteiro a cada marcação, a posição é recalculada depois de cada pintura */
function posPop(){
  const el = document.getElementById('fpop');
  if (!ST.pop || !el || el.hidden) return;
  const b = document.querySelector(`[data-fb="${ST.pop.qual}:${ST.pop.col}"]`);
  if (!b) { ST.pop = null; el.hidden = true; el.innerHTML = ''; return; }
  const r = b.getBoundingClientRect(), larg = el.offsetWidth || 250, alt = el.offsetHeight || 300;
  el.style.left = Math.max(6, Math.min(r.left, window.innerWidth - larg - 6)) + 'px';
  el.style.top = Math.max(6, Math.min(r.bottom + 4, window.innerHeight - alt - 6)) + 'px';
}

function marcaValor(qual, col, valor, ligado){
  const atual = new Set(ST.filtros[qual][col] || []);
  if (ligado) atual.add(valor); else atual.delete(valor);
  if (atual.size) ST.filtros[qual][col] = [...atual];
  else delete ST.filtros[qual][col];
}

/* ---------- aba Plantel: todas as colunas, filtro e ordenação por coluna ---------- */
/* Plantel: as colunas são as do arquivo, todas, na ordem dele. Filtro e
   ordenação por coluna, como na planilha. */
function colunasDoMes(d){
  return d.cab.map((r, i) => [i, String(r == null ? '' : r).replace(/\s+/g, ' ').trim()])
    .filter(([, r]) => r);
}

function painelPlantel(){
  const d = ST.meses[ST.mes];
  if (!d) return semArquivo();
  const ix = d.ix;
  const cols = colunasDoMes(d);
  // o plantel do mês é o do FECHAMENTO: linha tocada depois do dia 31 entra com
  // o valor do mês anterior (ver linhasEfetivas), senão a capa mostra um número
  // e o resumo contábil mostra outro
  const base = linhasEfetivas(ST.mes).slice();
  ST.ctx.plantel = {base, txtDe: txtPl, ordDe: ordPl, rotulos: Object.fromEntries(cols)};
  let linhas = aplicaFiltros(base, 'plantel', txtPl);
  if (ST.ordem.plantel.col != null) {
    const i = ST.ordem.plantel.col, dir = ST.ordem.plantel.dir, rot = d.cab[i];
    linhas.sort((a, b) => EH_NUM(rot)
      ? (num(a[i]) - num(b[i])) * dir
      : String(a[i] == null ? '' : a[i]).localeCompare(String(b[i] == null ? '' : b[i]), 'pt-BR') * dir);
  }
  const somaEsc = escopo => linhas.reduce((s, l) => s + patr(l, escopo, ix), 0);
  return `
    <div class="resumo-linha">
      <span>${linhas.length} de ${d.linhas.length} linhas</span>
      <span>Carla: <b>${rs(somaEsc('hpg'))}</b></span>
      <span>Carla + Eduardo: <b>${rs(somaEsc('carla_eduardo'))}</b></span>
    </div>
    ${chipsFiltro('plantel')}
    <div class="rolagem"><table class="t">
      <thead>
        <tr>${cols.map(([i, r]) => cabFiltro('plantel', i, r, EH_NUM(r), ST.ordem.plantel)).join('')}<th class="l">Dono</th></tr>
      </thead>
      <tbody>${linhas.map(l => `<tr>${cols.map(([i, r]) =>
        `<td class="${EH_NUM(r) ? '' : 'l'}">${esc(fmtCel(l, i, r))}</td>`).join('')}
        <td class="l">${ATRIB[donoDe(ST.mes, chaveCom(l, ix))] || '<span class="zero">—</span>'}${
          ST.sugeridos[chaveCom(l, ix)] ? ' <span class="sug">confirmar</span>' : ''}</td></tr>`).join('')}</tbody>
    </table></div>`;
}

/* o texto da célula é o mesmo que a lista de filtro oferece — se as duas
   strings divergirem, marcar um valor não acha linha nenhuma */
const txtPl = (l, i) => {
  const d = ST.meses[ST.mes];
  return (d ? fmtCel(l, i, d.cab[i]) : '') || '—';
};
const ordPl = (l, i) => {
  const d = ST.meses[ST.mes];
  return d && EH_NUM(d.cab[i]) ? num(l[i]) : String(l[i] == null ? '' : l[i]);
};

function fmtCel(l, i, rotulo){
  const v = l[i];
  if (v == null || v === '') return '';
  if (EH_DATA(rotulo)) return dataBR(v);
  if (EH_PCT(rotulo)) return pct(v);
  if (EH_MOEDA(rotulo)) return rs(num(v));
  if (/IDADE/.test(norm(rotulo))) return num(v).toFixed(1);
  return String(v);
}

/* ---------- aba Movimentações ---------- */
function painelMov(){
  const d = ST.meses[ST.mes];
  if (!d) return semArquivo();
  const sub = ST.sub;
  const menu = [['movimentacoes', 'Movimentações'], ['conciliacao', 'Conciliação'], ['checks', 'Checks']];
  const corpo = sub === 'conciliacao' ? subConciliacao() : sub === 'checks' ? subChecks() : subMovimentacoes();
  return `<div class="submenu">${menu.map(([id, t]) =>
      `<button type="button" data-sub="${id}" class="${sub === id ? 'on' : ''}">${t}</button>`).join('')}</div>
    ${corpo}`;
}

/* colunas da aba Movimentações: rótulo, como extrair e se é número (pra ordenar
   e filtrar por coluna, igual à planilha) */
const COLS_MOV = [
  ['Nome', m => m.nome, 0],
  ['Sufixo', m => m.sufixo, 0],
  ['Categoria', m => m.categoria, 0],
  ['Status', m => m.status, 0],
  ['Dono', m => ATRIB[m.dono] || '', 0],
  ['Cota', m => m.cota_atual, 1],
  ['Valor', m => m.valor_atual, 1],
  ['Valor inicial', m => m.patr_ant, 1],
  ['Compras', m => classeDe(m) === 'compra' ? m.delta : 0, 1],
  ['Embriões', m => classeDe(m) === 'embriao' ? m.delta : 0, 1],
  ['Venda', m => classeDe(m) === 'venda' ? m.delta : 0, 1],
  ['Morte/doação', m => ['morte', 'doacao'].includes(classeDe(m)) ? m.delta : 0, 1],
  ['Reavaliação', m => classeDe(m) === 'reavaliacao' ? m.delta : 0, 1],
  ['Valor final', m => m.patr_atual, 1],
];
const classeDe = m => {
  const d = ST.decisoes[`${ST.mes}|${m.chave}`];
  return d ? d.classe : m.sugestao;
};
// colunas que carregam movimento de patrimônio (ganham sinal e cor)
const COLS_DELTA = ['Compras', 'Embriões', 'Venda', 'Morte/doação', 'Reavaliação'];

/* o texto da célula é o mesmo que a lista de filtro oferece — se as duas
   strings divergirem, marcar um valor não acha linha nenhuma */
function txtMov(m, i){
  const [rot, pega, ehNum] = COLS_MOV[i];
  const v = pega(m);
  if (!ehNum) return String(v == null ? '' : v) || '—';
  if (rot === 'Cota') return pct(v);
  if (COLS_DELTA.includes(rot)) return v ? rs(v) : '—';
  return rs(v);
}
const ordMov = (m, i) => COLS_MOV[i][2] ? num(COLS_MOV[i][1](m)) : String(COLS_MOV[i][1](m) || '');

function subMovimentacoes(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  if (!ST.meses[mesAnterior(ST.mes)]) {
    return `<div class="aviso">Importe também ${rotMes(mesAnterior(ST.mes))} para comparar os dois meses.</div>`;
  }
  const base = mv.movs.filter(m => m.no_escopo || m.dono);
  ST.ctx.mov = {base, txtDe: txtMov, ordDe: ordMov,
                rotulos: Object.fromEntries(COLS_MOV.map(([rot], i) => [i, rot]))};
  let movs = aplicaFiltros(base, 'mov', txtMov);
  if (ST.ordem.mov.col != null) {
    const i = ST.ordem.mov.col, dir = ST.ordem.mov.dir, pega = COLS_MOV[i][1], ehNum = COLS_MOV[i][2];
    movs = movs.slice().sort((a, b) => ehNum
      ? (num(pega(a)) - num(pega(b))) * dir
      : String(pega(a) || '').localeCompare(String(pega(b) || ''), 'pt-BR') * dir);
  }
  const soma = k => movs.reduce((s, m) => {
    const dec = ST.decisoes[`${ST.mes}|${m.chave}`];
    return s + ((dec ? dec.classe : m.sugestao) === k ? m.delta : 0);
  }, 0);
  return `
    <div class="resumo-linha">
      <span>${movs.length}${movs.length === base.length ? '' : ' de ' + base.length} animais com movimentação em ${rotMes(ST.mes)}</span>
      <span>registrados: <b>${movs.filter(m => ST.decisoes[`${ST.mes}|${m.chave}`]).length}</b> de ${movs.length}</span>
      <span>Δ patrimônio: <b class="${clsN(movs.reduce((s, m) => s + m.delta, 0))}">${rs(movs.reduce((s, m) => s + m.delta, 0))}</b></span>
    </div>
    ${chipsFiltro('mov')}
    <div class="rolagem"><table class="t">
      <thead>
        <tr>${COLS_MOV.map(([rot, , ehNum], i) => {
            const th = cabFiltro('mov', i, rot, ehNum, ST.ordem.mov);
            // o select de registro fica logo depois do nome: é a coluna de ação,
            // e no fim da tabela ela caía fora da tela
            return i === 0 ? th + '<th class="l">Registro</th>' : th;
          }).join('')}<th class="l">O que foi feito</th></tr>
      </thead>
      <tbody>${movs.map(m => linhaMov(m)).join('')}
        <tr class="tot">${COLS_MOV.map(([rot, pega, ehNum], i) => {
          if (i === 0) return `<td>total</td><td></td>`;
          if (!ehNum || rot === 'Cota') return '<td></td>';
          const s = movs.reduce((a, m) => a + num(pega(m)), 0);
          return `<td class="${['Valor', 'Valor inicial', 'Valor final'].includes(rot) ? '' : clsN(s)}">${rs(s)}</td>`;
        }).join('')}<td></td></tr>
      </tbody></table></div>`;
}

function linhaMov(m){
  const dec = ST.decisoes[`${ST.mes}|${m.chave}`];
  const classe = dec ? dec.classe : null;
  const trancado = mesFechado(ST.mes);
  const celReg = `<td class="l">
      <select data-reg="${esc(m.chave)}"${trancado ? ' disabled title="mês fechado"' : ''}>
        <option value="">— registrar —</option>
        ${CLASSES_MOV.map(c => `<option value="${c}"${classe === c ? ' selected' : ''}>${c}${
          !classe && c === m.sugestao ? ' (sugerido)' : ''}</option>`).join('')}
      </select>${dec ? `<span class="autor">${esc(dec.autor || '')}</span>` : ''}
    </td>`;

  const oque = [];
  const canc = (m.log || []).map(x => cancelamento(x.ocorrencia)).find(c => c && c.tipo);
  if (canc) oque.push(canc.revenda
    ? '<b>venda cancelada e revendido</b>'
    : `<b>${canc.tipo} cancelada</b> (cota ${canc.sentido === 'volta' ? 'volta' : 'sai'})`);
  if (m.renome) oque.push(`renome: <b>${esc(m.renome.de)}</b> → <b>${esc(m.renome.para)}</b>`);
  if (m.entrou) oque.push('entrou no plantel');
  if (m.saiu) oque.push('saiu do controle');
  if (m.mudou_dono) oque.push(`dono: <b>${ATRIB[m.mudou_dono[0]] || '—'}</b> → <b>${ATRIB[m.mudou_dono[1]] || '—'}</b>`);
  if (m.mudou_status) oque.push(`status: ${esc(m.mudou_status[0])} → ${esc(m.mudou_status[1])}`);
  if (m.mudou_local) oque.push(`local: ${esc(m.mudou_local[0])} → ${esc(m.mudou_local[1])}`);
  if (m.cota_ant !== m.cota_atual) oque.push(`cota: ${pct(m.cota_ant)} → ${pct(m.cota_atual)}`);
  if (m.valor_ant !== m.valor_atual) oque.push(`valor: ${rs(m.valor_ant)} → ${rs(m.valor_atual)}`);
  for (const l of m.log) oque.push(
    `<span class="log" title="${esc(l.ocorrencia)}">${dataBR(l.data)} · ${esc(l.ocorrencia)}</span>`);
  if (m.posterior) oque.push(`<span class="log" title="${esc(m.posterior.map(x => x.ocorrencia).join(' | '))}">`
    + `⚠ ${m.posterior.length} ocorrência(s) posterior(es) ao mês — mantido o valor do mês anterior</span>`);

  const cels = COLS_MOV.map(([rot, pega, ehNum], i) => {
    const v = pega(m), t = esc(txtMov(m, i));
    let txt;
    if (!ehNum) txt = t === '—' ? '<span class="zero">—</span>' : t;
    else if (COLS_DELTA.includes(rot)) txt = v ? `<span class="${clsN(v)}">${t}</span>` : '—';
    else txt = t;
    const td = `<td class="${ehNum ? '' : 'l'}${i === 0 ? ' nome' : ''}"${
      i === 0 ? ` title="${esc(m.nome)}"` : ''}>${txt}</td>`;
    return i === 0 ? td + celReg : td;
  }).join('');

  return `<tr class="${dec ? 'reg' : ''}">${cels}<td class="l oque">${oque.join('<br>') || '—'}</td></tr>`;
}

function subConciliacao(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  const dmes = ST.meses[ST.mes];
  const semDono = Object.keys(ST.sugeridos)
    .filter(k => dmes.linhas.some(l => chaveCom(l, dmes.ix) === k));
  const semLog = mv.movs.filter(m => !m.log.length && Math.abs(m.delta) >= 1 && (m.no_escopo || m.dono));
  const logSemEfeito = mv.log.filter(l => l.tipo === 'financeira' &&
    !mv.movs.some(m => norm(m.nome) === norm(l.produto)));
  const naoRegistrados = mv.movs.filter(m => (m.no_escopo || m.dono) && !ST.decisoes[`${ST.mes}|${m.chave}`]);
  const bloco = (titulo, itens, render) => !itens.length ? '' :
    `<h3>${titulo} <span class="cont">${itens.length}</span></h3>${itens.map(render).join('')}`;
  if (!semDono.length && !semLog.length && !logSemEfeito.length && !naoRegistrados.length) {
    return `<div class="ok-vazio">Tudo conciliado em ${rotMes(ST.mes)}.</div>`;
  }
  return `
    ${bloco('Animal novo sem dono definido', semDono, k => {
      const dm = ST.meses[ST.mes], ixm = dm.ix;
      const l = dm.linhas.find(x => chaveCom(x, ixm) === k) || [];
      return `<div class="item"><b>${esc(l[ixm.nome])}</b> · ${esc(l[ixm.sufixo])} · ${esc(l[ixm.categoria])} ·
        cota ${pct(l[ixm.cota])} · ${rs(num(l[ixm.cota]) * num(l[ixm.valor]) + num(l[ixm.comissao]))}
        · sugerido: <b>${ATRIB[donoDe(ST.mes, k)]}</b>
        <span class="acoes">${['hpg', 'eduardo', 'nenhum'].map(dn =>
          `<button type="button" data-dono="${esc(k)}:${dn}">${ATRIB[dn]}</button>`).join('')}</span></div>`;
    })}
    ${bloco('Movimentação sem ocorrência no log do haras', semLog, m =>
      `<div class="item"><b>${esc(m.nome)}</b> · ${rs(m.delta)} ·
        cota ${pct(m.cota_ant)} → ${pct(m.cota_atual)} · valor ${rs(m.valor_ant)} → ${rs(m.valor_atual)}
        · status ${esc(m.status)}</div>`)}
    ${bloco('Ocorrência no log sem efeito no patrimônio', logSemEfeito, l =>
      `<div class="item"><b>${esc(l.produto)}</b> · ${dataBR(l.data)}<br><span class="log">${esc(l.ocorrencia)}</span></div>`)}
    ${bloco('Movimentação ainda não registrada', naoRegistrados, m =>
      `<div class="item"><b>${esc(m.nome)}</b> · ${rs(m.delta)} · sugestão: <b>${m.sugestao}</b></div>`)}`;
}

function subChecks(){
  const d = ST.meses[ST.mes], ant = ST.meses[mesAnterior(ST.mes)];
  if (!d) return semArquivo();
  const mv = movimentacaoDoMes(ST.mes);
  const iniC = patrMes(mesAnterior(ST.mes), 'hpg'), fimC = patrMes(ST.mes, 'hpg');
  const movC = (mv ? mv.movs : []).reduce((s, m) => s + (m.delta_carla || 0), 0);
  const iniCE = patrMes(mesAnterior(ST.mes), 'carla_eduardo'), fimCE = patrMes(ST.mes, 'carla_eduardo');
  const movCE = (mv ? mv.movs : []).reduce((s, m) => s + (m.delta_ce || 0), 0);
  // o resumo contábil só conta o que foi REGISTRADO; a movimentação apurada conta
  // tudo que mudou. Divergir aqui é sinal de mês incompleto, não de erro de conta.
  const movRegistrado = (mv ? mv.movs : [])
    .filter(m => ST.decisoes[`${ST.mes}|${m.chave}`])
    .reduce((s, m) => s + (m.delta_carla || 0), 0);

  const lib = d.liberado;
  const linhas = [
    ['Valor inicial + movimentações = valor final (Carla)', iniC + movC, fimC],
    ['Valor inicial + movimentações = valor final (Carla + Eduardo)', iniCE + movCE, fimCE],
  ];
  // mês fechado nunca passou pelo registro manual: cobrar isso ali é acusar erro
  // onde não há. O par só entra como check no mês aberto.
  if (!mesFechado(ST.mes)) linhas.push(
    ['Movimentação registrada = movimentação apurada (Carla)', movRegistrado, movC]);
  /* Dinheiro que cai numa causa sem linha no resumo (renome, sem efeito) não
     aparece em lugar nenhum: o saldo final muda e nenhuma linha explica. */
  const causasDoResumo = new Set(LINHAS_RESUMO.flatMap(([, cs]) => cs || []));
  const movEmCausa = (mv ? mv.movs : []).reduce((s, m) => {
    const dec = ST.decisoes[`${ST.mes}|${m.chave}`];
    return s + (causasDoResumo.has(dec ? dec.classe : m.sugestao) ? (m.delta_carla || 0) : 0);
  }, 0);
  linhas.push(['Causas do resumo = movimentação apurada (Carla)', movEmCausa, movC]);
  /* Confronto com o que foi divulgado. Vem junto com o mês (aba Resumo Contabil
     do mapa) porque é o número que valeu, e não se reproduz de trás pra frente.
     De mar/26 a jul/26 bate em R$ 0. Os dois meses que não batem são da fonte:
       jan/26  -257.425 = -8.925 (base dez/25) -244.000 (atribuição: 9 animais
               que o mapa de dez/25 conta como da Carla e o de jan/26 não —
               diferença entre os dois mapas, não movimentação do mês)
               -9.000 (reavaliação) -3.000 (compra) +7.500 (embrião);
       fev/26   -1.500 = OASIS DA PAO GRANDE, cuja ocorrência está datada
               09/nov/2026 no arquivo de fevereiro e 09/fev/2026 no de março —
               pela data do arquivo de fevereiro ela é posterior ao mês. */
  if (lib) {
    if (lib.saldo_fim != null) linhas.push(['Saldo final = Resumo Contábil divulgado', fimC, lib.saldo_fim]);
    if (lib.saldo_ini != null) linhas.push(['Saldo inicial = Resumo Contábil divulgado', iniC, lib.saldo_ini]);
  }
  const regs = (mv ? mv.movs : []).filter(m => m.no_escopo || m.dono);
  return `<div class="rolagem"><table class="t">
    <thead><tr><th class="l">Check de ${rotMes(ST.mes)}</th><th>Apurado</th><th>Esperado</th><th>Diferença</th><th class="l">Situação</th></tr></thead>
    <tbody>${linhas.map(([t, a, b]) => {
      const dif = +(a - b).toFixed(2);
      // tolerância de R$ 1: o Resumo Contábil divulgado carrega centavos de
      // arredondamento próprio (15.970.552,61 contra 15.970.552,71)
      const ok = Math.abs(dif) < 1;
      return `<tr><td class="l">${t}</td><td>${rs(a)}</td><td>${rs(b)}</td>
        <td class="${ok ? 'pos' : 'neg'}">${rs(dif)}</td>
        <td class="l">${ok ? '<span class="tag ok">confere</span>' : '<span class="tag ruim">diverge</span>'}</td></tr>`;
    }).join('')}
    <tr><td class="l">Movimentações registradas</td><td>${regs.filter(m => ST.decisoes[`${ST.mes}|${m.chave}`]).length}</td>
      <td>${regs.length}</td><td></td>
      <td class="l">${mesFechado(ST.mes) ? '<span class="tag">mês fechado</span>'
        : regs.every(m => ST.decisoes[`${ST.mes}|${m.chave}`]) && regs.length
        ? '<span class="tag ok">mês completo</span>' : '<span class="tag ruim">falta registrar</span>'}</td></tr>
    </tbody></table></div>`;
}

function painelResumo(){
  const r = resumoAno();
  const meses = Object.keys(r).sort();
  if (!meses.length) return `<div class="aviso">Importe pelo menos dois meses para o resumo contábil.</div>`;
  const val = (m, causas) => {
    if (!causas) return null;
    return causas.reduce((s, c) => s + (r[m].causas[c] || 0), 0);
  };
  return `<div class="rolagem"><table class="t">
    <thead><tr><th class="l">Título</th>${meses.map(m => `<th>${rotMes(m)}</th>`).join('')}<th>Ano</th></tr></thead>
    <tbody>${LINHAS_RESUMO.map(([rot, causas]) => {
      if (rot === 'Saldo inicial') return `<tr class="tot"><td>${rot}</td>${meses.map(m => `<td>${rs(r[m].ini)}</td>`).join('')}<td>${rs(r[meses[0]].ini)}</td></tr>`;
      if (rot === 'Saldo final') return `<tr class="tot"><td>${rot}</td>${meses.map(m => `<td>${rs(r[m].fim)}</td>`).join('')}<td>${rs(r[meses[meses.length - 1]].fim)}</td></tr>`;
      const vals = meses.map(m => val(m, causas));
      const soma = vals.reduce((a, b) => a + b, 0);
      return `<tr><td>${rot}</td>${vals.map(v => `<td class="${clsN(v)}">${v ? rs(v) : '—'}</td>`).join('')}<td class="${clsN(soma)}">${rs(soma)}</td></tr>`;
    }).join('')}
    <tr><td>Movimentações registradas</td>${meses.map(m => `<td>${r[m].registrado}/${r[m].total}</td>`).join('')}<td></td></tr>
    </tbody></table></div>`;
}

const semArquivo = () => `<div class="aviso">Nenhum arquivo importado. Use <b>Importar arquivo</b> e escolha o
  plantel enviado pelo haras (ou o mapa de movimentação, que traz a atribuição Carla/Eduardo).</div>`;

/* ---------- casca ---------- */
const ABAS = [['plantel', 'Plantel'], ['mov', 'Movimentações'], ['resumo', 'Resumo contábil']];

function pinta(){
  document.getElementById('abas').innerHTML = ABAS.map(([id, t]) =>
    `<button type="button" data-aba="${id}" class="${ST.aba === id ? 'on' : ''}">${t}</button>`).join('');
  const el = document.getElementById('painel');
  el.innerHTML = ST.aba === 'plantel' ? painelPlantel() : ST.aba === 'mov' ? painelMov() : painelResumo();
  posPop();
}

function liga(){
  document.body.addEventListener('click', async e => {
    const aba = e.target.closest('[data-aba]');
    if (aba) { ST.aba = aba.dataset.aba; ST.pop = null; pintaPop(); pinta(); return; }
    const sub = e.target.closest('[data-sub]');
    if (sub) { ST.sub = sub.dataset.sub; ST.pop = null; pintaPop(); pinta(); return; }
    const fb = e.target.closest('[data-fb]');
    if (fb) {
      const [qual, ci] = fb.dataset.fb.split(':');
      const col = +ci;
      ST.pop = (ST.pop && ST.pop.qual === qual && ST.pop.col === col)
        ? null : {qual, col, busca: '', foco: true};
      pintaPop(); return;
    }
    const fx = e.target.closest('[data-fx]');
    if (fx) {
      const [qual, ci] = fx.dataset.fx.split(':');
      delete ST.filtros[qual][+ci];
      ST.pop = null; pintaPop(); pinta(); return;
    }
    const fz = e.target.closest('[data-fzerar]');
    if (fz) {
      ST.filtros[fz.dataset.fzerar] = {};
      ST.pop = null; pintaPop(); pinta(); return;
    }
    const ord = e.target.closest('[data-ord]');
    if (ord) {
      const [qual, ci] = ord.dataset.ord.split(':');
      const i = +ci, atual = ST.ordem[qual];
      ST.ordem[qual] = {col: i, dir: atual.col === i ? -atual.dir : 1};
      pinta(); return;
    }
    const fa = e.target.closest('[data-fpop]');
    if (fa && ST.pop) {
      const {qual, col} = ST.pop;
      if (fa.dataset.fpop === 'fecha') { ST.pop = null; pintaPop(); return; }
      if (fa.dataset.fpop === 'nenhum') delete ST.filtros[qual][col];
      else {
        // 'marcar' vale só pro que está visível na busca, que é o que se vê
        const ctx = ST.ctx[qual], b = norm(ST.pop.busca || '');
        const opts = opcoesDe(ctx.base, qual, ctx.txtDe, ctx.ordDe, col)
          .map(([t]) => t).filter(t => !b || norm(t).includes(b));
        ST.filtros[qual][col] = opts;
      }
      ST.pop.foco = false;
      pinta(); pintaPop(); return;
    }
    // clique fora fecha o menu — sem engolir o clique, que pode ser de outro botão
    if (ST.pop && !e.target.closest('#fpop')) { ST.pop = null; pintaPop(); }
    const dono = e.target.closest('[data-dono]');
    if (dono) {
      const [k, dn] = dono.dataset.dono.split(':');
      await atribui(ST.mes, k, dn);
      pinta(); return;
    }
  });
  // a busca só encurta a lista do menu; a tabela não é redesenhada a cada tecla
  document.body.addEventListener('input', e => {
    if (e.target.id === 'fpopBusca' && ST.pop) {
      ST.pop.busca = e.target.value;
      ST.pop.foco = true;
      pintaPop();
    }
  });
  document.body.addEventListener('change', async e => {
    if (e.target.dataset.fv != null && ST.pop) {
      marcaValor(ST.pop.qual, ST.pop.col, e.target.dataset.fv, e.target.checked);
      ST.pop.foco = false;
      pinta(); pintaPop(); return;
    }
    const reg = e.target.closest('[data-reg]');
    if (!reg) return;
    const mv = movimentacaoDoMes(ST.mes);
    const mov = mv.movs.find(m => m.chave === reg.dataset.reg);
    if (!mov || !reg.value) return;
    await registra(ST.mes, mov, reg.value, '');
    pinta();
  });
}

(async function boot(){
  await carregaSnapshots();
  const meses = Object.keys(ST.meses).sort();
  ST.mes = meses[meses.length - 1] || null;
  topo(); liga(); pinta();
})();
