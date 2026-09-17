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
/* Mês aberto ou fechado. A verdade mora na tabela `plantel_mes_status` e é a
   MESMA que o banco consulta em `plantel_mes_fechado` — a tela não decide nada
   sozinha, só evita oferecer o que o RLS vai recusar.

   `FECHADO_ATE` sobrou como fallback do histórico: os meses anteriores a esta
   tela foram fechados fora dela, e criar linha pra eles seria inventar autor e
   data. Sem linha na tabela, vale o corte antigo. */
const FECHADO_ATE = '2026-07';
const mesFechado = m => !!m && (m in ST.statusMes ? ST.statusMes[m] : m <= FECHADO_ATE);
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
/* RECEPTORAS não é animal: é UMA linha agregada, e o número no nome é a
   CONTAGEM ("RECEPTORAS 166" em fev/26 vale 344.847,63; "RECEPTORAS 163" em
   mar/26 vale 338.460,65). Tratando o nome como identidade, a troca de número
   virava um animal saindo e outro entrando — em março, uma venda de R$ 344.848
   e uma compra de R$ 338.461, quando o que houve foram 3 receptoras a menos,
   R$ 6.387. A ponte por nome não salva: o log diz "ESTAVA RECEPTORAS 161",
   número que não bate com a linha do mês anterior. Chave fixa resolve. */
const RX_RECEPTORAS = /^RECEPTORAS(\s+\d+)?$/;
const ehLinhaReceptoras = nome => RX_RECEPTORAS.test(norm(nome));
const chaveCom = (l, ix) => ehLinhaReceptoras(l[ix.nome])
  ? 'PSEUDO:RECEPTORAS'
  : norm(l[ix.nome]) + '|' + norm(l[ix.letra]);
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
  /* Não existe mais estado de atribuição: o dono de cada linha sai do próprio
     arquivo (ver donoDaLinha). O que havia aqui — atrib por mês, inferido por
     sufixo e a fila 'sugeridos' de confirmação — foi removido em 16/09/2026. */
  mes: null,
  disponiveis: [],  // meses que existem no bucket/tabela, sem estarem carregados
  statusMes: {},    // 'AAAA-MM' -> true quando FECHADO (ausente = corte histórico)
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

/* Espaço sobrando é sujeira, não informação: o arquivo do haras é digitado à
   mão e jul/26 tem 'DA PAO GRANDE' (292 linhas) convivendo com 'DA PAO GRANDE '
   (15, espaço no fim). Como valor de texto, os dois são diferentes: viravam duas
   entradas na lista do filtro, e marcar uma deixava 15 animais de fora sem
   avisar. Limpar na ENTRADA resolve pra todo mundo de uma vez — filtro, chave do
   animal, agrupamento — em vez de cada consumidor lembrar de aparar. Data e
   número passam intactos. */
const limpaLinha = r => r.map(c => typeof c === 'string' ? c.replace(/\s+/g, ' ').trim() : c);

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
  const linhas = tudo.slice(iCab + 1).filter(r => r && r[ix.nome]).map(limpaLinha);

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
/* O mapa do Luxor NÃO é mais fonte de atribuição. Ele era, e em volta disso
   existia uma máquina inteira: gravar dono por animal e por mês, herdar do mês
   anterior, inferir por sufixo quando o mapa não cobria, e uma fila de
   'confirmar dono' na Conciliação. Tudo isso saiu em 16/09/2026 — o dono está
   NO ARQUIVO (ver donoDaLinha), e a regra reproduz a Carla de jul/26 no
   centavo. O mapa segue sendo importado só pelo Resumo Contábil divulgado, que
   é o número contra o qual os checks conferem. */

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
/* De quem é a LINHA, lido do próprio arquivo do haras. Regra dada pelo Arthur
   em 16/09/2026, e ela dispensa mapa, coluna 'Dono' na tela e classificação
   manual — tudo que eu tinha montado em volta disso era rodeio:

     NOME com '(CARLA)' ou '(EDUARDO)'   manda. É o caso dos dois donos no MESMO
                                         animal, que o haras lança como duas
                                         linhas (LATINO DA PAO GRANDE (CARLA) e
                                         LATINO DA PAO GRANDE (EDUARDO)).
     SUFIXO 'DA PAO GRANDE' ou 'OUTRO'   Carla.
     qualquer outro sufixo ('- E xx%')   Eduardo.

   Conferido contra jul/2026: pela regra a Carla soma R$ 15.970.553, que é
   exatamente o saldo do Resumo Contábil divulgado — o mesmo número que o mapa
   dava, sem precisar do mapa. */
function donoDaLinha(l, ix){
  ix = ix || (ST.meses[ST.mes] && ST.meses[ST.mes].ix) || {};
  const nome = norm(l[ix.nome]);
  if (nome.includes('(CARLA)')) return 'hpg';
  if (nome.includes('(EDUARDO)')) return 'eduardo';
  const suf = norm(l[ix.sufixo]);
  return (suf === 'DA PAO GRANDE' || suf === 'OUTRO') ? 'hpg' : 'eduardo';
}

function patr(l, escopo, ix, mes){
  ix = ix || (ST.meses[ST.mes] && ST.meses[ST.mes].ix) || {};
  const a = donoDaLinha(l, ix);
  if (escopo === 'hpg' && a !== 'hpg') return 0;
  if (escopo === 'eduardo' && a !== 'eduardo') return 0;
  /* 'carla_eduardo' não filtra nada de propósito: pela regra do sufixo TODA
     linha do arquivo é de um dos dois, então a soma dos dois é o arquivo
     inteiro. Antes existia um terceiro estado ('nenhum'), que era artefato de
     animal que o mapa não cobria — sem mapa, ele deixa de existir. */
  return num(l[ix.cota]) * num(l[ix.valor]) + num(l[ix.comissao]);
}
/* Soma das cotas do mês, na mesma base das linhas efetivas que o patrimônio
   usa — senão o check compararia populações diferentes. */
function cotasMes(mes){
  const d = ST.meses[mes];
  if (!d) return 0;
  return linhasEfetivas(mes).reduce((s, l) => s + num(l[d.ix.cota]), 0);
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
/* Nome antigo dentro da ocorrência de renome. O haras escreve de duas formas —
   "MUDOU O NOME - ESTAVA X" e "MUDOU DE NOME - ERA X E FICOU Y" — e a regra só
   conhecia a primeira. As 8 ocorrências com "ERA" viravam DOIS animais: o de
   nome velho sumindo e o de nome novo aparecendo. Em jun/26 isso fabricou uma
   venda de −R$ 45.000 e uma compra de +R$ 45.000 da PERSIA DA PAO GRANDE, que
   só trocou de nome. Medido contra os 156 renomes de jan a jul/2026: a regra
   antiga resolvia 148, esta resolve os 156. */
const RX_ESTAVA = /(?:ESTAVA|ERA)\s+(?:COMO\s+)?"?(.+?)"?(?=\s+-\s+(?:MUDOU|TINHA|FOI|PASSOU|ADICAO|ERA|E\s)|\s+(?:PASSOU|E\s+FICOU|E\s+PASSOU|FICOU)\s|$)/;

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
  // com o dono vindo do arquivo, TODA linha é de um dos dois — não existe mais
  // animal fora de escopo, que era artefato de mapa incompleto
  const noEscopo = () => true;

  const log = (d.log || []).filter(x => {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}` === mes;
  }).map(x => ({...x, tipo: tipoLog(x.ocorrencia)}));

  /* O log é agrupado pela MESMA identidade das linhas: sem isto, a ocorrência
     de junho ficava sob "RECEPTORAS 124" (nome antigo) e a linha do mês era
     "RECEPTORAS 121", então o motor não achava o log e não via o DEVOLUÇÃO. */
  const chaveLog = produto => ehLinhaReceptoras(produto) ? 'PSEUDO:RECEPTORAS' : norm(produto);
  const logPorNome = {};
  for (const x of log) (logPorNome[chaveLog(x.produto)] = logPorNome[chaveLog(x.produto)] || []).push(x);
  /* O log do arquivo é cumulativo: traz a vida inteira do animal, não só o mês.
     A tela usa o recorte do mês, mas a regra de cancelamento precisa do
     histórico (a venda cancelada pode ser de anos atrás). */
  const histPorNome = {};
  for (const x of (d.log || [])) {
    const k = chaveLog(x.produto);
    (histPorNome[k] = histPorNome[k] || []).push(
      {...x, data: x.data instanceof Date ? x.data : new Date(x.data)});
  }

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
    const trocouDono = !!(a && b && donoDaLinha(a, ixa) !== donoDaLinha(b, ixb));
    if (Math.abs(p1 - p0) < 0.01 && !ren && !mudouStatus && !mudouLocal && !trocouDono && a && b) continue;

    const linha = b || a, ixL = b ? ixb : ixa;
    const dono = donoDaLinha(linha, ixL);
    const donoAnt = a ? donoDaLinha(a, ixa) : dono;
    /* Trocar de dono MOVE patrimônio: animal que era da Carla e virou 'nenhum'
       sai do saldo dela mesmo sem venda registrada. Sem isto o check de julho
       acusava R$ 60 mil de diferença entre estoque e fluxo. */
    const noEsc = (dn, v) => dn === 'hpg' ? v : 0;
    const emCE = dn => ['hpg', 'eduardo'].includes(dn);
    const deltaCarla = +(noEsc(dono, p1) - noEsc(donoAnt, p0)).toFixed(2);
    const deltaCE = +((emCE(dono) ? p1 : 0) - (emCE(donoAnt) ? p0 : 0)).toFixed(2);
    const itensLog = logPorNome[chaveLog(linha[ixL.nome])]
      || (ren ? logPorNome[chaveLog(ren.para)] : []) || [];
    const comum = {
      chave: k, linha, nome: linha[ixL.nome], sufixo: linha[ixL.sufixo],
      categoria: linha[ixL.categoria], status: linha[ixL.status], dono,
      cota_ant: q0, cota_atual: q1, valor_ant: v0, valor_atual: v1,
      patr_ant: p0, patr_atual: p1,
      entrou: !a, saiu: !b, renome: ren || null,
      dono_ant: donoAnt,
      mudou_dono: donoAnt !== dono ? [donoAnt, dono] : null,
      mudou_status: mudouStatus ? [a[ixa.status], b[ixb.status]] : null,
      mudou_local: mudouLocal ? [a[ixa.local], b[ixb.local]] : null,
      posterior, no_escopo: noEscopo(linha),
    };
    /* proporção do delta que cabe a cada escopo, para repartir evento a evento
       sem recalcular dono */
    const fatiaCarla = (p1 - p0) ? deltaCarla / (p1 - p0) : (dono === 'hpg' ? 1 : 0);
    const fatiaCE = (p1 - p0) ? deltaCE / (p1 - p0) : (emCE(dono) ? 1 : 0);
    const empurra = (chave, delta, sugestao, log, rotulo) => movs.push({
      ...comum, chave, delta: +delta.toFixed(2),
      delta_carla: +(delta * fatiaCarla).toFixed(2),
      delta_ce: +(delta * fatiaCE).toFixed(2),
      sugestao, log, evento: rotulo || null,
    });

    /* Um movimento por EVENTO do log, quando o log explica a variação inteira.
       Sem isso o mês virava um número líquido só: o PLATAO de abril, que teve
       compra de 25% por R$ 40K e venda de 5%, aparecia como uma compra de
       R$ 10.000 — e a Controladoria lança os dois eventos e a reavaliação da
       diferença entre preço e valor de cota. A soma dos eventos é sempre o
       delta do mês; a decomposição reparte, não cria. */
    const evs = eventosDoMes(itensLog, q0, q1, v1 || v0, b ? num(b[ixb.comissao]) : 0);
    if (evs && evs.length) {
      let somado = 0;
      evs.forEach((e, i) => {
        const oc = norm(e.oc);
        const compra = e.dCota > 0;
        const cls = compra ? 'compra'
          : /DOAD|DOACAO|DEVOLU/.test(oc) ? 'doacao'
          : /MORREU|OBITO|ABORTOU/.test(oc) ? 'morte'
          : 'venda';
        const umLog = [{data: e.data, ocorrencia: e.texto}];
        if (compra && e.preco) {
          /* compra entra pelo PREÇO PAGO; a diferença contra o valor da fatia é
             reavaliação — 40.000 pagos por 12.500 de cota são compra 40.000 e
             reavaliação −27.500, que somam os mesmos +12.500 do patrimônio */
          empurra(`${k}#${i + 1}`, e.preco, 'compra', umLog, 'compra pelo preço pago');
          somado += e.preco;
          const ajuste = e.valorFatia - e.preco;
          if (Math.abs(ajuste) >= 1) {
            empurra(`${k}#${i + 1}r`, ajuste, 'reavaliacao', umLog,
                    'diferença entre o preço pago e o valor de cota');
            somado += ajuste;
          }
        } else {
          empurra(`${k}#${i + 1}`, e.valorFatia, cls, umLog, null);
          somado += e.valorFatia;
        }
      });
      // o que os eventos não explicam (valor unitário que mudou no mesmo mês)
      const resto = (p1 - p0) - somado;
      if (Math.abs(resto) >= 1) empurra(`${k}#resto`, resto, 'reavaliacao', itensLog,
                                        'variação de valor não explicada pelos eventos');
    } else {
      empurra(k, p1 - p0,
        sugere({q0, q1, v0, v1, p0, p1, ren, entrou: !a, saiu: !b, status: stB || stA,
                categoria: norm(linha[ixL.categoria]), itensLog,
                historico: histPorNome[chaveLog(linha[ixL.nome])] || [],
                receptoras: ehLinhaReceptoras(linha[ixL.nome])}),
        itensLog, null);
    }
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

/* Cancelamento de venda: o que ele é depende de QUANDO a venda foi.

   Regra do Arthur (17/09/2026): venda cancelada no MESMO ano é estorno da
   venda — as duas pernas se anulam dentro do exercício. Em ano DIFERENTE é
   recompra: o exercício anterior já fechou com a cota reduzida, então a cota
   voltando agora é entrada nova neste ano.

   Dá pra decidir porque o log do haras é CUMULATIVO — carrega a vida inteira do
   animal, não só o mês. NIOBIO DA PAO GRANDE, no arquivo de fevereiro:
     2025-10-17  VENDIDO 25% PARA VITOR BEZERRA ... COTA DE 75 PARA 25%
     2026-02-11  VENDA DE 25% PARA VITOR BEZERRA CANCELADA - DE 50 PARA 75%
   Anos diferentes, logo compra. Era uma das compras de fev/26 que o resumo
   original tinha e o nosso não via. */
const RX_VENDA_LOG = /VEND[AI]/;

function anoDaVendaOriginal(historico, quando){
  if (!historico || !historico.length || !quando) return null;
  const t = quando instanceof Date ? quando.getTime() : new Date(quando).getTime();
  let achado = null;
  for (const x of historico) {
    const d = x.data instanceof Date ? x.data : new Date(x.data);
    if (!(d.getTime() < t)) continue;                       // só o que veio antes
    const s = norm(x.ocorrencia);
    if (!RX_VENDA_LOG.test(s) || /CANCEL/.test(s)) continue;  // venda, e não outro cancelamento
    if (!achado || d > achado) achado = d;                  // a mais recente antes do cancelamento
  }
  return achado ? achado.getFullYear() : null;
}

function cancelamento(oc){
  const s = norm(oc);
  if (!/CANCEL/.test(s)) return null;
  if (RX_CANCEL_REVENDA.test(s)) return {tipo: 'venda', sentido: 'sai', revenda: true};
  if (RX_CANCEL_TERCEIRO.test(s)) return {tipo: 'venda', sentido: 'volta'};
  if (RX_CANCEL_COMPRA.test(s)) return {tipo: 'compra', sentido: 'sai'};
  if (RX_CANCEL_VENDA.test(s)) return {tipo: 'venda', sentido: 'volta'};
  return {tipo: null, sentido: null};
}

/* ---------- decomposição do mês em EVENTOS ----------

   O motor fazia UMA movimentação por animal por mês, com o delta líquido, e
   classificava esse líquido uma vez só. A Controladoria lança evento por
   evento, e é por isso que o PLATAO DA PAO GRANDE de abril não fechava:

     log: COMPRADO 25% DE ALEXANDRE DE MORAIS POR R$40K - COTA 46,25 -> 71,25%
          VENDIDO 5% PARA ANTONIO AZEVEDO - COTA 71,25 -> 66,25%
     mapa : compra 40.000 · venda -2.500 · reavaliação -27.500
     motor: compra 10.000            <- só o líquido dos dois

   Duas regras saem daí, e as duas estão escritas no próprio log:

   1. cada ocorrência que move a cota é um evento, com o valor da sua fatia
      (Δcota × valor unitário);
   2. compra com PREÇO declarado entra pelo preço, e a diferença entre o preço
      e o valor da fatia vira reavaliação — 40.000 pagos por 12.500 de cota são
      compra 40.000 e reavaliação −27.500, que somam os mesmos +12.500 do
      patrimônio. A decomposição nunca muda o total do mês; só reparte.

   Quando o log não explica a variação inteira, nada disso se aplica: a
   movimentação continua inteira, como antes. Explicar pela metade seria pior
   que não explicar. */

const RX_COTA_DE_PARA = /(?:MUDOU\s+)?(?:A\s+)?(?:%|COTA|PORCENTAGEM)\s*(?:DE|ERA)\s*([\d,.]+)\s*%?\s*(?:PARA|P\/|PASSOU\s+PARA|->)\s*([\d,.]+)\s*%/;
const RX_PRECO = /POR\s*R?\$?\s*([\d.,]+)\s*(K|MIL)?\b/;
const RX_PCT_EVENTO = /\b(VENDID[AO]|VENDA|COMPRAD[AO]|COMPRA|DOAD[AO]|DOACAO)\b[^.]{0,40}?([\d,.]+)\s*%/;

const _num = t => Number(String(t).replace(/\./g, '').replace(',', '.'));

/* Preço declarado na ocorrência: "POR R$40K", "POR R$ 40.000", "POR 40 MIL". */
function precoDeclarado(oc){
  const m = RX_PRECO.exec(oc);
  if (!m) return null;
  const v = _num(m[1]);
  if (!isFinite(v) || !v) return null;
  return m[2] ? v * 1000 : v;
}

/* Transição de cota da ocorrência, em fração (0,4625 -> 0,7125). */
function cotaDaOcorrencia(oc){
  const m = RX_COTA_DE_PARA.exec(oc);
  if (!m) return null;
  const de = _num(m[1]) / 100, para = _num(m[2]) / 100;
  if (!isFinite(de) || !isFinite(para)) return null;
  return {de, para};
}

/* Quebra o mês de UM animal em eventos. Devolve null quando o log não explica a
   variação de cota inteira — aí o chamador mantém a movimentação única. */
function eventosDoMes(itensLog, q0, q1, valor, comissao){
  if (!itensLog || !itensLog.length) return null;
  const passos = [];
  for (const it of itensLog) {
    const oc = norm(it.ocorrencia);
    const c = cotaDaOcorrencia(oc);
    if (!c) continue;
    passos.push({oc, de: c.de, para: c.para, preco: precoDeclarado(oc), data: it.data,
                 texto: it.ocorrencia});
  }
  if (!passos.length) return null;
  // a corrente tem de começar na cota do mês anterior e terminar na do mês
  const tol = 0.0005;
  if (Math.abs(passos[0].de - q0) > tol) return null;
  if (Math.abs(passos[passos.length - 1].para - q1) > tol) return null;
  for (let i = 1; i < passos.length; i++) {
    if (Math.abs(passos[i].de - passos[i - 1].para) > tol) return null;   // corrente quebrada
  }
  /* A comissão acompanha a cota proporcionalmente: ela é parte do patrimônio
     (cota × valor + comissão) e não tem evento próprio no log. */
  const porCota = q0 ? comissao / q0 : 0;
  return passos.map(p => {
    const dCota = p.para - p.de;
    const valorFatia = dCota * valor + dCota * porCota;
    return {...p, dCota, valorFatia};
  });
}

/* sugestão pelo padrão — só sugestão; o registro é o input de quem fecha.
   Dinheiro manda na ordem: renome é tipo 1-Nome e não move saldo, então só
   ganha a sugestão quando o patrimônio ficou igual. E patrimônio é
   cota × valor + comissão — olhar só cota e valor deixava mudança de comissão
   como 'sem_efeito'. Juntas, as duas falhas jogavam R$ 241 mil de janeiro/26
   em causas que o resumo não tem linha pra mostrar. */
function sugere({q0, q1, v0, v1, p0, p1, ren, entrou, saiu, status, categoria, itensLog,
                 historico, receptoras}){
  const st = status || '', cat = categoria || '';
  const oc = norm((itensLog || []).map(x => x.ocorrencia).join(' '));
  const mexeu = Math.abs((p1 || 0) - (p0 || 0)) >= 0.01;
  /* Animal que troca de dono entre um mapa e outro SEM mudar cota nem valor
     não é evento contábil: o fechamento nunca teve essa causa. O que existe é
     diferença de atribuição entre os dois mapas (jan/26: 9 animais que o mapa
     de dez/25 conta como da Carla e o de jan/26 não), e isso se resolve no
     mapa, não virando linha no resumo. Cai em 'sem_efeito' e aparece na coluna
     'O que foi feito' como troca de dono. */
  /* Receptora que ENTRA é compra (o haras informa o valor pago). Receptora que
     SAI pode ser venda ou baixa, e a direção do valor não distingue as duas — em
     jun/26 saíram 3 e a contabilidade lançou os R$ 6.223 em "Baixa mortes e
     doações", porque o log diz DEVOLUÇÃO, não venda. Então quem decide é o log;
     não dizendo nada, o motor não chuta. O valor já vem proporcional: a linha
     agregada carrega o valor do conjunto, então o delta É a fatia das que saíram. */
  if (receptoras) {
    /* O SINAL decide primeiro. Lendo a palavra antes do sinal, março virava
       "compra de -R$ 6.387": o log da linha agregada acumula os eventos do mês
       ("COMPRA DE 02" e devolução juntos) e a palavra COMPRA aparecia num mês
       cujo saldo foi negativo. Entrou valor é compra — o haras informa o preço
       pago. Saiu, o log distingue; devolução e venda caem na MESMA linha de
       baixa por venda (Arthur, 17/09/2026). Saiu sem o log dizer o quê, não
       chuta: vira pergunta na ficha. */
    if (p1 > p0) return 'compra';
    return /DEVOLU|VEND/.test(oc) ? 'venda' : null;
  }
  if (ren && !mexeu) return 'renome';
  if (entrou) return /EMBRI/.test(cat) ? 'embriao' : /NASCEU/.test(oc) ? 'embriao' : 'compra';
  if (saiu || (q0 && !q1)) {
    /* O óbito continua ganhando da venda quando os dois aparecem no mês, e isso
       foi MEDIDO: a aba Movimentações do mapa lança o JUSTICA DA PAO GRANDE X
       OLIMPO DO MH (vendido 50% em 13/03, abortou em 31/03) como venda, mas o
       Resumo Contábil de março quer aquele valor em morte — e é o Resumo que
       vale, porque é ele que tem mês. Inverter a ordem levou a divergência de
       março de R$ 22.250 para R$ 58.250. As duas abas do mapa discordam nesse
       animal; a classificação final é de quem fecha, pela ficha. */
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
  if (canc && canc.sentido === 'volta') {
    /* a cota volta: estorno da venda se ela foi neste ano, compra se foi em ano
       anterior (ver anoDaVendaOriginal) */
    const item = (itensLog || []).find(x => /CANCEL/.test(norm(x.ocorrencia)));
    const anoCanc = item && (item.data instanceof Date ? item.data : new Date(item.data)).getFullYear();
    const anoVenda = anoDaVendaOriginal(historico, item && item.data);
    return (anoVenda && anoCanc && anoVenda !== anoCanc) ? 'compra' : 'venda';
  }
  if (canc && canc.tipo === 'compra' && canc.sentido === 'sai') return 'compra';
  if (q1 > q0) return 'compra';
  if (q1 < q0) return /MORREU|OBITO/.test(oc) ? 'morte' : /DOAD/.test(oc) ? 'doacao' : 'venda';
  /* Aqui acabavam as regras e o resto virava 'reavaliacao'. Isso não é sugestão,
     é invenção: o Resumo Contábil divulgado tem ZERO reavaliação em fev, mai,
     jun e jul/2026, e nesses meses o motor fabricava exatamente o valor do que
     era outra coisa (jun: os R$ 6.223 que a contabilidade lançou como baixa).
     Reavaliação, quando acontece, é evento próprio e sai na aba REAV. PLANTEL —
     não se deduz de sobra. Sem regra que explique, o motor não sugere: a ficha
     pergunta sem palpite e a movimentação fica FORA do resumo até alguém
     classificar, em vez de entrar calada numa causa errada. */
  /* Reavaliação tem assinatura própria: a COTA fica igual e o VALOR UNITÁRIO
     muda — o animal é o mesmo, a participação é a mesma, só o preço mudou.
     Medido em jan/26: LIBERDADE DA PAO GRANDE X IMPERIO SAPECADO, cota 0,5 nos
     dois meses e valor de 30.000 para 24.000, são os R$ 18.000 que o Resumo
     Contábil publica como "(+/−) Reavaliações". Isto é diferente do que havia
     aqui antes, que era "qualquer sobra vira reavaliação" — aquilo inventava
     causa; isto reconhece uma. As receptoras não passam por aqui: têm regra
     própria acima, senão o valor do agregado mudando viraria reavaliação. */
  if (q1 === q0 && v1 !== v0) return 'reavaliacao';
  if (mexeu) return null;
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

/* YTD até o mês ESCOLHIDO, não até onde houver arquivo. O seletor de mês vale
   pra aba inteira: com mar/26 selecionado o resumo mostrava jan a ago e uma
   coluna ANO fechando em agosto — o acumulado não tem como ser de um período
   que quem está olhando não pediu. */
function resumoAno(){
  const meses = Object.keys(ST.meses).filter(m => !ST.mes || m <= ST.mes).sort();
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
      // sem decisão e sem sugestão: não entra em causa nenhuma até ser classificada
      if (!classe) continue;
      causas[classe] = +((causas[classe] || 0) + mo.delta_carla).toFixed(2);
    }
    /* lançamento manual entra na causa dele como qualquer outro: é dinheiro que
       aconteceu e que o diff dos arquivos não tinha como ver */
    for (const x of manuaisDoMes(m)) {
      causas[x.classe] = +((causas[x.classe] || 0) + x.valor).toFixed(2);
    }
    out[m] = {ini: +ini.toFixed(2), fim: +fim.toFixed(2), causas,
              registrado: (mv ? mv.movs : []).filter(mo => ST.decisoes[`${m}|${mo.chave}`]).length,
              total: (mv ? mv.movs : []).length};
  }
  return out;
}

/* ================= UI ================= */
function topo(){
  const meses = ST.disponiveis.length ? ST.disponiveis : Object.keys(ST.meses).sort();
  // sem <h1> aqui: o cabeçalho do hub já mostra o nome da aba, e repetir era
  // exatamente o tipo de poluição que o painel não precisa
  /* Aberto/fechado fica AQUI, colado no seletor: é propriedade do mês, não da
     Conciliação. Estava lá dentro e simplesmente não era encontrado. */
  const fechado = mesFechado(ST.mes);
  document.getElementById('topo').innerHTML = `
    <div class="ferramentas">
      <label class="mes-sel">Mês
        <select id="selMes">${meses.map(m => `<option value="${m}"${m === ST.mes ? ' selected' : ''}>${rotMes(m)}</option>`).join('')}</select>
      </label>
      ${!ST.mes ? '' : `<span class="tag ${fechado ? '' : 'ok'}">${fechado ? 'fechado' : 'aberto'}</span>
      <button type="button" class="botao-acao${fechado ? '' : ' primario'}"
        data-mes-status="${ST.mes}:${fechado ? 'abrir' : 'fechar'}">${
        fechado ? 'Reabrir mês' : 'Fechar mês'}</button>`}
      <label class="botao-arquivo">Importar arquivo
        <input type="file" id="arq" accept=".xlsx,.xlsm" hidden>
      </label>
      <span id="statusImp"></span>
    </div>`;
  document.getElementById('selMes').onchange = async e => {
    ST.mes = e.target.value;
    /* o resumo é acumulado ATÉ o mês escolhido, então trocar de mês nele pede o
       período inteiro; nas outras abas bastam o mês e o anterior */
    if (ST.aba === 'resumo') await garanteMeses(ST.disponiveis.filter(m => m <= ST.mes));
    else await garantePar(ST.mes);
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
      fecha();
      document.getElementById('statusImp').textContent =
        'mapa lido — a atribuição Carla/Eduardo sai do próprio arquivo do haras, '
        + 'pelo sufixo; do mapa só entra o Resumo Contábil divulgado';
      pinta();
    };
    return;
  }
  document.getElementById('impOk').onclick = async () => {
    const mes = document.getElementById('impMes').value;
    ST.meses[mes] = d;
    ST.mes = mes;
    // mês novo entra na lista do seletor sem precisar recarregar a aba
    if (!ST.disponiveis.includes(mes)) ST.disponiveis = [...ST.disponiveis, mes].sort();
    fecha();
    await salvaSnapshot(mes, d, arquivo);
    topo(); pinta();
  };
}

async function salvaSnapshot(mes, d, arquivo){
  const c = sb();
  if (!c) return;
  const pacote = {
    mes, arquivo,
    linhas: {cab: d.cab, ix: d.ix, rows: d.linhas},
    log: d.log.map(x => ({...x, data: (x.data instanceof Date ? x.data : new Date(x.data)).toISOString()})),
  };
  // 1) tabela: registro durável do fechamento
  try {
    await c.from('plantel_snapshot').upsert({...pacote, atribuicao: {}}, {onConflict: 'mes'});
  } catch (err) { console.warn('[plantel] snapshot não gravado', err.message || err); }
  // 2) arquivo: é o que a aba lê depois. Falhar aqui não perde dado — a leitura
  //    cai pra tabela —, mas tira a velocidade, então avisa alto.
  try {
    const corpo = new Blob([JSON.stringify(pacote)], {type: 'application/json'});
    const { error } = await c.storage.from(balde())
      .upload(arqDoMes(mes), corpo, {upsert: true, contentType: 'application/json'});
    if (error) throw error;
  } catch (err) {
    console.warn('[plantel] arquivo do mês não gravado', err.message || err);
    document.getElementById('statusImp').textContent =
      'mês salvo, mas o arquivo rápido não subiu: ' + (err.message || err);
  }
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

/* ---------- leitura: arquivo no bucket, não consulta no banco ----------

   A aba fazia `select('mes,arquivo,linhas,log')` SEM filtro: trazia todos os
   meses do Postgres, que tinha de ler o JSONB, converter pra texto e streamar
   pela API — sem CDN e com trabalho de banco por byte. As abas rápidas do hub
   (semanal, comitê, e o indicadores no P&C) fazem outra coisa: URL assinada +
   GET de objeto estático servido pela borda (assets/auth.js, loadData).

   Agora é o mesmo caminho: um arquivo por mês, `plantel.<AAAA-MM>.json`, e só
   os meses que a tela precisa. Abrir a aba carrega DOIS — o mês e o anterior,
   porque a movimentação compara os dois. O resto entra sob demanda.

   A tabela continua gravada e serve de fallback: mês que ainda não tem arquivo
   (os que existiam antes desta mudança) é lido dela, sem o usuário perceber. */

const balde = () => { try { return window.parent.HUB_BUCKET || 'hpg-data'; } catch (e) { return 'hpg-data'; } };
const arqDoMes = m => `plantel.${m}.json`;

/* Assina TODOS os meses pedidos numa chamada e baixa em paralelo.

   Era uma `createSignedUrl` por mês, e cada uma é uma ida ao servidor: abrir a
   aba virava lista + 2 assinaturas + 2 downloads, em série. Isso ficou MAIS
   lento que a consulta única de antes, mesmo trazendo um oitavo do dado — o
   custo aqui é número de idas e voltas, não byte. Com `createSignedUrls` (no
   plural) é uma assinatura só para todos. */
async function baixaMeses(lista){
  const c = sb();
  if (!c || !lista.length) return {};
  const out = {};
  try {
    const { data, error } = await c.storage.from(balde())
      .createSignedUrls(lista.map(arqDoMes), 60);
    if (error || !data) return out;
    await Promise.all(data.map(async it => {
      if (!it || it.error || !it.signedUrl) return;
      const m = /plantel\.(\d{4}-\d{2})\.json/.exec(it.path || '');
      if (!m) return;
      /* cache-busting pelo mesmo motivo do porteiro: o CDN na frente do bucket
         serve HIT mesmo com 'no-store' no objeto, e um mês reimportado podia
         ficar preso na borda parecendo cache do navegador. */
      const url = it.signedUrl + (it.signedUrl.includes('?') ? '&' : '?') + 'cb=' + Date.now();
      try {
        const res = await fetch(url, {cache: 'no-store'});
        if (res.ok) out[m[1]] = normalizaSnapshot(await res.json());
      } catch (e) { /* esse mês cai pro fallback da tabela */ }
    }));
  } catch (e) { /* idem */ }
  return out;
}

async function baixaMesDaTabela(mes){
  const c = sb();
  if (!c) return null;
  try {
    const { data } = await c.from('plantel_snapshot')
      .select('mes,arquivo,linhas,log').eq('mes', mes).maybeSingle();
    return data ? normalizaSnapshot(data) : null;
  } catch (e) { return null; }
}

/* Garante que os meses pedidos estão em ST.meses. Arquivo primeiro, tabela
   depois — e o que não existe em lugar nenhum simplesmente não entra. */
async function garanteMeses(lista){
  const faltam = [...new Set(lista)].filter(m => m && !ST.meses[m]);
  if (!faltam.length) return;
  const doBucket = await baixaMeses(faltam);
  Object.assign(ST.meses, doBucket);
  // só quem não tem arquivo vai à tabela (mês antigo, antes do backfill)
  const resto = faltam.filter(m => !ST.meses[m]);
  await Promise.all(resto.map(async m => {
    const d = await baixaMesDaTabela(m);
    if (d) ST.meses[m] = d;
  }));
}

/* O mês e o anterior: é o par que a aba Movimentações precisa pra existir. */
const garantePar = mes => garanteMeses([mes, mesAnterior(mes)]);

/* Quais meses existem, sem baixar nenhum. Vem da listagem do bucket (barata) e
   se completa com a tabela, que ainda tem os meses sem arquivo. */
async function listaMeses(){
  const c = sb();
  if (!c) return [];
  const achados = new Set(Object.keys(ST.meses));
  try {
    const { data } = await c.storage.from(balde()).list('', {limit: 1000});
    for (const o of data || []) {
      const m = /^plantel\.(\d{4}-\d{2})\.json$/.exec(o.name || '');
      if (m) achados.add(m[1]);
    }
  } catch (e) { /* sem listagem, a tabela resolve abaixo */ }
  /* A tabela só é consultada quando o bucket não devolveu nada: depois do
     backfill ela diria exatamente o mesmo, e a consulta extra custava uma ida e
     volta na abertura de toda sessão. */
  if (!achados.size) {
    try {
      const { data } = await c.from('plantel_snapshot').select('mes');
      for (const r of data || []) achados.add(r.mes);
    } catch (e) { /* idem */ }
  }
  return [...achados].sort();
}

async function carregaDecisoes(){
  const c = sb();
  if (!c) return;
  try {
    const { data } = await c.from('plantel_mov_classificacao')
      .select('mes,chave,classe,nota,autor,nome,valor');
    for (const r of data || []) ST.decisoes[`${r.mes}|${r.chave}`] = r;
  } catch (err) { console.warn('[plantel] decisões não carregadas', err.message || err); }
  try {
    const { data } = await c.from('plantel_mes_status').select('mes,fechado');
    for (const r of data || []) ST.statusMes[r.mes] = !!r.fechado;
  } catch (err) { console.warn('[plantel] status dos meses não carregado', err.message || err); }
}

/* Fecha ou reabre o mês. Reabrir NÃO apaga nada: o mês volta a aceitar registro
   com tudo que já foi classificado no lugar — é ajuste, não recomeço (pra
   recomeçar existe o limpar, ao lado). */
async function mudaStatusMes(mes, fechado){
  const c = sb();
  if (c) {
    const { error } = await c.from('plantel_mes_status')
      .upsert({mes, fechado}, {onConflict: 'mes'});
    if (error) { alert('não deu pra mudar o status do mês: ' + error.message); return false; }
  }
  ST.statusMes[mes] = fechado;
  return true;
}

/* ---------- lançamento manual ----------
   Animal que o arquivo do haras ainda não tem (compra recém-fechada, por
   exemplo) não aparece no diff entre dois meses, então o cálculo não o inventa.
   Aqui a pessoa lança: nome, classe e valor. A chave começa com 'MANUAL:' —
   não colide com chave de animal, que é NOME|LETRA. */
const ehManual = chave => String(chave || '').startsWith('MANUAL:');

async function lancaManual(mes, nome, classe, valor, nota){
  const chave = 'MANUAL:' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const k = `${mes}|${chave}`;
  const linha = {mes, chave, classe, nome, valor, nota: nota || null,
                 autor: eu() || '(local)'};
  ST.decisoes[k] = linha;
  const c = sb();
  if (c) {
    const { error } = await c.from('plantel_mov_classificacao')
      .upsert({mes, chave, classe, nome, valor, nota: nota || null}, {onConflict: 'mes,chave'});
    if (error) { delete ST.decisoes[k]; alert('não gravou: ' + error.message); return false; }
  }
  return true;
}

/* Lançamentos manuais do mês, no formato que o resto da tela já entende. */
const manuaisDoMes = mes => Object.entries(ST.decisoes)
  .filter(([k, d]) => k.startsWith(mes + '|') && ehManual(k.slice(mes.length + 1)) && d)
  .map(([k, d]) => ({chave: k.slice(mes.length + 1), nome: d.nome || '(sem nome)',
                     classe: d.classe, valor: num(d.valor), nota: d.nota, autor: d.autor}));

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

/* Apaga TODAS as decisões do mês e recomeça. Existe porque quem roda o
   fechamento pela primeira vez classifica meia dúzia testando, percebe que o
   critério estava errado e quer o mês limpo — reclassificar item por item para
   um valor que também não vale não é recomeçar. O banco só deixa em mês aberto
   (policy pmc_delete). */
async function reiniciaMes(mes){
  const c = sb();
  if (c) {
    const { error } = await c.from('plantel_mov_classificacao').delete().eq('mes', mes);
    if (error) { alert('não deu pra reiniciar: ' + error.message); return false; }
  }
  for (const k of Object.keys(ST.decisoes)) if (k.startsWith(mes + '|')) delete ST.decisoes[k];
  return true;
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
      /* os VALORES, não a contagem: '2 itens' obriga a abrir o menu de novo pra
         lembrar o que está filtrado, que é justamente o que a etiqueta existe
         pra evitar. Lista inteira enquanto couber; passou disso, mostra os
         primeiros e diz quantos sobraram (o title traz todos). */
      let txt = v.join('; ');
      if (txt.length > 64) {
        const cabem = [];
        let n = 0;
        for (const x of v) {
          if (n + String(x).length > 52 && cabem.length) break;
          cabem.push(x); n += String(x).length + 2;
        }
        txt = cabem.join('; ') + ` +${v.length - cabem.length}`;
      }
      return `<span class="chip" title="${esc(v.join('; '))}">`
        + `<b>${esc(ctx.rotulos[ci] || '')}</b><span class="chip-v">(${esc(txt)})</span>`
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
      <span>Eduardo: <b>${rs(somaEsc('eduardo'))}</b></span>
      <span>Carla + Eduardo: <b>${rs(somaEsc('hpg') + somaEsc('eduardo'))}</b></span>
    </div>
    ${chipsFiltro('plantel')}
    <div class="rolagem"><table class="t">
      <thead>
        <tr>${cols.map(([i, r]) => cabFiltro('plantel', i, r, EH_NUM(r), ST.ordem.plantel)).join('')}</tr>
      </thead>
      <tbody>${linhas.map(l => `<tr>${cols.map(([i, r]) => {
          const txt = fmtCel(l, i, r);
          return `<td class="${EH_NUM(r) ? '' : 'l'} cort" title="${esc(txt)}">${esc(txt)}</td>`;
        }).join('')}</tr>`).join('')}</tbody>
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
  /* espaço duplicado/no fim vira valor DIFERENTE na lista do filtro: 'DA PAO
     GRANDE' aparecia duas vezes (292 e 15), e marcar uma não pegava a outra.
     O arquivo do haras é digitado à mão; normalizar aqui junta as duas. */
  return String(v).replace(/\s+/g, ' ').trim();
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
  // mesma normalização de espaço do plantel (ver fmtCel)
  if (!ehNum) return String(v == null ? '' : v).replace(/\s+/g, ' ').trim() || '—';
  if (rot === 'Cota') return pct(v);
  if (COLS_DELTA.includes(rot)) return v ? rs(v) : '—';
  return rs(v);
}
const ordMov = (m, i) => COLS_MOV[i][2] ? num(COLS_MOV[i][1](m)) : String(COLS_MOV[i][1](m) || '');

/* ---------- a fila, em fichas ----------

   Era uma tabela de 16 colunas. Três problemas que a ficha resolve de uma vez:
   a descrição do que aconteceu — a única informação da linha que não se deduz de
   nenhuma outra coluna — ficava no fim da tabela, longe do botão de registro;
   o `<select>` de "foi outra coisa" destoava no meio de botões; e nada cabia na
   tela sem rolagem horizontal.

   Cada pendência vira um cartão que faz UMA pergunta: foi isto? A sugestão vem
   como o botão aceso; as outras classes ficam ao lado, do mesmo tamanho, porque
   escolher outra é tão legítimo quanto confirmar. */
function fichaMov(m){
  const deltas = [];
  if (m.mudou_status) deltas.push(['status', `${esc(m.mudou_status[0])} → ${esc(m.mudou_status[1])}`]);
  if (m.cota_ant !== m.cota_atual) deltas.push(['cota', `${pct(m.cota_ant)} → ${pct(m.cota_atual)}`]);
  if (m.valor_ant !== m.valor_atual) deltas.push(['valor', `${rs(m.valor_ant)} → ${rs(m.valor_atual)}`]);
  if (m.mudou_local) deltas.push(['local', `${esc(m.mudou_local[0])} → ${esc(m.mudou_local[1])}`]);
  if (m.mudou_dono) deltas.push(['dono', `${ATRIB[m.mudou_dono[0]] || '—'} → ${ATRIB[m.mudou_dono[1]] || '—'}`]);
  if (m.renome) deltas.push(['nome', `${esc(m.renome.de)} → ${esc(m.renome.para)}`]);
  if (m.entrou) deltas.push(['entrada', 'entrou no plantel']);
  if (m.saiu) deltas.push(['saída', 'saiu do controle']);

  const canc = (m.log || []).map(x => cancelamento(x.ocorrencia)).find(c => c && c.tipo);
  const avisos = [];
  if (canc) avisos.push(canc.revenda ? 'venda cancelada e revendido'
    : `${canc.tipo} cancelada (cota ${canc.sentido === 'volta' ? 'volta' : 'sai'})`);
  if (m.posterior) avisos.push(`${m.posterior.length} ocorrência(s) POSTERIOR(es) ao mês `
    + '— mantido o valor do mês anterior');

  // sem sugestão, todas as classes viram opção de primeira linha
  const outras = CLASSES_MOV.filter(c => c !== m.sugestao);
  return `<div class="ficha">
    <div class="ficha-topo">
      <div class="ficha-nome" title="${esc(m.nome)}">${esc(m.nome)}</div>
      <div class="ficha-delta ${clsN(m.delta)}">${rs(m.delta)}</div>
    </div>
    <div class="ficha-meta">${[m.categoria, m.sufixo, ATRIB[m.dono]].filter(Boolean).map(esc).join(' · ')}</div>
    ${!deltas.length ? '' : `<dl class="ficha-mudou">${deltas.map(([r, v]) =>
      `<dt>${r}</dt><dd>${v}</dd>`).join('')}</dl>`}
    <!-- a conta do animal: de onde saiu, o que mexeu, onde parou. Patrimônio é
         cota × valor + comissão, por isso ele não é o "valor" da linha acima -->
    <div class="ficha-conta">
      <span><i>valor inicial</i><b>${rs(m.patr_ant)}</b></span>
      <span class="seta-conta">→</span>
      <span><i>movimentação</i><b class="${clsN(m.delta)}">${rs(m.delta)}</b></span>
      <span class="seta-conta">→</span>
      <span><i>valor final</i><b>${rs(m.patr_atual)}</b></span>
      ${m.cota_atual ? `<span class="conta-obs">${pct(m.cota_atual)} de ${rs(m.valor_atual)}</span>` : ''}
    </div>
    ${(m.log || []).map(l => `<p class="ficha-log">${dataBR(l.data)} · ${esc(l.ocorrencia)}</p>`).join('')}
    ${avisos.map(a => `<p class="ficha-aviso">⚠ ${esc(a)}</p>`).join('')}
    <div class="ficha-perg">
      ${m.sugestao
        ? `<span class="ficha-perg-rot">Foi isto?</span>
           <button type="button" class="cls-bt sugerida" data-conf="${esc(m.chave)}:${esc(m.sugestao)}">
             ✓ ${esc(m.sugestao)}</button>`
        : `<span class="ficha-perg-rot sem-palpite">Sem sugestão — o que foi?</span>`}
      ${outras.map(c => `<button type="button" class="cls-bt" data-conf="${esc(m.chave)}:${c}">${c}</button>`).join('')}
    </div>
  </div>`;
}

function subMovimentacoes(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  if (!ST.meses[mesAnterior(ST.mes)]) {
    return `<div class="aviso">Importe também ${rotMes(mesAnterior(ST.mes))} para comparar os dois meses.</div>`;
  }
  /* Movimentações é a FILA do que falta decidir. Mês FECHADO não tem fila: o
     mês foi encerrado, não há o que perguntar, e mostrar as fichas com os
     botões "Foi isto?" convidava a uma ação que o banco recusaria. O mês
     encerrado se lê na Conciliação, que é onde ele fica inteiro à vista. */
  const todas = mv.movs.filter(m => m.no_escopo || m.dono);
  const fechado = mesFechado(ST.mes);
  if (fechado) {
    return `<div class="ok-vazio">${rotMes(ST.mes)} está fechado — não há fila.
      As ${todas.length} movimentações do mês estão na aba <b>Conciliação</b>.
      Para mexer, use <b>Reabrir mês</b> no topo.</div>`;
  }
  const base = todas.filter(m => !ST.decisoes[`${ST.mes}|${m.chave}`]);
  ST.ctx.mov = {base, txtDe: txtMov, ordDe: ordMov,
                rotulos: Object.fromEntries(COLS_MOV.map(([rot], i) => [i, rot]))};
  let movs = aplicaFiltros(base, 'mov', txtMov);
  if (ST.ordem.mov.col != null) {
    const i = ST.ordem.mov.col, dir = ST.ordem.mov.dir, pega = COLS_MOV[i][1], ehNum = COLS_MOV[i][2];
    movs = movs.slice().sort((a, b) => ehNum
      ? (num(pega(a)) - num(pega(b))) * dir
      : String(pega(a) || '').localeCompare(String(pega(b) || ''), 'pt-BR') * dir);
  }
  const nReg = todas.filter(m => ST.decisoes[`${ST.mes}|${m.chave}`]).length;
  if (!base.length) {
    return `<div class="ok-vazio">Nada na fila: as ${todas.length} movimentações de
      ${rotMes(ST.mes)} já foram confirmadas. Elas estão na aba Conciliação.</div>`;
  }
  // as colunas de texto viram o filtro da fila; as de dinheiro não fazem sentido aqui
  const FILTRAVEIS = [0, 1, 2, 3, 4];
  return `
    <div class="resumo-linha">
      <span>${movs.length}${movs.length === base.length ? '' : ' de ' + base.length} na fila de ${rotMes(ST.mes)}</span>
      <span>registrados: <b>${nReg}</b> de ${todas.length}</span>
      <span>Δ patrimônio: <b class="${clsN(movs.reduce((s, m) => s + m.delta, 0))}">${
        rs(movs.reduce((s, m) => s + m.delta, 0))}</b></span>
    </div>
    <div class="barra-filtro">
      <span class="barra-rot">filtrar</span>
      ${FILTRAVEIS.map(i => `<button type="button" class="fbtn-rot${
        (ST.filtros.mov[i] || []).length ? ' on' : ''}" data-fb="mov:${i}">${COLS_MOV[i][0]} ▾</button>`).join('')}
      <span class="barra-rot">ordenar</span>
      <button type="button" class="fbtn-rot${ST.ordem.mov.col == null ? ' on' : ''}" data-ord="mov:">maior Δ</button>
      <button type="button" class="fbtn-rot${ST.ordem.mov.col === 0 ? ' on' : ''}" data-ord="mov:0">nome</button>
    </div>
    ${chipsFiltro('mov')}
    <div class="rolagem fichas">${movs.map(m => fichaMov(m)).join('')}</div>`;
}

function subConciliacao(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  /* Saiu a fila 'animal novo sem dono definido': ela existia porque o mapa não
     cobria todo animal e alguém tinha de escolher entre Carla e Eduardo. Com o
     dono vindo do sufixo, não há o que escolher. */
  const semLog = mv.movs.filter(m => !m.log.length && Math.abs(m.delta) >= 1 && (m.no_escopo || m.dono));
  const logSemEfeito = mv.log.filter(l => l.tipo === 'financeira' &&
    !mv.movs.some(m => norm(m.nome) === norm(l.produto)));
  /* A divisão é essa — Movimentações é a FILA do que falta decidir; Conciliação
     é onde o mês fechado fica à vista: tudo que já foi registrado, mais os
     cruzamentos entre as duas fontes (movimento que o log não menciona,
     ocorrência do log sem efeito no patrimônio). */
  /* Mês fechado conta como resolvido por inteiro: ele foi encerrado fora desta
     tela (ou por ela), e listar só o que passou pelo botão diria que um mês
     fechado está vazio. Mês aberto lista o que já foi confirmado. */
  const registradas = mv.movs
    .filter(m => (m.no_escopo || m.dono)
                 && (mesFechado(ST.mes) || ST.decisoes[`${ST.mes}|${m.chave}`]))
    .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
  const bloco = (titulo, itens, render) => !itens.length ? '' :
    `<h3>${titulo} <span class="cont">${itens.length}</span></h3>${itens.map(render).join('')}`;
  const trancado = mesFechado(ST.mes);
  const manuais = manuaisDoMes(ST.mes);
  return `
    <div class="barra-mes">
      <span class="nota-acao">${trancado
        ? `${rotMes(ST.mes)} está fechado. Reabrir (no topo) mantém tudo que já foi classificado.`
        : `${rotMes(ST.mes)} está aberto. Fechar (no topo) trava o registro; dá pra reabrir depois.`}</span>
    </div>
    ${trancado ? '' : `
      <h3>Lançar movimentação à mão</h3>
      <div class="item form-manual">
        <input id="mNome" placeholder="Nome do animal" class="cresce">
        <select id="mClasse">${CLASSES_MOV.map(c => `<option value="${c}">${c}</option>`).join('')}</select>
        <input id="mValor" placeholder="Valor (R$)" inputmode="decimal">
        <input id="mNota" placeholder="Observação (opcional)" class="cresce">
        <button type="button" class="botao-acao primario" id="mAdd">Lançar</button>
      </div>`}
    ${!manuais.length ? '' : `
      <h3>Lançamentos manuais <span class="cont">${manuais.length}</span></h3>
      ${manuais.map(x => `<div class="item"><b>${esc(x.nome)}</b> ·
        <span class="reg-ok">${esc(x.classe)}</span> · <b class="${clsN(x.valor)}">${rs(x.valor)}</b>
        ${x.nota ? `· ${esc(x.nota)}` : ''}
        <span class="autor">${esc(x.autor || '')}</span>
        ${trancado ? '' : `<span class="acoes"><button type="button" data-rmman="${esc(x.chave)}">remover</button></span>`}
      </div>`).join('')}`}
    ${!registradas.length ? '' : `
      <h3>Movimentações registradas em ${rotMes(ST.mes)} <span class="cont">${registradas.length}</span>${
        trancado ? '' : `<button type="button" class="h3-acao" data-reiniciar="${ST.mes}"
          title="apaga todas as classificações deste mês">limpar classificações</button>`}</h3>
      <div class="rolagem"><table class="t">
        <thead><tr><th class="l">Animal</th><th class="l">Classificação</th><th>Δ patrimônio</th>
          <th class="l">Quem registrou</th></tr></thead>
        <tbody>${registradas.map(m => {
          // sem decisão (mês fechado sem registro na tela) vale a classe apurada
          const d = ST.decisoes[`${ST.mes}|${m.chave}`] || {classe: m.sugestao, autor: 'apurado'};
          return `<tr><td class="l cort" title="${esc(m.nome)}">${esc(m.nome)}</td>
            <td class="l">${trancado ? `<span class="reg-ok">✓ ${esc(d.classe)}</span>`
              : `<select data-reg="${esc(m.chave)}">${CLASSES_MOV.map(c =>
                  `<option value="${c}"${d.classe === c ? ' selected' : ''}>${c}</option>`).join('')}</select>`}</td>
            <td class="${clsN(m.delta)}">${rs(m.delta)}</td>
            <td class="l">${esc(d.autor || '')}</td></tr>`;
        }).join('')}
        <tr class="tot"><td>total</td><td></td>
          <td class="${clsN(registradas.reduce((s, m) => s + m.delta, 0))}">${
            rs(registradas.reduce((s, m) => s + m.delta, 0))}</td><td></td></tr>
        </tbody></table></div>`}
    ${bloco('Movimentação sem ocorrência no log do haras', semLog, m =>
      `<div class="item"><b>${esc(m.nome)}</b> · ${rs(m.delta)} ·
        cota ${pct(m.cota_ant)} → ${pct(m.cota_atual)} · valor ${rs(m.valor_ant)} → ${rs(m.valor_atual)}
        · status ${esc(m.status)}</div>`)}
    ${bloco('Ocorrência no log sem efeito no patrimônio', logSemEfeito, l =>
      `<div class="item"><b>${esc(l.produto)}</b> · ${dataBR(l.data)}<br><span class="log">${esc(l.ocorrencia)}</span></div>`)}
`;
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
  /* Cotas: o MESMO check do patrimônio, na outra unidade. O saldo em R$ pode
     fechar com a cota errada (valor sobe e cota cai na mesma proporção), então
     conferir só dinheiro deixa passar troca de participação. A soma das cotas do
     arquivo do haras é o lado direito; o esquerdo é o que estamos fechando —
     mês anterior mais o que as movimentações do mês mexeram. */
  const cotIni = cotasMes(mesAnterior(ST.mes)), cotFim = cotasMes(ST.mes);
  const cotMov = (mv ? mv.movs : []).reduce((s, m) => s + (num(m.cota_atual) - num(m.cota_ant)), 0);
  const linhas = [
    ['Valor inicial + movimentações = valor final (Carla)', iniC + movC, fimC],
    ['Valor inicial + movimentações = valor final (Carla + Eduardo)', iniCE + movCE, fimCE],
  ];
  // sem o mês anterior carregado não há 'inicial': o check acusaria diferença
  // que é só ausência de base
  if (ST.meses[mesAnterior(ST.mes)]) linhas.push(
    ['Cotas (%): inicial + movimentações = plantel do haras', cotIni + cotMov, cotFim, 'cota']);
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
  /* Lançamento manual é, por definição, o que o diff dos arquivos NÃO viu — ele
     entra no resumo contábil e não no apurado, então a diferença entre os dois
     passa a ser exatamente ele. Fica como linha própria pra que essa diferença
     tenha nome, em vez de virar um check vermelho sem explicação. */
  const manuais = manuaisDoMes(ST.mes);
  if (manuais.length) linhas.push(
    [`Lançamentos manuais (${manuais.length}) — fora do apurado`,
     manuais.reduce((s, x) => s + x.valor, 0), 0]);
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
    <tbody>${linhas.map(([t, a, b, un]) => {
      const dif = +(a - b).toFixed(4);
      /* tolerância de R$ 1: o Resumo Contábil divulgado carrega centavos de
         arredondamento próprio (15.970.552,61 contra 15.970.552,71). Em COTA a
         régua é outra — 0,0001 de cota é 0,01% e já é diferença de verdade. */
      const ok = un === 'cota' ? Math.abs(dif) < 0.0001 : Math.abs(dif) < 1;
      const f = un === 'cota' ? pct : rs;
      return `<tr><td class="l">${t}</td><td>${f(a)}</td><td>${f(b)}</td>
        <td class="${ok ? 'pos' : 'neg'}">${f(dif)}</td>
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
    if (aba) {
      ST.aba = aba.dataset.aba; ST.pop = null; pintaPop();
      /* O resumo é YTD: precisa de TODOS os meses. Pintar antes de carregar
         mostrava um resumo com um mês só — o único que estava na memória —, que
         se corrigia sozinho no clique seguinte e parecia filtro quebrado. Agora
         avisa que está carregando e só desenha com o ano inteiro na mão. */
      const ateAqui = ST.disponiveis.filter(m => m <= ST.mes);
      if (ST.aba === 'resumo' && ateAqui.some(m => !ST.meses[m])) {
        document.getElementById('painel').innerHTML =
          '<div class="aviso">Carregando os meses do ano para o resumo…</div>';
        await garanteMeses(ateAqui);
      }
      pinta();
      return;
    }
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
      // índice vazio = ordem natural da apuração (maior Δ primeiro)
      if (ci === '') { ST.ordem[qual] = {col: null, dir: 1}; pinta(); return; }
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
    /* clique na célula corta/descorta o texto: a reticência resolve a largura
       da coluna, mas some com o valor. O title já mostra no hover; o clique
       serve pra quem quer ler sem segurar o mouse, e afeta só aquela célula —
       a altura das outras linhas não muda. */
    const st = e.target.closest('[data-mes-status]');
    if (st) {
      const [mes, acao] = st.dataset.mesStatus.split(':');
      const fechar = acao === 'fechar';
      const pend = fechar ? (movimentacaoDoMes(mes) || {movs: []}).movs
        .filter(m => (m.no_escopo || m.dono) && !ST.decisoes[`${mes}|${m.chave}`]).length : 0;
      const txt = fechar
        ? (pend ? `Fechar ${rotMes(mes)} com ${pend} movimentação(ões) ainda na fila?\n\n`
                  + 'Elas ficam sem classificação e o mês para de aceitar registro.'
                : `Fechar ${rotMes(mes)}? O mês para de aceitar registro.`)
        : `Reabrir ${rotMes(mes)}? Nada do que já foi classificado se perde.`;
      if (confirm(txt) && await mudaStatusMes(mes, fechar)) { topo(); pinta(); }
      return;
    }
    const rmm = e.target.closest('[data-rmman]');
    if (rmm) {
      const chave = rmm.dataset.rmman;
      if (confirm('Remover este lançamento manual?')) {
        const c = sb();
        if (c) {
          const { error } = await c.from('plantel_mov_classificacao').delete()
            .eq('mes', ST.mes).eq('chave', chave);
          if (error) { alert('não removeu: ' + error.message); return; }
        }
        delete ST.decisoes[`${ST.mes}|${chave}`];
        pinta();
      }
      return;
    }
    if (e.target.id === 'mAdd') {
      const nome = (document.getElementById('mNome').value || '').trim();
      const classe = document.getElementById('mClasse').value;
      const bruto = (document.getElementById('mValor').value || '').replace(/\./g, '').replace(',', '.');
      const valor = Number(bruto);
      const nota = (document.getElementById('mNota').value || '').trim();
      if (!nome) { alert('falta o nome do animal'); return; }
      if (!isFinite(valor) || !valor) { alert('falta o valor (ex.: 60000 ou -60000)'); return; }
      if (await lancaManual(ST.mes, nome, classe, valor, nota)) pinta();
      return;
    }
    const rei = e.target.closest('[data-reiniciar]');
    if (rei) {
      const mes = rei.dataset.reiniciar;
      const n = Object.keys(ST.decisoes).filter(k => k.startsWith(mes + '|')).length;
      if (confirm(`Apagar as ${n} classificação(ões) de ${rotMes(mes)} e recomeçar?\n\n`
                  + 'Elas voltam para a fila da aba Movimentações. Não dá pra desfazer.')) {
        if (await reiniciaMes(mes)) pinta();
      }
      return;
    }
    const conf = e.target.closest('[data-conf]');
    if (conf) {
      // corta no ULTIMO ':' — nome de animal com dois-pontos partiria a chave
      const bruto = conf.dataset.conf, corte = bruto.lastIndexOf(':');
      const chave = bruto.slice(0, corte), classe = bruto.slice(corte + 1);
      const mv = movimentacaoDoMes(ST.mes);
      const mov = mv && mv.movs.find(m => m.chave === chave);
      if (mov) { await registra(ST.mes, mov, classe, ''); pinta(); }
      return;
    }
    const cel = e.target.closest('td.cort');
    if (cel) { cel.classList.toggle('aberta'); return; }
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
  ST.disponiveis = await listaMeses();
  ST.mes = ST.disponiveis[ST.disponiveis.length - 1] || null;
  topo(); liga(); pinta();              // pinta já, mesmo sem o mês na mão
  /* O mês pedido primeiro e sozinho: é o que a aba Plantel precisa pra mostrar
     alguma coisa. O mês ANTERIOR só serve pra aba Movimentações comparar, então
     entra depois, sem segurar a primeira pintura. */
  await garanteMeses([ST.mes]);
  pinta();
  garanteMeses([mesAnterior(ST.mes)]).then(pinta);
  carregaDecisoes().then(pinta);
})();
