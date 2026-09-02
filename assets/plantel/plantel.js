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
const CLASSES_MOV = ['compra', 'embriao', 'venda', 'morte', 'doacao', 'reavaliacao', 'renome', 'sem_efeito'];
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
  meses: {},        // 'AAAA-MM' -> {arquivo, linhas, log, atribuicao, temSplit}
  mes: null,
  decisoes: {},     // 'mes|chave' -> {classe, nota, autor}
  aba: 'plantel',
  sub: 'movimentacoes',
  ordem: {col: null, dir: 1},
  filtros: {},
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

/* atribuição Carla/Eduardo: do próprio arquivo quando ele tem o split, senão
   herdada do mês anterior. Animal sem atribuição fica pendente — não se
   adivinha de quem é o animal. */
/* De quem é cada animal. Do arquivo quando ele traz o split (mapa do Luxor),
   senão herdada do mês anterior. Não achou nem um nem outro: SUGERE pelo sufixo
   e marca como sugestão — animal novo aparece na conciliação pra ser confirmado,
   porque a atribuição não é fórmula (no mapa a cota inteira vai pra um dos dois
   e há animal em nenhum dos dois). */
function montaAtribuicao(mes){
  const d = ST.meses[mes], ix = d.ix;
  const anterior = mesAnterior(mes);
  const ant = anterior && ST.meses[anterior];
  const herdada = ant ? ant.atribuicao || {} : {};
  const at = {}, sugeridos = [];
  for (const l of d.linhas) {
    const k = chaveCom(l, ix);
    if (d.temSplit) {
      const temHpg = num(l[ix.hpgCota]) || num(l[ix.hpgVal]);
      const temEd = num(l[ix.edCota]) || num(l[ix.edVal]);
      at[k] = temHpg ? 'hpg' : temEd ? 'eduardo' : 'nenhum';
    } else if (k in herdada) {
      at[k] = herdada[k];
    } else {
      const suf = norm(l[ix.sufixo]);
      at[k] = suf === 'DA PAO GRANDE' || suf === 'OUTRO' ? 'hpg' : suf.includes(' - E ') ? 'eduardo' : 'nenhum';
      sugeridos.push(k);
    }
  }
  d.atribuicao = at;
  d.sugeridos = sugeridos;
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
function patr(l, at, escopo, ix){
  ix = ix || (ST.meses[ST.mes] && ST.meses[ST.mes].ix) || {};
  const a = at[chaveCom(l, ix)];
  if (escopo === 'hpg' && a !== 'hpg') return 0;
  if (escopo === 'carla_eduardo' && a !== 'hpg' && a !== 'eduardo') return 0;
  return num(l[ix.cota]) * num(l[ix.valor]) + num(l[ix.comissao]);
}
const patrMes = (mes, escopo) => {
  const d = ST.meses[mes];
  return d ? d.linhas.reduce((s, l) => s + patr(l, d.atribuicao || {}, escopo, d.ix), 0) : 0;
};

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
  const at = d.atribuicao || {};
  const noEscopo = l => ['hpg', 'eduardo'].includes(at[chave(l)]);

  const log = (d.log || []).filter(x => {
    const dt = x.data instanceof Date ? x.data : new Date(x.data);
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}` === mes;
  }).map(x => ({...x, tipo: tipoLog(x.ocorrencia)}));

  const logPorNome = {};
  for (const x of log) (logPorNome[norm(x.produto)] = logPorNome[norm(x.produto)] || []).push(x);

  // pontes de identidade: nome antigo -> chave nova (só quando a antiga sumiu)
  const idxAnt = {}, idxAtual = {};
  if (ant) for (const l of ant.linhas) idxAnt[norm(l[COL_NOME])] = chave(l);
  for (const l of d.linhas) idxAtual[norm(l[COL_NOME])] = chave(l);
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

  const mapa = ls => {
    const o = {};
    for (const l of ls) o[ponte[chave(l)] || chave(l)] = l;
    return o;
  };
  const A = ant ? mapa(ant.linhas) : {};
  const B = mapa(d.linhas);
  const atAnt = ant ? ant.atribuicao || {} : {};

  const movs = [];
  for (const k of new Set([...Object.keys(A), ...Object.keys(B)])) {
    const a = A[k], b = B[k];
    const q0 = a ? num(a[COL_COTA]) : 0, q1 = b ? num(b[COL_COTA]) : 0;
    const v0 = a ? num(a[COL_VALOR]) : 0, v1 = b ? num(b[COL_VALOR]) : 0;
    const p0 = q0 * v0, p1 = q1 * v1;
    const ren = renomes.find(r => r.chave === k);
    const mudouStatus = a && b && norm(a[6]) !== norm(b[6]);
    const mudouLocal = a && b && norm(a[7]) !== norm(b[7]);
    if (Math.abs(p1 - p0) < 0.01 && !ren && !mudouStatus && !mudouLocal && a && b) continue;

    const linha = b || a;
    const dono = at[k] || atAnt[k] || null;
    const itensLog = logPorNome[norm(linha[COL_NOME])] || (ren ? logPorNome[norm(ren.para)] : []) || [];
    movs.push({
      chave: k, linha, nome: linha[COL_NOME], sufixo: linha[COL_SUFIXO],
      categoria: linha[5], status: linha[6], dono,
      cota_ant: q0, cota_atual: q1, valor_ant: v0, valor_atual: v1,
      patr_ant: p0, patr_atual: p1, delta: +(p1 - p0).toFixed(2),
      entrou: !a, saiu: !b, renome: ren || null,
      mudou_status: mudouStatus ? [a[6], b[6]] : null,
      mudou_local: mudouLocal ? [a[7], b[7]] : null,
      log: itensLog,
      sugestao: sugere({q0, q1, v0, v1, ren, entrou: !a, saiu: !b, linha, itensLog}),
      no_escopo: noEscopo(linha),
    });
  }
  movs.sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta));
  return {movs, log, renomes, novos: d.novos || []};
}

/* sugestão pelo padrão — só sugestão; o registro é o input de quem fecha */
function sugere({q0, q1, v0, v1, ren, entrou, saiu, linha, itensLog}){
  const st = norm(linha[6]), cat = norm(linha[5]);
  const oc = norm((itensLog || []).map(x => x.ocorrencia).join(' '));
  if (ren) return 'renome';
  if (entrou) return /EMBRI/.test(cat) ? 'embriao' : /NASCEU/.test(oc) ? 'embriao' : 'compra';
  if (saiu || (q0 && !q1)) {
    if (/MORREU|OBITO|ABORTOU/.test(oc) || /OBITO/.test(st)) return 'morte';
    if (/DOAD|DOACAO/.test(oc)) return 'doacao';
    return 'venda';
  }
  if (q1 > q0) return 'compra';
  if (q1 < q0) return /MORREU|OBITO/.test(oc) ? 'morte' : /DOAD/.test(oc) ? 'doacao' : 'venda';
  if (v1 !== v0) return 'reavaliacao';
  return 'sem_efeito';
}

/* ================= resumo contábil (YTD, só Carla) ================= */
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
    const at = d.atribuicao || {}, atA = da.atribuicao || {};
    const ini = da.linhas.reduce((s, l) => s + patr(l, atA, 'hpg'), 0);
    const fim = d.linhas.reduce((s, l) => s + patr(l, at, 'hpg'), 0);
    const mv = movimentacaoDoMes(m);
    const causas = {};
    for (const mo of (mv ? mv.movs : [])) {
      if ((at[mo.chave] || atA[mo.chave]) !== 'hpg') continue;
      const dec = ST.decisoes[`${m}|${mo.chave}`];
      const classe = dec ? dec.classe : mo.sugestao;
      causas[classe] = +( (causas[classe] || 0) + mo.delta ).toFixed(2);
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
  document.getElementById('topo').innerHTML = `
    <h1>Plantel Haras Pao Grande - Movimentação Jan a Dez ${(ST.mes || '2026').slice(0, 4)}</h1>
    <div class="ferramentas">
      <label class="mes-sel">Mês
        <select id="selMes">${meses.map(m => `<option value="${m}"${m === ST.mes ? ' selected' : ''}>${rotMes(m)}</option>`).join('')}</select>
      </label>
      <label class="botao-arquivo">Importar arquivo
        <input type="file" id="arq" accept=".xlsx,.xlsm" hidden>
      </label>
      <span id="statusImp"></span>
    </div>`;
  document.getElementById('selMes').onchange = e => { ST.mes = e.target.value; ST.ordem = {col: null, dir: 1}; ST.filtros = {}; pinta(); };
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
      ${d.temSplit ? '· traz atribuição Carla/Eduardo' : '· sem atribuição (será herdada do mês anterior)'}
      ${ultimo ? `<br><span class="log">última ocorrência no arquivo: ${ultimo.toLocaleDateString('pt-BR')}</span>` : ''}</div>
    <label>Mês de fechamento
      <select id="impMes">${MESES_PT.map((r, i) => {
        const m = `${ano}-${String(i + 1).padStart(2, '0')}`;
        return `<option value="${m}"${m === mesSugerido ? ' selected' : ''}>${r}/${ano.slice(2)}</option>`;
      }).join('')}</select></label>
    <button type="button" id="impOk">Processar</button>
    <button type="button" id="impNao" class="secundario">Cancelar</button>`;
  document.getElementById('impNao').onclick = () => { caixa.hidden = true; caixa.innerHTML = ''; };
  document.getElementById('impOk').onclick = async () => {
    const mes = document.getElementById('impMes').value;
    ST.meses[mes] = d;
    montaAtribuicao(mes);
    ST.mes = mes;
    caixa.hidden = true; caixa.innerHTML = '';
    await salvaSnapshot(mes, d, arquivo);
    topo(); pinta();
  };
}

async function salvaSnapshot(mes, d, arquivo){
  const c = sb();
  if (!c) return;
  try {
    await c.from('plantel_snapshot').upsert({
      mes, arquivo, linhas: d.linhas, log: d.log.map(x => ({...x, data: (x.data instanceof Date ? x.data : new Date(x.data)).toISOString()})),
      atribuicao: d.atribuicao, importado_por: eu(),
    }, {onConflict: 'mes'});
  } catch (err) { console.warn('[plantel] snapshot não gravado', err.message || err); }
}

async function carregaSnapshots(){
  const c = sb();
  if (!c) return;
  try {
    const { data } = await c.from('plantel_snapshot').select('mes,arquivo,linhas,log,atribuicao');
    for (const r of data || []) {
      ST.meses[r.mes] = {arquivo: r.arquivo, linhas: r.linhas || [], log: (r.log || []).map(x => ({...x, data: new Date(x.data)})),
                         atribuicao: r.atribuicao || {}, temSplit: false, cab: []};
    }
    const { data: dec } = await c.from('plantel_mov_classificacao').select('mes,chave,classe,nota,autor');
    for (const r of dec || []) ST.decisoes[`${r.mes}|${r.chave}`] = r;
  } catch (err) { console.warn('[plantel] sem dados salvos', err.message || err); }
}

async function registra(mes, mov, classe, nota){
  ST.decisoes[`${mes}|${mov.chave}`] = {classe, nota, autor: eu() || '(local)'};
  const c = sb();
  if (c) {
    const { error } = await c.from('plantel_mov_classificacao')
      .upsert({mes, chave: mov.chave, classe, nota: nota || null, nome: mov.nome}, {onConflict: 'mes,chave'});
    if (error) { alert('não gravou: ' + error.message); return false; }
  }
  return true;
}

async function atribui(mes, k, dono){
  ST.meses[mes].atribuicao[k] = dono;
  ST.meses[mes].novos = (ST.meses[mes].novos || []).filter(x => x !== k);
  const c = sb();
  if (c) await c.from('plantel_snapshot').update({atribuicao: ST.meses[mes].atribuicao}).eq('mes', mes);
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
  const at = d.atribuicao || {}, ix = d.ix;
  const cols = colunasDoMes(d);
  let linhas = d.linhas.slice();
  for (const [ci, txt] of Object.entries(ST.filtros)) {
    if (!txt) continue;
    const i = +ci, t = norm(txt);
    linhas = linhas.filter(l => norm(fmtCel(l, i, d.cab[i])).includes(t));
  }
  if (ST.ordem.col != null) {
    const i = ST.ordem.col, dir = ST.ordem.dir, rot = d.cab[i];
    linhas.sort((a, b) => EH_NUM(rot)
      ? (num(a[i]) - num(b[i])) * dir
      : String(a[i] == null ? '' : a[i]).localeCompare(String(b[i] == null ? '' : b[i]), 'pt-BR') * dir);
  }
  const somaEsc = escopo => linhas.reduce((s, l) => s + patr(l, at, escopo, ix), 0);
  return `
    <div class="resumo-linha">
      <span>${linhas.length} de ${d.linhas.length} linhas</span>
      <span>Carla: <b>${rs(somaEsc('hpg'))}</b></span>
      <span>Carla + Eduardo: <b>${rs(somaEsc('carla_eduardo'))}</b></span>
      ${Object.values(ST.filtros).some(v => v) ? '<button type="button" id="limpaF">limpar filtros</button>' : ''}
    </div>
    <div class="rolagem"><table class="t">
      <thead>
        <tr>${cols.map(([i, r]) => `<th data-ord="${i}" class="${EH_NUM(r) ? '' : 'l'}${ST.ordem.col === i ? ' ord' : ''}">${esc(r)}${ST.ordem.col === i ? (ST.ordem.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('')}<th class="l">Dono</th></tr>
        <tr class="filtros-linha">${cols.map(([i]) => `<th><input data-f="${i}" value="${esc(ST.filtros[i] || '')}"></th>`).join('')}<th></th></tr>
      </thead>
      <tbody>${linhas.map(l => `<tr>${cols.map(([i, r]) =>
        `<td class="${EH_NUM(r) ? '' : 'l'}">${esc(fmtCel(l, i, r))}</td>`).join('')}
        <td class="l">${ATRIB[at[chaveCom(l, ix)]] || '<span class="zero">—</span>'}${
          (d.sugeridos || []).includes(chaveCom(l, ix)) ? ' <span class="sug">sugerido</span>' : ''}</td></tr>`).join('')}</tbody>
    </table></div>`;
}

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

function subMovimentacoes(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  if (!ST.meses[mesAnterior(ST.mes)]) {
    return `<div class="aviso">Importe também ${rotMes(mesAnterior(ST.mes))} para comparar os dois meses.</div>`;
  }
  const movs = mv.movs.filter(m => m.no_escopo || m.dono);
  const soma = k => movs.reduce((s, m) => {
    const dec = ST.decisoes[`${ST.mes}|${m.chave}`];
    return s + ((dec ? dec.classe : m.sugestao) === k ? m.delta : 0);
  }, 0);
  return `
    <div class="resumo-linha">
      <span>${movs.length} animais com movimentação em ${rotMes(ST.mes)}</span>
      <span>registrados: <b>${movs.filter(m => ST.decisoes[`${ST.mes}|${m.chave}`]).length}</b> de ${movs.length}</span>
      <span>Δ patrimônio: <b class="${clsN(movs.reduce((s, m) => s + m.delta, 0))}">${rs(movs.reduce((s, m) => s + m.delta, 0))}</b></span>
    </div>
    <div class="rolagem"><table class="t">
      <thead><tr>
        <th class="l">Nome</th><th class="l">Sufixo</th><th class="l">Categoria</th><th class="l">Status</th>
        <th class="l">Dono</th><th>Cota</th><th>Valor</th><th>Valor inicial</th>
        <th>Compras</th><th>Embriões</th><th>Venda</th><th>Morte/doação</th><th>Reavaliação</th><th>Valor final</th>
        <th class="l">O que foi feito</th><th class="l">Registro</th>
      </tr></thead>
      <tbody>${movs.map(m => linhaMov(m)).join('')}
        <tr class="tot">
          <td colspan="7">total</td>
          <td>${rs(movs.reduce((s, m) => s + m.patr_ant, 0))}</td>
          <td class="${clsN(soma('compra'))}">${rs(soma('compra'))}</td>
          <td class="${clsN(soma('embriao'))}">${rs(soma('embriao'))}</td>
          <td class="${clsN(soma('venda'))}">${rs(soma('venda'))}</td>
          <td class="${clsN(soma('morte') + soma('doacao'))}">${rs(soma('morte') + soma('doacao'))}</td>
          <td class="${clsN(soma('reavaliacao'))}">${rs(soma('reavaliacao'))}</td>
          <td>${rs(movs.reduce((s, m) => s + m.patr_atual, 0))}</td>
          <td colspan="2"></td>
        </tr>
      </tbody></table></div>`;
}

function linhaMov(m){
  const dec = ST.decisoes[`${ST.mes}|${m.chave}`];
  const classe = dec ? dec.classe : null;
  const cel = k => {
    const v = (classe || m.sugestao) === k ? m.delta : 0;
    return `<td class="${clsN(v)}">${v ? rs(v) : '—'}</td>`;
  };
  const oque = [];
  if (m.renome) oque.push(`renome: <b>${esc(m.renome.de)}</b> → <b>${esc(m.renome.para)}</b>`);
  if (m.entrou) oque.push('entrou no plantel');
  if (m.saiu) oque.push('saiu do controle');
  if (m.mudou_status) oque.push(`status: ${esc(m.mudou_status[0])} → ${esc(m.mudou_status[1])}`);
  if (m.mudou_local) oque.push(`local: ${esc(m.mudou_local[0])} → ${esc(m.mudou_local[1])}`);
  if (m.cota_ant !== m.cota_atual) oque.push(`cota: ${pct(m.cota_ant)} → ${pct(m.cota_atual)}`);
  if (m.valor_ant !== m.valor_atual) oque.push(`valor: ${rs(m.valor_ant)} → ${rs(m.valor_atual)}`);
  for (const l of m.log) oque.push(`<span class="log">${dataBR(l.data)} · ${esc(l.ocorrencia)}</span>`);
  return `<tr class="${dec ? 'reg' : ''}">
    <td class="l nome" title="${esc(m.nome)}">${esc(m.nome)}</td>
    <td class="l">${esc(m.sufixo)}</td><td class="l">${esc(m.categoria)}</td><td class="l">${esc(m.status)}</td>
    <td class="l">${ATRIB[m.dono] || '<span class="zero">—</span>'}</td>
    <td>${pct(m.cota_atual)}</td><td>${rs(m.valor_atual)}</td>
    <td>${rs(m.patr_ant)}</td>
    ${cel('compra')}${cel('embriao')}${cel('venda')}
    <td class="${clsN((classe || m.sugestao) === 'morte' || (classe || m.sugestao) === 'doacao' ? m.delta : 0)}">${['morte', 'doacao'].includes(classe || m.sugestao) ? rs(m.delta) : '—'}</td>
    ${cel('reavaliacao')}
    <td>${rs(m.patr_atual)}</td>
    <td class="l oque">${oque.join('<br>') || '—'}</td>
    <td class="l">
      <select data-reg="${esc(m.chave)}">
        <option value="">— registrar —</option>
        ${CLASSES_MOV.map(c => `<option value="${c}"${classe === c ? ' selected' : ''}>${c}${!classe && c === m.sugestao ? ' (sugerido)' : ''}</option>`).join('')}
      </select>
      ${dec ? `<span class="autor">${esc(dec.autor || '')}</span>` : ''}
    </td></tr>`;
}

function subConciliacao(){
  const mv = movimentacaoDoMes(ST.mes);
  if (!mv) return semArquivo();
  const semDono = (mv.novos || []);
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
      const l = ST.meses[ST.mes].linhas.find(x => chave(x) === k) || [];
      return `<div class="item"><b>${esc(l[COL_NOME])}</b> · ${esc(l[COL_SUFIXO])} · ${esc(l[5])} ·
        cota ${pct(l[COL_COTA])} · ${rs(num(l[COL_COTA]) * num(l[COL_VALOR]))}
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
  const at = d.atribuicao || {};
  const mv = movimentacaoDoMes(ST.mes);
  const iniC = ant ? ant.linhas.reduce((s, l) => s + patr(l, ant.atribuicao || {}, 'hpg'), 0) : 0;
  const fimC = d.linhas.reduce((s, l) => s + patr(l, at, 'hpg'), 0);
  const movC = (mv ? mv.movs : []).filter(m => (at[m.chave] || (ant && (ant.atribuicao || {})[m.chave])) === 'hpg')
    .reduce((s, m) => s + m.delta, 0);
  const iniCE = ant ? ant.linhas.reduce((s, l) => s + patr(l, ant.atribuicao || {}, 'carla_eduardo'), 0) : 0;
  const fimCE = d.linhas.reduce((s, l) => s + patr(l, at, 'carla_eduardo'), 0);
  const movCE = (mv ? mv.movs : []).filter(m => m.no_escopo || m.dono).reduce((s, m) => s + m.delta, 0);
  const somaCotaValor = d.linhas.reduce((s, l) => s + patr(l, at, 'hpg'), 0);

  const linhas = [
    ['Valor inicial + movimentações = valor final (Carla)', iniC + movC, fimC],
    ['Valor inicial + movimentações = valor final (Carla + Eduardo)', iniCE + movCE, fimCE],
    ['Σ (cota × valor) dos animais da Carla = saldo do resumo contábil', somaCotaValor, fimC],
  ];
  const regs = (mv ? mv.movs : []).filter(m => m.no_escopo || m.dono);
  return `<div class="rolagem"><table class="t">
    <thead><tr><th class="l">Check de ${rotMes(ST.mes)}</th><th>Apurado</th><th>Esperado</th><th>Diferença</th><th class="l">Situação</th></tr></thead>
    <tbody>${linhas.map(([t, a, b]) => {
      const dif = +(a - b).toFixed(2);
      return `<tr><td class="l">${t}</td><td>${rs(a)}</td><td>${rs(b)}</td>
        <td class="${Math.abs(dif) < 0.01 ? 'pos' : 'neg'}">${rs(dif)}</td>
        <td class="l">${Math.abs(dif) < 0.01 ? '<span class="tag ok">confere</span>' : '<span class="tag ruim">diverge</span>'}</td></tr>`;
    }).join('')}
    <tr><td class="l">Movimentações registradas</td><td>${regs.filter(m => ST.decisoes[`${ST.mes}|${m.chave}`]).length}</td>
      <td>${regs.length}</td><td></td>
      <td class="l">${regs.every(m => ST.decisoes[`${ST.mes}|${m.chave}`]) && regs.length
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
}

function liga(){
  document.body.addEventListener('click', async e => {
    const aba = e.target.closest('[data-aba]');
    if (aba) { ST.aba = aba.dataset.aba; pinta(); return; }
    const sub = e.target.closest('[data-sub]');
    if (sub) { ST.sub = sub.dataset.sub; pinta(); return; }
    const ord = e.target.closest('[data-ord]');
    if (ord) {
      const i = +ord.dataset.ord;
      ST.ordem = {col: i, dir: ST.ordem.col === i ? -ST.ordem.dir : 1};
      pinta(); return;
    }
    if (e.target.id === 'limpaF') { ST.filtros = {}; pinta(); return; }
    const dono = e.target.closest('[data-dono]');
    if (dono) {
      const [k, dn] = dono.dataset.dono.split(':');
      await atribui(ST.mes, k, dn);
      pinta(); return;
    }
  });
  document.body.addEventListener('input', e => {
    if (e.target.dataset.f != null) {
      ST.filtros[e.target.dataset.f] = e.target.value;
      const pos = e.target.dataset.f;
      pinta();
      const novo = document.querySelector(`[data-f="${pos}"]`);
      if (novo) { novo.focus(); novo.setSelectionRange(novo.value.length, novo.value.length); }
    }
  });
  document.body.addEventListener('change', async e => {
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
