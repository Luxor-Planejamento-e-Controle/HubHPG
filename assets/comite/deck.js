/* Comitê Mensal HPG — renderiza o spec como deck e exporta o mesmo spec em PPTX.

   Regra do desenho: o spec (assets/comite/spec.js, gerado por
   hub/tools/build_comite.py) é a ÚNICA fonte. HTML e PPTX são duas saídas do
   mesmo objeto — se um slide muda, muda nos dois, porque nenhum dos dois guarda
   conteúdo próprio. Slide sem fonte de dado vira `pendente` e diz isso na tela;
   número não se inventa pra encher slide.

   O spec traz um deck por MÊS (`SPEC.decks`), e o seletor da barra troca de mês
   sem recarregar a página. */
'use strict';

const SPEC = window.COMITE_SPEC;
const C = {bg:'04223B', bg2:'072B49', card:'0A3050', line:'1B486B', ink:'EAF0F4',
           ink2:'B6C8D6', ink3:'93AABC', amber:'CA9703', pos:'4CC38A', neg:'F07A7A'};
const LOGO = 'assets/pg-logo.png';
/* área útil do slide: 1280×720 menos cabeçalho (96) e rodapé (44) */
const BODY_H = 720 - 96 - 44;

/* ---- formatação ---- */
const nf = (v, d) => v.toLocaleString('pt-BR', {minimumFractionDigits:d, maximumFractionDigits:d});
const rs = v => v == null ? '—' : (v < 0 ? '-' : '') + 'R$ ' + nf(Math.abs(v), 0);
const dk = v => v == null ? '—' : (v >= 0 ? '+' : '−') + nf(Math.abs(v), 0) + 'k';
const dpct = v => v == null ? 'N/A' : (v >= 0 ? '+' : '−') + nf(Math.abs(v) * 100, 0) + '%';
const cls = v => v == null ? '' : v >= 0 ? 'pos' : 'neg';
const esc = s => String(s == null ? '' : s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const brdata = s => {
  if (!s) return '—';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s));
  return m ? `${m[3]}/${m[2]}/${m[1].slice(2)}` : String(s);
};
/* R$ k, como no deck original: R$18.703k */
function rsk(v){
  if (v == null || v === 0) return '—';
  const s = v < 0 ? '-' : '', a = Math.abs(v) / 1000;
  return s + 'R$' + nf(a, a < 100 ? 1 : 0) + 'k';
}
/* Altura de linha e corpo de fonte que fazem N linhas caberem em `alt`.
   Sem isso a tabela do DRE (40+ linhas) vazava por cima do rodapé — foi o que
   obrigou a diminuir o zoom do navegador pra conseguir ler. */
function ajusta(n, alt = BODY_H, maxH = 26){
  const h = Math.max(11, Math.min(maxH, alt / Math.max(n, 1)));
  return {h, fs: Math.max(8, Math.min(13, h * 0.55))};
}

/* ---- render HTML ---- */
const head = s => `<div class="s-head"><h1>${esc(s.titulo)}</h1>${s.sub ? `<div class="sub">${esc(s.sub)}</div>` : ''}</div>
  <img class="s-mark" src="${LOGO}" alt="">`;
const foot = (s, i, n) => `<div class="s-foot"><span>${esc(SPEC.labels[mesAtual])}</span>
  ${s.obs ? `<span class="obs">${esc(s.obs)}</span>` : ''}<span class="spacer"></span><span>${i + 1} / ${n}</span></div>`;
const kpiRow = ks => `<div class="kpis" style="grid-template-columns:repeat(${ks.length},1fr)">` +
  ks.map(k => `<div class="kpi"><div class="v">${esc(k.v)}</div><div class="l">${esc(k.l)}</div><div class="s">${esc(k.s)}</div></div>`).join('') + `</div>`;

function tabelaHTML(cols, rows, fmtCel, alt, larguras){
  const {h, fs} = ajusta(rows.length + 1, alt);
  return `<table class="t" style="font-size:${fs.toFixed(1)}px">
    ${larguras ? `<colgroup>${larguras.map(w => `<col style="width:${w}">`).join('')}</colgroup>` : ''}
    <thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(r => {
      const cel = r.cells || r;
      const nv = r.nivel != null ? r.nivel : (r.total ? 0 : null);
      return `<tr class="${nv === 0 ? 'tot' : nv === 1 ? 'sub' : ''}" style="height:${h.toFixed(1)}px">` +
        cel.map((c, j) => fmtCel(c, j, r)).join('') + `</tr>`;
    }).join('')}</tbody></table>`;
}

const R = {
  capa: s => `<div class="capa"><img src="${LOGO}" alt="">
    <h1>${esc(s.titulo)}</h1><div class="mes">${esc(s.mes)}</div><div class="org">${esc(s.org)}</div></div>`,

  encerramento: s => `<div class="fim"><img src="${LOGO}" alt=""><h1>${esc(s.titulo)}</h1></div>`,

  divisor: s => `<div class="divisor"><div class="n">0${s.n}</div>
    <h1>${esc(s.titulo)}</h1><div class="sub">${esc(s.sub)}</div></div>`,

  agenda: s => head(s) + `<div class="s-body"><div class="agenda">` +
    s.itens.map(it => `<div class="it"><div class="n">${esc(it.n)}</div>
      <h3>${esc(it.titulo)}</h3><p>${esc(it.sub)}</p></div>`).join('') + `</div></div>`,

  /* S04–S07, S10, S13, S14 — Orçado | Realizado | ∆ R$ k | ∆ %
     Três níveis: 0 grupo (dourado), 1 subgrupo (branco, recuado), 2 folha. */
  dre: s => head(s) + `<div class="s-body">` + tabelaHTML(
    ['NATUREZA', 'ORÇADO', 'REALIZADO', '∆ R$ k', '∆ %'],
    s.linhas.map(l => ({nivel: l.nivel == null ? (l.total ? 0 : 2) : l.nivel, cells: [l.nome, ...l.v]})),
    (c, j, r) => j === 0
      ? `<td class="nome n${r.nivel}">${esc(c)}</td>`
      : j === 3 ? `<td class="${cls(c)}">${dk(c)}</td>`
      : j === 4 ? `<td class="${cls(c)}">${dpct(c)}</td>`
      : `<td>${rs(c)}</td>`,
    BODY_H, ['40%', '', '', '12%', '10%']) + `</div>`,

  /* KPIs + tabela (S11, S16–S20, S29, S37) */
  kpis_tabela: s => head(s) + `<div class="s-body">${kpiRow(s.kpis)}` + tabelaHTML(
    s.tabela.cols, s.tabela.rows,
    (c, j) => `<td${j === 0 ? ' class="nome"' : ''}>${esc(c)}</td>`,
    BODY_H - 108) + `</div>`,

  /* tabela pura (S30, S32–S35) */
  tabela: s => head(s) + `<div class="s-body">` + (s.rows.length
    ? tabelaHTML(s.cols, s.rows, (c, j) => {
        if ((s.moeda || []).includes(j)) return `<td>${rs(c)}</td>`;
        if ((s.data || []).includes(j)) return `<td>${brdata(c)}</td>`;
        return `<td${j === 0 ? ' class="nome"' : ''}>${esc(c)}</td>`;
      }, BODY_H)
    : vazio('○', 'nada neste recorte', 'Sem registro para o filtro deste slide',
            'A base foi lida e respondeu vazio — não é falta de fonte.')) + `</div>`,

  /* S12 — KPIs + matriz título × meses */
  matriz: s => head(s) + `<div class="s-body">${kpiRow(s.kpis)}` + tabelaHTML(
    s.cols, s.rows.map(r => ({total: /Saldo/.test(r[0]), cells: r})),
    (c, j) => j === 0 ? `<td class="nome">${esc(c)}</td>`
                      : `<td class="${typeof c === 'number' && c ? cls(c) : ''}">${c ? rsk(c) : '—'}</td>`,
    BODY_H - 108) + `</div>`,

  /* S09 — investimentos mês a mês */
  lista_mes: s => {
    const n = s.meses.reduce((a, m) => a + 1 + m.itens.length, 0);
    const {fs} = ajusta(n, BODY_H, 22);
    return head(s) + `<div class="s-body"><div class="lista" style="font-size:${fs.toFixed(1)}px">` +
      s.meses.map(m => `<div class="m"><span class="mes">${esc(m.mes)}</span>
          <span class="tag">Animais e produtos</span><span class="tot">${rs(m.total)}</span></div>` +
        m.itens.map(it => `<div class="it"><span></span><span class="d">${esc(it.desc)}</span><span>${rs(it.valor)}</span></div>`).join('')
      ).join('') + `</div></div>`;
  },

  /* S08 — comentários do DRE: categoria · texto · delta */
  comentarios: s => {
    const {h, fs} = ajusta(s.itens.length, BODY_H, 56);
    return head(s) + `<div class="s-body"><div class="coment">` + s.itens.map(i =>
      `<div class="li" style="min-height:${h.toFixed(0)}px;font-size:${Math.min(12.5, fs + 1.5).toFixed(1)}px">
        <div class="cat">${esc(i.cat)}</div>
        <div class="txt">${esc(i.txt)}</div>
        <div class="d ${/^[-−]/.test(i.delta) ? 'neg' : 'pos'}">${esc(i.delta)}</div>
      </div>`).join('') + `</div></div>`;
  },

  /* S24+ — resultados de exposição: animal e seus prêmios, em duas colunas */
  resultados: s => head(s) + `<div class="s-body"><div class="premios">` +
    s.animais.map(a => `<div class="an"><div class="nome">${esc(a.nome)}</div>` +
      a.premios.map(p => `<div class="p">${esc(p)}</div>`).join('') + `</div>`).join('') +
    `</div></div>`,

  /* S38 — histórico de manejo, mês a mês */
  manejo: s => {
    const {h} = ajusta(s.itens.length, BODY_H, 74);
    return head(s) + `<div class="s-body"><div class="manejo">` + s.itens.map(([m, t]) =>
      `<div class="li" style="min-height:${h.toFixed(0)}px"><div class="m">${esc(m)}</div>
        <div class="t">${esc(t)}</div></div>`).join('') + `</div></div>`;
  },

  /* S39+ — fotos do mês */
  // a grade vem do spec (s.grade = [colunas, linhas]) para bater com a do PPTX; o
  // último slide do mês raramente fecha com 6 fotos, e julho/26 tem uma só
  fotos: s => head(s) + `<div class="s-body"><div class="fotos" style="grid-template-columns:repeat(${(s.grade||[3])[0]},1fr)">` +
    // f ja vem como data URI: as fotos nao existem como arquivo no site (repo e
    // site sao publicos), vem embutidas no spec, que sai do bucket privado.
    s.fotos.map(f => `<div class="f" style="background-image:url('${f}')"></div>`).join('') +
    `</div></div>`,

  pendente: s => head(s) + `<div class="s-body">` +
    vazio('◇', 'slide em aberto', s.titulo, s.sub || '', s.fonte, s.motivo) + `</div>`,
};

/* Estado vazio comum aos dois casos — cartão centralizado, e não um bloco de
   texto solto no meio do slide, que era como ficava antes. */
function vazio(glyph, tag, titulo, texto, fonte, motivo){
  return `<div class="vazio"><div class="box">
    <div class="glyph">${glyph}</div>
    <span class="tag">${esc(tag)}</span>
    <h2>${esc(titulo)}</h2>
    ${texto ? `<p>${esc(texto)}</p>` : ''}
    ${fonte ? `<div class="campos">
      <div><div class="l">Base que vai alimentar</div><div class="v">${esc(fonte)}</div></div>
      <div><div class="l">Por que ainda não tem</div><div class="v">${esc(motivo)}</div></div>
    </div>` : ''}
  </div></div>`;
}

/* ==================== conteúdo ao vivo (Supabase) ====================
   Comentários, exposições, manejo e fotos são "sem fonte" (ver cabeçalho de
   tools/build_comite.py) — editados pelo hub desde 31/08/2026, não mais só
   no _docs/comite_conteudo.json. O SPEC (spec.js) é o que o PIPELINE gerou
   da última vez que alguém rodou — pode estar atrasado em relação ao
   Supabase, que é a fonte principal agora. Pra editor E visualizador verem
   sempre o mais recente sem depender de alguém lembrar de rodar o build de
   novo, o deck busca o mês corrente no Supabase toda vez que troca de mês
   (ou salva algo) e RECONSTRÓI esses 4 tipos de slide na hora — mesma regra
   de slide_comentarios/slides_exposicoes/slide_manejo/_fotos_grupo_por_tema
   do build_comite.py, só que em JS. Se o Supabase não responder ou o mês
   não tiver linha lá, o SPEC (baked) continua valendo — não quebra nada. */
const FOTOS_POR_SLIDE = 6;
const GRADE_FOTOS = {1:[1,1], 2:[2,1], 3:[3,1], 4:[2,2], 5:[3,2], 6:[3,2]};
const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const ABR_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const TIPOS_EDITAVEIS = new Set(['comentarios', 'manejo', 'fotos', 'resultados']);
const ehExposicaoProg = s => s.t === 'tabela' && s.n === 23;
const ehEditavel = s => TIPOS_EDITAVEIS.has(s.t) || ehExposicaoProg(s);

function hubSb(){ try { return window.parent.HUB && window.parent.HUB.sb; } catch (e) { return null; } }
function hubEmail(){ try { return window.parent.HUB && window.parent.HUB.email; } catch (e) { return null; } }
function hubIsAdmin(){ try { return window.parent.HUB && window.parent.HUB.role === 'admin'; } catch (e) { return false; } }

let souEditor = false;
async function checaEditor(){
  if (hubIsAdmin()) {
    souEditor = true;
  } else {
    const sb = hubSb(), email = hubEmail();
    if (!sb || !email) {
      souEditor = false;
    } else {
      try {
        const { data } = await sb.from('comite_editores').select('email').eq('email', email).maybeSingle();
        souEditor = !!data;
      } catch (e) { souEditor = false; }
    }
  }
  const btn = document.getElementById('editar');
  if (btn) btn.hidden = !souEditor;
}

const conteudoCache = {};   // {mes: linha do Supabase | null} — invalidado ao salvar
async function buscaConteudoAoVivo(mes){
  if (mes in conteudoCache) return conteudoCache[mes];
  const sb = hubSb();
  if (!sb) { conteudoCache[mes] = null; return null; }
  try {
    const { data, error } = await sb.from('comite_conteudo').select('*').eq('mes', mes).maybeSingle();
    conteudoCache[mes] = error ? null : data;
  } catch (e) { conteudoCache[mes] = null; }
  return conteudoCache[mes];
}

/* foto do bucket privado -> data URI, mesmo formato que o resto do arquivo já
   usa (ver dataURI() mais abaixo, reaproveitada aqui) — assim a exportação
   PPTX continua funcionando sem mudar nada nela (ela espera base64 embutido,
   não URL remota). */
async function fotoDataUri(path){
  const sb = hubSb();
  if (!sb) return null;
  try {
    const { data, error } = await sb.storage.from('comite-fotos').createSignedUrl(path, 3600);
    if (error || !data) return null;
    return await dataURI(data.signedUrl);
  } catch (e) { return null; }
}

/* monta os slides sem fonte a partir do conteúdo ao vivo — mesma regra do
   Python. Devolve {comentarios:[...], exposicoes:[...], manejo:[...],
   fotos:[...]} só com as chaves que têm conteúdo; ausente = spec baked vale. */
async function montaSlidesAoVivo(mes){
  const c = await buscaConteudoAoVivo(mes);
  if (!c) return null;
  const [ano, mNum] = mes.split('-').map(Number);
  const out = {};

  if (c.comentarios && c.comentarios.length) {
    out.comentarios = [{t:'comentarios', n:8,
      titulo:`COMENTÁRIOS — VARIAÇÕES YTD JAN–${ABR_PT[mNum-1].toUpperCase()} ${ano}`,
      sub:'DRE 2026 | HPG · principais destaques acumulados por categoria',
      itens: c.comentarios}];
  }

  const exp = c.exposicoes || {}, prog = exp.programacao || [], res = exp.resultados || [];
  if (prog.length || res.length) {
    const s = [];
    if (prog.length) s.push({t:'tabela', n:23, titulo:`EXPOSIÇÕES ${ano} — PROGRAMAÇÃO`,
      sub:'Calendário de participações previstas', cols:['EVENTO','DATA','LOCAL','STATUS'], rows: prog});
    res.forEach((r, k) => s.push({t:'resultados', n:24+k, titulo:r.titulo, sub:r.sub || '', animais:r.animais}));
    out.exposicoes = s;
  }

  if (c.manejo && c.manejo.length) {
    out.manejo = [{t:'manejo', n:38, titulo:'MANEJO — PONTOS DE MELHORIA E DECISÕES', itens: c.manejo}];
  }

  if (c.fotos && c.fotos.length) {
    const grupos = typeof c.fotos[0] === 'string' ? [{tema:'', arquivos:c.fotos}] : c.fotos;
    const s = [];
    for (const g of grupos) {
      const urls = (await Promise.all((g.arquivos || []).map(fotoDataUri))).filter(Boolean);
      if (!urls.length) continue;
      const n = Math.ceil(urls.length / FOTOS_POR_SLIDE);
      for (let k = 0; k < n; k++) {
        const bloco = urls.slice(k * FOTOS_POR_SLIDE, (k + 1) * FOTOS_POR_SLIDE);
        const [cols, rows] = GRADE_FOTOS[bloco.length];
        let sub = g.tema ? `Obras e melhorias realizadas · ${g.tema}` : `Registros de ${MESES_PT[mNum-1]} ${ano}`;
        if (n > 1) sub += ` (${k+1}/${n})`;
        s.push({t:'fotos', n:39, titulo:'MANEJO — FOTOS E REGISTROS', sub, grade:[cols, rows], fotos:bloco});
      }
    }
    if (s.length) out.fotos = s;
  }
  return out;
}

/* troca, na lista `slides` já carregada, os slides de um tipo pelos novos —
   por tipo (comentarios/manejo/fotos) ou pelo par tabela(n=23)+resultados
   (exposições, que são dois tipos de slide pra 1 conteúdo só). */
function substituiSlidesDoTipo(chave, novos){
  let primeiro = -1;
  slides = slides.filter((s, i) => {
    const bate = chave === 'exposicoes' ? (ehExposicaoProg(s) || s.t === 'resultados') : s.t === chave;
    if (bate && primeiro === -1) primeiro = i;
    return !bate;
  });
  if (primeiro === -1) primeiro = slides.length;
  slides.splice(primeiro, 0, ...novos);
}

async function aplicaConteudoAoVivo(mes){
  const vivo = await montaSlidesAoVivo(mes);
  if (mes !== mesAtual || !vivo) return;   // usuário já trocou de mês, ou nada pra aplicar
  for (const [chave, novos] of Object.entries(vivo)) substituiSlidesDoTipo(chave, novos);
  listaSlides();
  go(Math.min(idx, slides.length - 1));
}

/* ==================== editor ====================
   Um painel só, corpo diferente por tipo. `estado` guarda os dados sendo
   editados (array de linhas, ou objeto — depende do tipo); os inputs
   escrevem direto em `estado` via listener delegado (sem re-render a cada
   tecla), e só reconstrói o HTML quando uma linha é adicionada/removida. */
let estado = null, tipoAtual = null;

function painelEditor(){
  let el = document.getElementById('editorOverlay');
  if (el) return el;
  el = document.createElement('div');
  el.id = 'editorOverlay';
  el.innerHTML = `<div id="editorPanel">
    <div class="ed-head"><h3 id="edTitulo"></h3>
      <button type="button" id="edFechar" aria-label="Fechar">✕</button></div>
    <div id="edCorpo"></div>
    <div class="ed-rodape">
      <span id="edStatus"></span>
      <button type="button" id="edCancelar">Cancelar</button>
      <button type="button" id="edSalvar" class="primary">Salvar</button>
    </div>
  </div>`;
  document.body.appendChild(el);
  el.addEventListener('click', e => { if (e.target === el) fechaEditor(); });
  document.getElementById('edFechar').onclick = fechaEditor;
  document.getElementById('edCancelar').onclick = fechaEditor;
  document.getElementById('edSalvar').onclick = salvaEditor;
  // delegado: linhas de comentários (data-i/data-f) escrevem direto em
  // `estado` sem precisar de handler próprio por input. Os outros tipos
  // (manejo/exposições/fotos) usam atributos diferentes (data-mi, data-pi,
  // data-gi...) com oninput própria montada em cada render — não passam
  // por aqui, o filtro abaixo já ignora.
  document.getElementById('edCorpo').addEventListener('input', e => {
    const t = e.target, i = t.dataset.i, f = t.dataset.f;
    if (i == null) return;
    estado[+i][f] = t.value;
  });
  return el;
}
function fechaEditor(){
  const el = document.getElementById('editorOverlay');
  if (el) el.style.display = 'none';
}

async function abreEditor(s){
  if (!souEditor || !ehEditavel(s)) return;
  const el = painelEditor();
  el.style.display = 'flex';
  document.getElementById('edStatus').textContent = 'carregando…';
  document.getElementById('edSalvar').disabled = true;
  const c = (await buscaConteudoAoVivo(mesAtual)) || {};

  if (s.t === 'comentarios') {
    tipoAtual = 'comentarios';
    estado = JSON.parse(JSON.stringify(c.comentarios || []));
    document.getElementById('edTitulo').textContent = `Comentários — ${SPEC.labels[mesAtual]}`;
    renderComentarios();
  } else if (s.t === 'manejo') {
    tipoAtual = 'manejo';
    estado = JSON.parse(JSON.stringify(c.manejo || []));
    document.getElementById('edTitulo').textContent = `Manejo — ${SPEC.labels[mesAtual]}`;
    renderManejo();
  } else if (ehExposicaoProg(s) || s.t === 'resultados') {
    tipoAtual = 'exposicoes';
    const exp = c.exposicoes || {};
    estado = {programacao: JSON.parse(JSON.stringify(exp.programacao || [])),
              resultados: JSON.parse(JSON.stringify(exp.resultados || []))};
    document.getElementById('edTitulo').textContent = `Exposições — ${SPEC.labels[mesAtual]}`;
    renderExposicoes();
  } else if (s.t === 'fotos') {
    tipoAtual = 'fotos';
    const fotos = c.fotos || [];
    estado = JSON.parse(JSON.stringify(
      typeof fotos[0] === 'string' ? [{tema:'', arquivos: fotos}] : fotos));
    document.getElementById('edTitulo').textContent = `Fotos — ${SPEC.labels[mesAtual]}`;
    await renderFotos();
  }
  document.getElementById('edStatus').textContent = '';
  document.getElementById('edSalvar').disabled = false;
}

/* ---- comentários: linhas {cat, txt, delta} ---- */
function renderComentarios(){
  const corpo = document.getElementById('edCorpo');
  corpo.innerHTML = estado.map((r, i) => `
    <div class="ed-linha">
      <input data-i="${i}" data-f="cat" value="${escAttr(r.cat)}" placeholder="Categoria">
      <textarea data-i="${i}" data-f="txt" placeholder="Texto do comentário" rows="2">${esc(r.txt || '')}</textarea>
      <input data-i="${i}" data-f="delta" value="${escAttr(r.delta)}" placeholder="+R$Xk" class="ed-curta">
      <button type="button" class="ed-rm" data-rm="${i}">✕</button>
    </div>`).join('') + `<button type="button" id="edAdd" class="ed-add">+ categoria</button>`;
  corpo.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rm, 1); renderComentarios(); });
  document.getElementById('edAdd').onclick = () => { estado.push({cat:'', txt:'', delta:''}); renderComentarios(); };
}

/* ---- manejo: linhas [mes, texto] ---- */
function renderManejo(){
  const corpo = document.getElementById('edCorpo');
  corpo.innerHTML = estado.map((r, i) => `
    <div class="ed-linha">
      <input data-mi="${i}" data-mf="0" value="${escAttr(r[0])}" placeholder="Mês" class="ed-curta">
      <textarea data-mi="${i}" data-mf="1" placeholder="Intervenções e decisões" rows="2">${esc(r[1] || '')}</textarea>
      <button type="button" class="ed-rm" data-rm="${i}">✕</button>
    </div>`).join('') + `<button type="button" id="edAdd" class="ed-add">+ mês</button>`;
  corpo.querySelectorAll('[data-mi]').forEach(inp => inp.oninput = () => {
    estado[+inp.dataset.mi][+inp.dataset.mf] = inp.value;
  });
  corpo.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rm, 1); renderManejo(); });
  document.getElementById('edAdd').onclick = () => { estado.push(['', '']); renderManejo(); };
}

/* ---- exposições: programação (linhas [evento,data,local,status]) +
   resultados ({titulo,sub,animais}, animais editado como texto "Nome:
   prêmio1; prêmio2" por linha — mais simples que formulário aninhado) ---- */
function animaisParaTexto(animais){
  return (animais || []).map(a => `${a.nome}: ${(a.premios || []).join('; ')}`).join('\n');
}
function textoParaAnimais(txt){
  return txt.split('\n').map(l => l.trim()).filter(Boolean).map(l => {
    const [nome, resto] = l.split(':');
    return {nome: (nome || '').trim(), premios: (resto || '').split(';').map(p => p.trim()).filter(Boolean)};
  });
}
function renderExposicoes(){
  const corpo = document.getElementById('edCorpo');
  const prog = estado.programacao.map((r, i) => `
    <div class="ed-linha ed-linha4">
      <input data-pi="${i}" data-pf="0" value="${escAttr(r[0])}" placeholder="Evento">
      <input data-pi="${i}" data-pf="1" value="${escAttr(r[1])}" placeholder="Data">
      <input data-pi="${i}" data-pf="2" value="${escAttr(r[2])}" placeholder="Local">
      <input data-pi="${i}" data-pf="3" value="${escAttr(r[3])}" placeholder="Status">
      <button type="button" class="ed-rm" data-rmp="${i}">✕</button>
    </div>`).join('');
  const res = estado.resultados.map((r, i) => `
    <div class="ed-bloco">
      <input data-ri="${i}" data-rf="titulo" value="${escAttr(r.titulo)}" placeholder="Título do slide">
      <input data-ri="${i}" data-rf="sub" value="${escAttr(r.sub)}" placeholder="Subtítulo">
      <textarea data-ri="${i}" data-rf="animais" rows="4" placeholder="Um animal por linha: Nome: prêmio 1; prêmio 2">${esc(animaisParaTexto(r.animais))}</textarea>
      <button type="button" class="ed-rm" data-rmr="${i}">✕ remover resultado</button>
    </div>`).join('');
  corpo.innerHTML = `<h4>Programação</h4>${prog}
    <button type="button" id="edAddP" class="ed-add">+ evento</button>
    <h4>Resultados</h4>${res}
    <button type="button" id="edAddR" class="ed-add">+ resultado</button>`;
  corpo.querySelectorAll('[data-pi]').forEach(inp => inp.oninput = () => {
    estado.programacao[+inp.dataset.pi][+inp.dataset.pf] = inp.value;
  });
  corpo.querySelectorAll('[data-rmp]').forEach(b => b.onclick = () => { estado.programacao.splice(+b.dataset.rmp, 1); renderExposicoes(); });
  corpo.querySelectorAll('[data-ri]').forEach(inp => inp.oninput = () => {
    const r = estado.resultados[+inp.dataset.ri], f = inp.dataset.rf;
    r[f === 'animais' ? '_txt' : f] = inp.value;
  });
  corpo.querySelectorAll('[data-rmr]').forEach(b => b.onclick = () => { estado.resultados.splice(+b.dataset.rmr, 1); renderExposicoes(); });
  document.getElementById('edAddP').onclick = () => { estado.programacao.push(['', '', '', '']); renderExposicoes(); };
  document.getElementById('edAddR').onclick = () => { estado.resultados.push({titulo:'', sub:'', animais:[]}); renderExposicoes(); };
}

/* ---- fotos: grupos {tema, arquivos[]} — upload direto no bucket ao
   escolher arquivo (não espera o Salvar geral, senão perde o arquivo se
   fechar sem salvar); remover/reordenar só mexe no array em memória até
   Salvar gravar a lista final. ---- */
async function renderFotos(){
  const corpo = document.getElementById('edCorpo');
  const blocos = await Promise.all(estado.map(async (g, gi) => {
    const thumbs = await Promise.all((g.arquivos || []).map(async (p, ai) => {
      const uri = await fotoDataUri(p);
      return `<div class="ed-foto">
        <img src="${uri || ''}" alt="">
        <div class="ed-foto-acoes">
          <button type="button" data-mv="${gi}:${ai}:-1" ${ai === 0 ? 'disabled' : ''}>↑</button>
          <button type="button" data-mv="${gi}:${ai}:1" ${ai === g.arquivos.length - 1 ? 'disabled' : ''}>↓</button>
          <button type="button" data-rmf="${gi}:${ai}">✕</button>
        </div></div>`;
    }));
    return `<div class="ed-grupo">
      <div class="ed-grupo-head">
        <input data-gi="${gi}" value="${escAttr(g.tema)}" placeholder="Tema (opcional)">
        <label class="ed-upload">+ foto<input type="file" accept="image/*" multiple data-up="${gi}" hidden></label>
        <button type="button" class="ed-rm" data-rmg="${gi}">✕ remover grupo</button>
      </div>
      <div class="ed-fotos-grade">${thumbs.join('') || '<span class="ed-vazio">sem foto neste grupo</span>'}</div>
    </div>`;
  }));
  corpo.innerHTML = blocos.join('') + `<button type="button" id="edAddG" class="ed-add">+ grupo (tema)</button>`;

  corpo.querySelectorAll('[data-gi]').forEach(inp => inp.oninput = () => { estado[+inp.dataset.gi].tema = inp.value; });
  corpo.querySelectorAll('[data-rmg]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rmg, 1); renderFotos(); });
  corpo.querySelectorAll('[data-rmf]').forEach(b => b.onclick = () => {
    const [gi, ai] = b.dataset.rmf.split(':').map(Number);
    estado[gi].arquivos.splice(ai, 1); renderFotos();
  });
  corpo.querySelectorAll('[data-mv]').forEach(b => b.onclick = () => {
    const [gi, ai, dir] = b.dataset.mv.split(':').map(Number);
    const arr = estado[gi].arquivos, novo = ai + dir;
    if (novo < 0 || novo >= arr.length) return;
    [arr[ai], arr[novo]] = [arr[novo], arr[ai]];
    renderFotos();
  });
  corpo.querySelectorAll('[data-up]').forEach(inp => inp.onchange = async () => {
    const gi = +inp.dataset.up;
    document.getElementById('edStatus').textContent = 'enviando foto…';
    for (const file of inp.files) {
      const path = await sobeFoto(file);
      if (path) estado[gi].arquivos.push(path);
    }
    document.getElementById('edStatus').textContent = '';
    renderFotos();
  });
  document.getElementById('edAddG').onclick = () => { estado.push({tema:'', arquivos:[]}); renderFotos(); };
}
async function sobeFoto(file){
  const sb = hubSb();
  if (!sb) return null;
  const nome = `${Date.now()}_${file.name}`.replace(/[^\w.-]/g, '_');
  const path = `${mesAtual}/${nome}`;
  try {
    const { error } = await sb.storage.from('comite-fotos').upload(path, file, {contentType: file.type, upsert: true});
    if (error) { alert('Falha no upload: ' + error.message); return null; }
    return path;
  } catch (e) { alert('Falha no upload: ' + e.message); return null; }
}

async function salvaEditor(){
  const sb = hubSb();
  if (!sb) { alert('Sem sessão do hub — não dá pra salvar.'); return; }
  document.getElementById('edStatus').textContent = 'salvando…';
  document.getElementById('edSalvar').disabled = true;
  let coluna, valor;
  if (tipoAtual === 'comentarios') { coluna = 'comentarios'; valor = estado; }
  else if (tipoAtual === 'manejo') { coluna = 'manejo'; valor = estado; }
  else if (tipoAtual === 'fotos') { coluna = 'fotos'; valor = estado; }
  else if (tipoAtual === 'exposicoes') {
    coluna = 'exposicoes';
    valor = {
      programacao: estado.programacao,
      resultados: estado.resultados.map(r => ({titulo: r.titulo, sub: r.sub, animais: textoParaAnimais(r._txt ?? animaisParaTexto(r.animais))})),
    };
  }
  const { error } = await sb.from('comite_conteudo').upsert({mes: mesAtual, [coluna]: valor});
  document.getElementById('edSalvar').disabled = false;
  if (error) { document.getElementById('edStatus').textContent = ''; alert('Falha ao salvar: ' + error.message); return; }
  delete conteudoCache[mesAtual];
  fechaEditor();
  aplicaConteudoAoVivo(mesAtual);
}

function escAttr(s){ return esc(s).replace(/"/g, '&quot;'); }

/* ---- navegação ---- */
let mesAtual = SPEC.padrao;
/* mensal = deck inteiro; trimestral = sem os slides que recortam o MÊS.
   O comitê trimestral olha o acumulado do ANO, não o do trimestre — então os
   slides de YTD, que já existem, são justamente os que ficam. */
let modo = 'mensal';
const doModo = d => modo === 'trimestral' ? d.filter(x => !x.so_mensal) : d;
let slides = doModo(SPEC.decks[mesAtual]);
let idx = 0;
const stage = document.getElementById('stage');

function render(){
  const s = slides[idx], n = slides.length;
  const body = (R[s.t] || R.pendente)(s);
  stage.innerHTML = `<div class="slide" id="slide">${body}${s.t === 'capa' || s.t === 'encerramento' ? '' : foot(s, idx, n)}</div>`;
  fit();
  document.getElementById('pos').textContent = `${idx + 1} / ${n}`;
  document.getElementById('prev').disabled = idx === 0;
  document.getElementById('next').disabled = idx === n - 1;
  const sel = document.getElementById('ir');
  if (sel.value !== String(idx)) sel.value = String(idx);
  location.hash = `#${mesAtual}/${idx + 1}`;
  const btnEditar = document.getElementById('editar');
  if (btnEditar) btnEditar.disabled = !ehEditavel(s);
}
/* escala o slide de 1280×720 pra caber na área disponível, mantendo a proporção.
   translate(-50%,-50%) faz parte do transform (não só do CSS base) porque
   setar style.transform aqui SUBSTITUI o valor inteiro — só "scale(...)"
   perdia a centralização do CSS. Achado em 31/08/2026: com place-items:center
   (grid), o slide de 1280px dentro de um stage de celular (bem menor) saía
   fora da tela — grid/flex têm min-width:auto implícito no filho, que recusa
   encolher o item pra centralizar certo quando ele é muito maior que o
   container. position:absolute + translate(-50%,-50%) centraliza certo
   sempre, não depende do tamanho relativo dos dois. */
function fit(){
  const el = document.getElementById('slide');
  if (!el) return;
  const box = stage.getBoundingClientRect();
  const pad = document.body.classList.contains('play') ? 0 : 32;
  const s = Math.min((box.width - pad) / 1280, (box.height - pad) / 720);
  el.style.transform = `translate(-50%,-50%) scale(${s})`;
}
const go = i => { idx = Math.max(0, Math.min(slides.length - 1, i)); render(); };

function trocaMes(mes){
  if (!SPEC.decks[mes]) return;
  mesAtual = mes;
  slides = doModo(SPEC.decks[mes]);
  document.getElementById('mes').value = mes;
  listaSlides();
  atualizaAviso();
  go(Math.min(idx, slides.length - 1));
  aplicaConteudoAoVivo(mes);   // async, de propósito — reflete quando chegar, sem travar a troca
}
function listaSlides(){
  document.getElementById('ir').innerHTML = slides
    .map((s, i) => `<option value="${i}">${String(i + 1).padStart(2, '0')} · ${esc((s.titulo || s.t).slice(0, 52))}</option>`).join('');
}
function atualizaAviso(){
  const p = slides.filter(s => s.t === 'pendente').length;
  document.getElementById('aviso').textContent =
    (p ? `${p} de ${slides.length} slides sem fonte` : 'todos os slides com fonte') +
    (SPEC.avisos && SPEC.avisos.length ? ` · ${SPEC.avisos[0]}` : '');
}

window.addEventListener('resize', fit);
document.addEventListener('keydown', e => {
  if (e.target.tagName === 'SELECT') return;
  if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') { go(idx + 1); e.preventDefault(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { go(idx - 1); e.preventDefault(); }
  else if (e.key === 'Home') go(0);
  else if (e.key === 'End') go(slides.length - 1);
  else if (e.key === 'Escape' && document.body.classList.contains('play')) sairPlay();
  else if (e.key.toLowerCase() === 'p') play();
});

function play(){
  document.body.classList.add('play');
  (document.documentElement.requestFullscreen ? document.documentElement.requestFullscreen() : Promise.resolve())
    .catch(() => {});
  fit();
}
function sairPlay(){
  document.body.classList.remove('play');
  if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  fit();
}
document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement) sairPlay(); });

/* ---- export PPTX (mesmo spec, outra saída) ---- */
async function exportarPptx(btn){
  btn.disabled = true; btn.textContent = 'Gerando…';
  try {
    const p = new PptxGenJS();
    p.defineLayout({name:'HPG', width:10, height:5.625});      // 16:9, igual ao deck original
    p.layout = 'HPG';
    p.title = `Relatório de Desempenho Estratégico — ${SPEC.labels[mesAtual]}`;
    const logo = await dataURI(LOGO);
    // fotos precisam virar base64 antes: o pptxgen não busca arquivo sozinho
    const imgs = {};
    for (const s of slides) for (const f of (s.fotos || [])) {
      // f JA e data URI (fotos vem embutidas no spec, nao como arquivo)
      if (!(f in imgs)) imgs[f] = f;
    }
    slides.forEach((s, i) => pptSlide(p, s, i, logo, imgs));
    const rotulo = modo === 'trimestral' ? 'TRIMESTRAL' : 'MENSAL';
    await p.writeFile({fileName: `RELATORIO ${rotulo}_PG_${mesAtual}.pptx`});
    btn.textContent = 'Exportar PPTX';
  } catch (e) {
    btn.textContent = 'Falhou — ver console';
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}
/* o pptxgen precisa da imagem em base64; o logo é do mesmo diretório, fetch resolve */
async function dataURI(url){
  try {
    const b = await (await fetch(url)).blob();
    return await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); });
  } catch { return null; }
}

function pptSlide(p, s, i, logo, imgs){
  const sl = p.addSlide();
  sl.background = {color: C.bg};
  const T = (t, o) => sl.addText(t, Object.assign({fontFace:'Segoe UI', color:C.ink}, o));

  if (s.t === 'capa' || s.t === 'encerramento'){
    if (logo) sl.addImage({data:logo, x:6.1, y:1.1, w:2.6, h:2.6});
    T(s.titulo, {x:0.6, y:s.t === 'capa' ? 1.9 : 2.4, w:5.4, h:1.1, fontSize:s.t === 'capa' ? 28 : 26,
                 bold:true, color:s.t === 'capa' ? C.ink : C.amber});
    if (s.mes) T(s.mes, {x:0.6, y:3.0, w:5.4, h:0.4, fontSize:15, color:C.amber});
    if (s.org) T(s.org, {x:0.6, y:3.8, w:5.4, h:0.4, fontSize:10, color:C.ink3, charSpacing:3});
    return;
  }
  if (s.t === 'divisor'){
    T('0' + s.n, {x:6.6, y:0.4, w:2.9, h:1.4, fontSize:72, bold:true, color:C.line, align:'right'});
    T(s.titulo, {x:0.6, y:2.1, w:6, h:0.8, fontSize:26, bold:true});
    T(s.sub, {x:0.6, y:2.95, w:8.8, h:0.4, fontSize:12, color:C.ink3});
    return;
  }

  /* título comprido quebrava em duas linhas, cobria o subtítulo e a tabela — que
     começa em y fixo — não recuava. Reduz a fonte quando é longo, em vez de deixar
     quebrar. */
  const tituloLongo = (s.titulo || '').length > 58;
  T(s.titulo, {x:0.41, y:0.13, w:8.6, h:0.4, fontSize:tituloLongo ? 12 : 15,
               bold:true, valign:'top'});
  if (s.sub) T(s.sub, {x:0.41, y:0.52, w:8.6, h:0.3, fontSize:9, color:C.ink3});
  if (logo) sl.addImage({data:logo, x:9.05, y:0.14, w:0.55, h:0.55});
  T(`${SPEC.labels[mesAtual]}   ·   ${i + 1}/${slides.length}${s.obs ? '   ·   ' + s.obs : ''}`,
    {x:0.41, y:5.25, w:9.2, h:0.25, fontSize:8, color:C.ink3});

  /* rowH acompanha o nº de linhas, como no HTML — 4,1 in é a altura útil */
  /* rowH no pptxgen é altura MÍNIMA: o PowerPoint expande a linha para caber o
     texto mais a margem interna da célula. Com a margem padrão (~0,05in em cima e
     embaixo) uma linha de 6pt renderiza ~0,19in, não os 0,155in declarados — 27
     linhas viravam 5,1in a partir de y=0,95 e a tabela saía pela borda de baixo,
     por cima do rodapé. Era isso que quebrava os slides de DRE.
     Margem zerada + altura calculada com o piso REAL de renderização. */
  /* `margin` aqui é por CÉLULA e em PONTOS. Passar 3 nas laterais deixava a largura
     útil de coluna estreita menor que um caractere, e o PowerPoint quebrava a cada
     letra — a tabela virava uma coluna de caracteres empilhados e empurrava o resto
     para fora do slide. Margem vertical zerada resolve a altura sem estrangular a
     largura; a lateral fica no padrão. */
  const tbl = (rows, opts) => {
    const y = (opts && opts.y) || 0.95;
    const alt = 5.15 - y;
    const fs = Math.max(6, Math.min(9, alt / rows.length * 26));
    return sl.addTable(rows, Object.assign({
      x:0.41, y, w:9.2, border:{type:'solid', pt:0.4, color:C.line},
      fontFace:'Segoe UI', fontSize:fs,
      rowH:Math.max(fs * 1.2 / 72, Math.min(0.3, alt / rows.length)),
      color:C.ink, valign:'middle',
    }, opts));
  };
  const th = t => ({text:String(t), options:{bold:true, fontSize:7.5, color:C.ink3, fill:{color:C.bg2}, align:'right'}});
  const kpis = () => {
    const w = 9.2 / s.kpis.length;
    s.kpis.forEach((k, j) => {
      const x = 0.41 + j * w;
      sl.addShape(p.ShapeType.roundRect, {x, y:0.95, w:w - 0.14, h:0.82,
        fill:{color:C.card}, line:{color:C.line, width:0.5}, rectRadius:0.05});
      T(k.v, {x:x + 0.13, y:1.0, w:w - 0.4, h:0.34, fontSize:17, bold:true, color:C.amber});
      T(k.l, {x:x + 0.13, y:1.34, w:w - 0.4, h:0.22, fontSize:8, color:C.ink2});
      T(k.s, {x:x + 0.13, y:1.53, w:w - 0.4, h:0.2, fontSize:7, color:C.ink3});
    });
  };

  if (s.t === 'agenda'){
    s.itens.forEach((it, k) => {
      const x = 0.41 + k * 1.83;
      sl.addShape(p.ShapeType.rect, {x, y:1.15, w:1.7, h:0.04, fill:{color:C.amber}});
      T(it.n, {x, y:1.3, w:1.7, h:0.4, fontSize:18, bold:true, color:C.amber});
      T(it.titulo, {x, y:1.8, w:1.7, h:0.35, fontSize:10, bold:true});
      T(it.sub, {x, y:2.2, w:1.7, h:0.7, fontSize:8, color:C.ink3});
    });
    return;
  }
  if (s.t === 'dre'){
    const rows = [[th('NATUREZA'), th('ORÇADO'), th('REALIZADO'), th('∆ R$ k'), th('∆ %')]];
    s.linhas.forEach(l => rows.push([
      {text:(l.total ? '' : '   ') + l.nome, options:{align:'left', bold:l.total, color:l.total ? C.amber : C.ink2}},
      {text:rs(l.v[0]), options:{align:'right', bold:l.total}},
      {text:rs(l.v[1]), options:{align:'right', bold:l.total}},
      {text:dk(l.v[2]), options:{align:'right', color:l.v[2] == null ? C.ink : l.v[2] >= 0 ? C.pos : C.neg}},
      {text:dpct(l.v[3]), options:{align:'right', color:l.v[3] == null ? C.ink3 : l.v[3] >= 0 ? C.pos : C.neg}},
    ]));
    tbl(rows, {colW:[3.7, 1.7, 1.7, 1.1, 1.0]});
    return;
  }
  if (s.t === 'tabela'){
    if (!s.rows.length){ T('Nenhum registro nesse recorte.', {x:0.41, y:2.4, w:8, h:0.4, fontSize:12, color:C.ink3}); return; }
    const rows = [s.cols.map(th)];
    s.rows.forEach(r => rows.push(r.map((c, j) => ({
      text:(s.moeda || []).includes(j) ? rs(c) : (s.data || []).includes(j) ? brdata(c) : String(c == null ? '—' : c),
      options:{align:j === 0 ? 'left' : 'right'}}))));
    const w0 = 9.2 / s.cols.length;
    tbl(rows, {colW:[w0 * 1.6].concat(new Array(s.cols.length - 1).fill((9.2 - w0 * 1.6) / (s.cols.length - 1)))});
    return;
  }
  if (s.t === 'kpis_tabela' || s.t === 'matriz'){
    kpis();
    if (s.t === 'kpis_tabela'){
      const rows = [s.tabela.cols.map(th)];
      s.tabela.rows.forEach(r => rows.push(r.map((c, j) =>
        ({text:String(c), options:{align:j === 0 ? 'left' : 'right'}}))));
      const w0 = 9.2 / s.tabela.cols.length;
      tbl(rows, {y:1.95, colW:[w0 * 1.7].concat(new Array(s.tabela.cols.length - 1)
        .fill((9.2 - w0 * 1.7) / (s.tabela.cols.length - 1)))});
    } else {
      const rows = [s.cols.map(th)];
      s.rows.forEach(r => rows.push(r.map((c, j) => j === 0
        ? {text:String(c), options:{align:'left', bold:/Saldo/.test(c), color:/Saldo/.test(c) ? C.amber : C.ink2}}
        : {text:c ? rsk(c) : '—', options:{align:'right', color:typeof c === 'number' && c < 0 ? C.neg : C.ink}})));
      tbl(rows, {y:1.95, colW:[1.9].concat(new Array(s.cols.length - 1).fill((9.2 - 1.9) / (s.cols.length - 1)))});
    }
    return;
  }
  if (s.t === 'lista_mes'){
    let y = 0.95;
    const passo = Math.min(0.22, 4.2 / s.meses.reduce((a, m) => a + 1 + m.itens.length, 0));
    s.meses.forEach(m => {
      T(m.mes, {x:0.41, y, w:0.9, h:passo, fontSize:9, bold:true, color:C.amber});
      T('ANIMAIS E PRODUTOS', {x:1.35, y, w:3, h:passo, fontSize:7.5, color:C.ink3});
      T(rs(m.total), {x:7.6, y, w:2.0, h:passo, fontSize:10, bold:true, align:'right'});
      y += passo + 0.04;
      m.itens.forEach(it => {
        T(it.desc, {x:1.35, y, w:6.1, h:passo, fontSize:8, color:C.ink2});
        T(rs(it.valor), {x:7.6, y, w:2.0, h:passo, fontSize:8, align:'right', color:C.ink2});
        y += passo;
      });
      y += 0.04;
    });
    return;
  }
  if (s.t === 'comentarios'){
    const rows = [[th('CATEGORIA'), th('DESTAQUE DO ACUMULADO'), th('∆')]];
    s.itens.forEach(it => rows.push([
      {text:it.cat, options:{align:'left', bold:true, color:C.amber}},
      {text:it.txt, options:{align:'left', color:C.ink2}},
      {text:it.delta, options:{align:'right', bold:true, color:/^[-−]/.test(it.delta) ? C.neg : C.pos}},
    ]));
    tbl(rows, {colW:[1.9, 6.3, 1.0], rowH:Math.min(0.42, 4.1 / rows.length), fontSize:8, valign:'top'});
    return;
  }
  if (s.t === 'resultados'){
    // duas colunas, quebrando por altura acumulada
    const col = [[], []];
    const alt = a => 0.26 + a.premios.length * 0.24;
    let h0 = 0;
    s.animais.forEach(a => { const k = h0 <= 2.0 ? 0 : 1; col[k].push(a); if (!k) h0 += alt(a); });
    col.forEach((lista, k) => {
      let y = 0.95;
      lista.forEach(a => {
        T(a.nome, {x:0.41 + k * 4.7, y, w:4.4, h:0.24, fontSize:10, bold:true, color:C.amber});
        y += 0.26;
        a.premios.forEach(pr => {
          T('🏆  ' + pr, {x:0.55 + k * 4.7, y, w:4.3, h:0.24, fontSize:9, color:C.ink});
          y += 0.24;
        });
        y += 0.08;
      });
    });
    return;
  }
  if (s.t === 'manejo'){
    const rows = [[th('MÊS'), th('INTERVENÇÕES E DECISÕES')]];
    s.itens.forEach(([m, t]) => rows.push([
      {text:m, options:{align:'left', bold:true, color:C.amber}},
      {text:t, options:{align:'left', color:C.ink}},
    ]));
    tbl(rows, {colW:[0.9, 8.3], rowH:Math.min(0.6, 4.1 / rows.length), fontSize:8.5, valign:'top'});
    return;
  }
  if (s.t === 'fotos'){
    // mesma grade do HTML, vinda do spec — não recalcular aqui, senão as duas saídas
    // divergem quando o mês fecha com menos de 6 fotos
    const [cols, linhas] = s.grade || [3, 2];
    const gap = 0.12;
    const w = (9.2 - gap * (cols - 1)) / cols;
    const alt = (4.15 - gap * (linhas - 1)) / linhas;
    s.fotos.forEach((f, k) => {
      const d = imgs && imgs[f];
      if (!d) return;
      sl.addImage({data:d, x:0.41 + (k % cols) * (w + gap), y:0.95 + Math.floor(k / cols) * (alt + gap),
                   w, h:alt, sizing:{type:'cover', w, h:alt}});
    });
    return;
  }
  // pendente
  T('SEM DADO AINDA', {x:0.41, y:1.6, w:3, h:0.3, fontSize:9, bold:true, color:C.amber, charSpacing:2});
  T('BASE QUE VAI ALIMENTAR', {x:0.41, y:2.2, w:8.6, h:0.22, fontSize:8, color:C.ink3, charSpacing:1});
  T(s.fonte, {x:0.41, y:2.44, w:8.6, h:0.5, fontSize:10});
  T('POR QUE AINDA NÃO TEM', {x:0.41, y:3.1, w:8.6, h:0.22, fontSize:8, color:C.ink3, charSpacing:1});
  T(s.motivo, {x:0.41, y:3.34, w:8.6, h:0.5, fontSize:10});
}

/* ---- boot ---- */
document.getElementById('mes').innerHTML = SPEC.meses
  .map(m => `<option value="${m}">${esc(SPEC.labels[m])}</option>`).join('');
document.getElementById('mes').onchange = e => trocaMes(e.target.value);
document.getElementById('modo').onchange = e => {
  modo = e.target.value;
  idx = 0;                       // a numeração muda; voltar ao início evita cair fora
  trocaMes(mesAtual);
};
document.getElementById('prev').onclick = () => go(idx - 1);
document.getElementById('next').onclick = () => go(idx + 1);
document.getElementById('ir').onchange = e => go(+e.target.value);
document.getElementById('play').onclick = play;
document.getElementById('pptx').onclick = e => exportarPptx(e.currentTarget);

document.getElementById('editar').onclick = () => abreEditor(slides[idx]);

const alvo = /^#([\d-]+)\/(\d+)$/.exec(location.hash);
if (alvo && SPEC.decks[alvo[1]]) { idx = +alvo[2] - 1; mesAtual = alvo[1]; }
trocaMes(mesAtual);
checaEditor();
