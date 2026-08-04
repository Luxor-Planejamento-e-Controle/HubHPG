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
  fotos: s => head(s) + `<div class="s-body"><div class="fotos">` +
    s.fotos.map(f => `<div class="f" style="background-image:url('assets/comite/${f}')"></div>`).join('') +
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

/* ---- navegação ---- */
let mesAtual = SPEC.padrao;
let slides = SPEC.decks[mesAtual];
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
}
/* escala o slide de 1280×720 pra caber na área disponível, mantendo a proporção */
function fit(){
  const el = document.getElementById('slide');
  if (!el) return;
  const box = stage.getBoundingClientRect();
  const pad = document.body.classList.contains('play') ? 0 : 32;
  el.style.transform = `scale(${Math.min((box.width - pad) / 1280, (box.height - pad) / 720)})`;
}
const go = i => { idx = Math.max(0, Math.min(slides.length - 1, i)); render(); };

function trocaMes(mes){
  if (!SPEC.decks[mes]) return;
  mesAtual = mes;
  slides = SPEC.decks[mes];
  document.getElementById('mes').value = mes;
  listaSlides();
  atualizaAviso();
  go(Math.min(idx, slides.length - 1));
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
      if (!(f in imgs)) imgs[f] = await dataURI('assets/comite/' + f);
    }
    slides.forEach((s, i) => pptSlide(p, s, i, logo, imgs));
    await p.writeFile({fileName: `RELATORIO MENSAL_PG_${mesAtual}.pptx`});
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

  T(s.titulo, {x:0.41, y:0.13, w:8.6, h:0.4, fontSize:15, bold:true});
  if (s.sub) T(s.sub, {x:0.41, y:0.52, w:8.6, h:0.3, fontSize:9, color:C.ink3});
  if (logo) sl.addImage({data:logo, x:9.05, y:0.14, w:0.55, h:0.55});
  T(`${SPEC.labels[mesAtual]}   ·   ${i + 1}/${slides.length}${s.obs ? '   ·   ' + s.obs : ''}`,
    {x:0.41, y:5.25, w:9.2, h:0.25, fontSize:8, color:C.ink3});

  /* rowH acompanha o nº de linhas, como no HTML — 4,1 in é a altura útil */
  const tbl = (rows, opts) => {
    const y = (opts && opts.y) || 0.95;
    const alt = 5.15 - y;
    return sl.addTable(rows, Object.assign({
      x:0.41, y, w:9.2, border:{type:'solid', pt:0.4, color:C.line},
      fontFace:'Segoe UI', fontSize:Math.max(6, Math.min(9, alt / rows.length * 26)),
      rowH:Math.max(0.11, Math.min(0.3, alt / rows.length)),
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
document.getElementById('prev').onclick = () => go(idx - 1);
document.getElementById('next').onclick = () => go(idx + 1);
document.getElementById('ir').onchange = e => go(+e.target.value);
document.getElementById('play').onclick = play;
document.getElementById('pptx').onclick = e => exportarPptx(e.currentTarget);

const alvo = /^#([\d-]+)\/(\d+)$/.exec(location.hash);
if (alvo && SPEC.decks[alvo[1]]) { idx = +alvo[2] - 1; mesAtual = alvo[1]; }
trocaMes(mesAtual);
