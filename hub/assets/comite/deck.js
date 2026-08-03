/* Comitê Mensal HPG — renderiza o spec como deck e exporta o mesmo spec em PPTX.

   Regra do desenho: o spec (assets/comite/spec.js, gerado por
   hub/tools/build_comite.py) é a ÚNICA fonte. HTML e PPTX são duas saídas do
   mesmo objeto — se um slide muda, muda nos dois, porque nenhum dos dois guarda
   conteúdo próprio. Slide sem fonte de dado vira `pendente` e diz isso na tela;
   número não se inventa pra encher slide. */
'use strict';

const SPEC = window.COMITE_SPEC;
const C = {bg:'04223B', bg2:'072B49', card:'0A3050', line:'1B486B', ink:'EAF0F4',
           ink2:'B6C8D6', ink3:'93AABC', amber:'CA9703', pos:'4CC38A', neg:'F07A7A'};
const LOGO = 'assets/pg-logo.png';

/* ---- formatação ---- */
const nf = (v, d) => v.toLocaleString('pt-BR', {minimumFractionDigits:d, maximumFractionDigits:d});
const rs = v => v == null ? '—' : (v < 0 ? '-' : '') + 'R$ ' + nf(Math.abs(v), 0);
const dk = v => v == null ? '—' : (v >= 0 ? '+' : '−') + nf(Math.abs(v), 0) + 'k';
/* a coluna ∆% vem como fração (0,86) na planilha; quando é texto ('N/A') o
   build já mandou null */
const dpct = v => v == null ? 'N/A' : (v >= 0 ? '+' : '−') + nf(Math.abs(v) * (Math.abs(v) <= 5 ? 100 : 1), 0) + '%';
const cls = v => v == null ? '' : v >= 0 ? 'pos' : 'neg';
const esc = s => String(s == null ? '' : s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));

/* ---- render HTML ---- */
const head = s => `<div class="s-head"><h1>${esc(s.titulo)}</h1>${s.sub ? `<div class="sub">${esc(s.sub)}</div>` : ''}</div>
  <img class="s-mark" src="${LOGO}" alt="">`;
const foot = (s, i, n) => `<div class="s-foot"><span>${esc(SPEC.mesLabel)}</span><span class="spacer"></span><span>${i + 1} / ${n}</span></div>`;

const R = {
  capa: s => `<div class="capa"><img src="${LOGO}" alt="">
    <h1>${esc(s.titulo)}</h1><div class="mes">${esc(s.mes)}</div><div class="org">${esc(s.org)}</div></div>`,

  encerramento: s => `<div class="fim"><img src="${LOGO}" alt=""><h1>${esc(s.titulo)}</h1></div>`,

  divisor: s => `<div class="divisor"><div class="n">0${s.n}</div>
    <h1>${esc(s.titulo)}</h1><div class="sub">${esc(s.sub)}</div></div>`,

  agenda: s => head(s) + `<div class="s-body"><div class="agenda">` +
    s.itens.map(it => `<div class="it"><div class="n">${esc(it.n)}</div>
      <h3>${esc(it.titulo)}</h3><p>${esc(it.sub)}</p></div>`).join('') + `</div></div>`,

  /* S04/S07/S10/S13/S14 — Orçado | Realizado | ∆ R$ k | ∆ % */
  dre: s => {
    const h = 610 / Math.max(s.linhas.length, 1);
    return head(s) + `<div class="s-body"><table class="t" style="font-size:${h < 22 ? 12 : 13}px">
      <colgroup><col style="width:40%"><col><col><col style="width:12%"><col style="width:10%"></colgroup>
      <thead><tr><th>NATUREZA</th><th>ORÇADO</th><th>REALIZADO</th><th>∆ R$ k</th><th>∆ %</th></tr></thead>
      <tbody>${s.linhas.map(l => `<tr class="${l.total ? 'tot' : ''}" style="height:${Math.min(h, 26)}px">
        <td class="nome ${l.total ? '' : 'filho'}">${esc(l.nome)}</td>
        <td>${rs(l.v[0])}</td><td>${rs(l.v[1])}</td>
        <td class="${cls(l.v[2])}">${dk(l.v[2])}</td>
        <td class="${cls(l.v[3])}">${dpct(l.v[3])}</td></tr>`).join('')}</tbody></table></div>`;
  },

  /* S11 — KPIs + tabela de composição */
  kpis_tabela: s => head(s) + `<div class="s-body">
    <div class="kpis" style="grid-template-columns:repeat(${s.kpis.length},1fr)">
      ${s.kpis.map(k => `<div class="kpi"><div class="v">${esc(k.v)}</div><div class="l">${esc(k.l)}</div><div class="s">${esc(k.s)}</div></div>`).join('')}
    </div>
    <table class="t" style="font-size:13px"><colgroup><col style="width:56%"><col><col style="width:18%"></colgroup>
      <thead><tr>${s.tabela.cols.map((c, i) => `<th${i === 0 ? '' : ''}>${esc(c)}</th>`).join('')}</tr></thead>
      <tbody>${s.tabela.rows.map(r => `<tr style="height:26px">${r.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`,

  /* S12 — KPIs + matriz título × meses */
  matriz: s => head(s) + `<div class="s-body">
    <div class="kpis" style="grid-template-columns:repeat(${s.kpis.length},1fr)">
      ${s.kpis.map(k => `<div class="kpi"><div class="v">${esc(k.v)}</div><div class="l">${esc(k.l)}</div><div class="s">${esc(k.s)}</div></div>`).join('')}
    </div>
    <table class="t" style="font-size:12.5px"><thead><tr>${s.cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead>
      <tbody>${s.rows.map(r => {
        const tot = /Saldo/.test(r[0]);
        return `<tr class="${tot ? 'tot' : ''}" style="height:30px">` + r.map((c, i) =>
          i === 0 ? `<td class="nome">${esc(c)}</td>`
                  : `<td class="${typeof c === 'number' && c ? cls(c) : ''}">${c ? rsk(c) : '—'}</td>`).join('') + `</tr>`;
      }).join('')}</tbody></table></div>`,

  /* S09 — investimentos mês a mês */
  lista_mes: s => head(s) + `<div class="s-body"><div class="lista">` +
    s.meses.map(m => `<div class="m"><span class="mes">${esc(m.mes)}</span>
        <span class="tag">Animais e produtos</span><span class="tot">${rs(m.total)}</span></div>` +
      m.itens.map(it => `<div class="it"><span></span><span class="d">${esc(it.desc)}</span><span>${rs(it.valor)}</span></div>`).join('')
    ).join('') + `</div></div>`,

  pendente: s => head(s) + `<div class="s-body"><div class="pend">
      <span class="tag">sem dado ainda</span>
      <div class="campo"><div class="l">Base que vai alimentar</div><div class="v">${esc(s.fonte)}</div></div>
      <div class="campo"><div class="l">Por que ainda não tem</div><div class="v">${esc(s.motivo)}</div></div>
    </div></div>`,
};
/* R$ k, como no deck original: R$18.703k */
function rsk(v){
  if (v == null || v === 0) return '—';
  const s = v < 0 ? '-' : '', a = Math.abs(v) / 1000;
  return s + 'R$' + nf(a, a < 100 ? 1 : 0) + 'k';
}

/* ---- navegação ---- */
let idx = 0;
const stage = document.getElementById('stage');
const total = SPEC.slides.length;

function render(){
  const s = SPEC.slides[idx];
  const body = (R[s.t] || R.pendente)(s);
  stage.innerHTML = `<div class="slide" id="slide">${body}${s.t === 'capa' || s.t === 'encerramento' ? '' : foot(s, idx, total)}</div>`;
  fit();
  document.getElementById('pos').textContent = `${idx + 1} / ${total}`;
  document.getElementById('prev').disabled = idx === 0;
  document.getElementById('next').disabled = idx === total - 1;
  const sel = document.getElementById('ir');
  if (sel.value !== String(idx)) sel.value = String(idx);
  location.hash = '#' + (idx + 1);
}
/* escala o slide de 1280×720 pra caber na área disponível, mantendo a proporção */
function fit(){
  const el = document.getElementById('slide');
  if (!el) return;
  const box = stage.getBoundingClientRect();
  const pad = document.body.classList.contains('play') ? 0 : 32;
  el.style.transform = `scale(${Math.min((box.width - pad) / 1280, (box.height - pad) / 720)})`;
}
const go = i => { idx = Math.max(0, Math.min(total - 1, i)); render(); };

window.addEventListener('resize', fit);
document.addEventListener('keydown', e => {
  if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') { go(idx + 1); e.preventDefault(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { go(idx - 1); e.preventDefault(); }
  else if (e.key === 'Home') go(0);
  else if (e.key === 'End') go(total - 1);
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
const IN = {W: 10, H: 5.625};                 // 16:9, igual ao deck original
const px = v => v / 128;                       // 1280 px de largura -> 10 in

async function exportarPptx(btn){
  btn.disabled = true; btn.textContent = 'Gerando…';
  try {
    const p = new PptxGenJS();
    p.defineLayout({name:'HPG', width:IN.W, height:IN.H});
    p.layout = 'HPG';
    p.title = `Relatório de Desempenho Estratégico — ${SPEC.mesLabel}`;
    const logo = await dataURI(LOGO);
    SPEC.slides.forEach((s, i) => pptSlide(p, s, i, logo));
    await p.writeFile({fileName: `RELATORIO MENSAL_PG_${SPEC.mes}.pptx`});
    btn.textContent = 'Exportar PPTX';
  } catch (e) {
    btn.textContent = 'Falhou — ver console';
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}
/* o pptxgen precisa da imagem em base64; file:// não deixa ler canvas de outra
   origem, mas o logo é do mesmo diretório, então fetch resolve */
async function dataURI(url){
  try {
    const b = await (await fetch(url)).blob();
    return await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); });
  } catch { return null; }
}

function pptSlide(p, s, i, logo){
  const sl = p.addSlide();
  sl.background = {color: C.bg};
  const T = (t, o) => sl.addText(t, Object.assign({fontFace:'Segoe UI', color:C.ink}, o));

  if (s.t === 'capa' || s.t === 'encerramento'){
    if (logo) sl.addImage({data:logo, x:6.1, y:1.1, w:2.6, h:2.6});
    T(s.titulo, {x:0.6, y:s.t === 'capa' ? 1.9 : 2.4, w:5.4, h:1.1, fontSize:s.t === 'capa' ? 28 : 26, bold:true,
                 color:s.t === 'capa' ? C.ink : C.amber});
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

  // cabeçalho padrão das demais
  T(s.titulo, {x:0.41, y:0.13, w:8.6, h:0.4, fontSize:15, bold:true});
  if (s.sub) T(s.sub, {x:0.41, y:0.52, w:8.6, h:0.3, fontSize:9, color:C.ink3});
  if (logo) sl.addImage({data:logo, x:9.05, y:0.14, w:0.55, h:0.55});
  T(`${SPEC.mesLabel}   ·   ${i + 1}/${SPEC.slides.length}`,
    {x:0.41, y:5.25, w:9.2, h:0.25, fontSize:8, color:C.ink3});

  const tbl = (rows, opts) => sl.addTable(rows, Object.assign({
    x:0.41, y:0.95, w:9.2, border:{type:'solid', pt:0.4, color:C.line},
    fontFace:'Segoe UI', fontSize:9, color:C.ink, valign:'middle',
  }, opts));
  const th = t => ({text:t, options:{bold:true, fontSize:8, color:C.ink3, fill:{color:C.bg2}, align:'right'}});

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
    tbl(rows, {colW:[3.7, 1.7, 1.7, 1.1, 1.0], rowH:Math.min(0.24, 4.1 / rows.length), fontSize:8});
    return;
  }
  if (s.t === 'kpis_tabela' || s.t === 'matriz'){
    const w = 9.2 / s.kpis.length;
    s.kpis.forEach((k, j) => {
      const x = 0.41 + j * w;
      sl.addShape(p.ShapeType.roundRect, {x, y:0.95, w:w - 0.14, h:0.82,
        fill:{color:C.card}, line:{color:C.line, width:0.5}, rectRadius:0.05});
      T(k.v, {x:x + 0.13, y:1.0, w:w - 0.4, h:0.34, fontSize:17, bold:true, color:C.amber});
      T(k.l, {x:x + 0.13, y:1.34, w:w - 0.4, h:0.22, fontSize:8, color:C.ink2});
      T(k.s, {x:x + 0.13, y:1.53, w:w - 0.4, h:0.2, fontSize:7, color:C.ink3});
    });
    if (s.t === 'kpis_tabela'){
      const rows = [s.tabela.cols.map(th)];
      s.tabela.rows.forEach(r => rows.push(r.map((c, j) =>
        ({text:String(c), options:{align:j === 0 ? 'left' : 'right'}}))));
      tbl(rows, {y:1.95, colW:[5.2, 2.0, 2.0], rowH:0.26});
    } else {
      const rows = [s.cols.map(th)];
      s.rows.forEach(r => rows.push(r.map((c, j) => j === 0
        ? {text:String(c), options:{align:'left', bold:/Saldo/.test(c), color:/Saldo/.test(c) ? C.amber : C.ink2}}
        : {text:c ? rsk(c) : '—', options:{align:'right', color:typeof c === 'number' && c < 0 ? C.neg : C.ink}})));
      const cw = [1.9].concat(new Array(s.cols.length - 1).fill((9.2 - 1.9) / (s.cols.length - 1)));
      tbl(rows, {y:1.95, colW:cw, rowH:0.3});
    }
    return;
  }
  if (s.t === 'lista_mes'){
    let y = 0.95;
    s.meses.forEach(m => {
      T(m.mes, {x:0.41, y, w:0.9, h:0.22, fontSize:9, bold:true, color:C.amber});
      T('ANIMAIS E PRODUTOS', {x:1.35, y, w:3, h:0.22, fontSize:7.5, color:C.ink3});
      T(rs(m.total), {x:7.6, y, w:2.0, h:0.22, fontSize:10, bold:true, align:'right'});
      y += 0.26;
      m.itens.forEach(it => {
        T(it.desc, {x:1.35, y, w:6.1, h:0.2, fontSize:8, color:C.ink2});
        T(rs(it.valor), {x:7.6, y, w:2.0, h:0.2, fontSize:8, align:'right', color:C.ink2});
        y += 0.22;
      });
      y += 0.06;
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
document.getElementById('ir').innerHTML = SPEC.slides
  .map((s, i) => `<option value="${i}">${String(i + 1).padStart(2, '0')} · ${esc((s.titulo || s.t).slice(0, 52))}</option>`).join('');
document.getElementById('prev').onclick = () => go(idx - 1);
document.getElementById('next').onclick = () => go(idx + 1);
document.getElementById('ir').onchange = e => go(+e.target.value);
document.getElementById('play').onclick = play;
document.getElementById('pptx').onclick = e => exportarPptx(e.currentTarget);
const pend = SPEC.slides.filter(s => s.t === 'pendente').length;
document.getElementById('aviso').textContent =
  (pend ? `${pend} de ${total} slides ainda sem fonte` : '') +
  (SPEC.avisos && SPEC.avisos.length ? ` · ${SPEC.avisos[0]}` : '');
go(Math.max(0, (parseInt(location.hash.slice(1), 10) || 1) - 1));
