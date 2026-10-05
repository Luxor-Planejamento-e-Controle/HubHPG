/* SGPG — Sistema de Gestão Pao Grande. Roda em iframe do hub (mesma origem).

   O sistema cresce módulo a módulo; MODULOS lista os que existem. O primeiro é
   "Exposições e premiações": o registro de cada exposição e do que os animais
   ganharam nela, com o histórico consultável. O comitê mensal lê este registro para
   os slides de programação e resultados — não há mais editor de exposições lá.

   Sessão e cliente do Supabase vêm do hub (window.parent.HUB). Quem grava é decidido
   pelo banco (hub_sgpg_editor: admin e sgpg_editores); aqui só se esconde o botão que
   o banco recusaria. */
'use strict';

const STATUS = ['Próxima', 'Aguardando', 'Realizada', 'Cancelada'];
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho',
               'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const MODULOS = [{id: 'exposicoes', titulo: 'Exposições e premiações'}];

const hub = () => { try { return window.parent.HUB || null; } catch (e) { return null; } };
const cliente = () => { const h = hub(); return h && h.sb; };
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g,
  c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const $ = id => document.getElementById(id);

let EXPOS = [], PREM = [], souEditor = false;
let modulo = 'exposicoes', vista = 'lista', ano = new Date().getFullYear();
let aberta = null;            // exposição aberta na ficha: objeto (id null = nova)
let filtro = {animal: '', ano: '', expo: ''};

/* ---- datas ---- */
const dt = iso => { const [a, m, d] = iso.split('-'); return {a: +a, m: +m, d: +d}; };
const dd = n => String(n).padStart(2, '0');
function periodo(e) {
  const i = dt(e.inicio), f = dt(e.fim);
  if (e.so_mes) return `${MESES[i.m - 1]}/${i.a}`;
  if (e.inicio === e.fim) return `${dd(i.d)}/${dd(i.m)}/${i.a}`;
  if (i.a === f.a && i.m === f.m) return `${dd(i.d)} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
  if (i.a === f.a) return `${dd(i.d)}/${dd(i.m)} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
  return `${dd(i.d)}/${dd(i.m)}/${i.a} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
}
function fimDoMes(aaaaMm) {
  const [a, m] = aaaaMm.split('-').map(Number);
  return `${aaaaMm}-${dd(new Date(a, m, 0).getDate())}`;
}

/* ---- dados ---- */
async function checaEditor() {
  const h = hub();
  if (h && h.role === 'admin') { souEditor = true; return; }
  const c = cliente(), email = h && h.email;
  if (!c || !email) { souEditor = false; return; }
  const {data} = await c.from('sgpg_editores').select('email').eq('email', email).maybeSingle();
  souEditor = !!data;
}
async function carrega() {
  const c = cliente();
  const [e, p] = await Promise.all([
    c.from('sgpg_exposicao').select('*').order('inicio'),
    c.from('sgpg_premiacao').select('*').order('exposicao_id').order('ordem'),
  ]);
  if (e.error) throw e.error;
  if (p.error) throw p.error;
  EXPOS = e.data || []; PREM = p.data || [];
}
const premiosDe = id => PREM.filter(p => p.exposicao_id === id);

/* ---- casca ---- */
function render() {
  $('modulos').innerHTML = MODULOS.map(m =>
    `<button type="button" class="${m.id === modulo ? 'on' : ''}" data-mod="${m.id}">${esc(m.titulo)}</button>`).join('');
  const sub = [['lista', 'Exposições'], ['historico', 'Histórico']];
  $('painel').innerHTML = `<div class="submenu">${sub.map(([id, t]) =>
    `<button type="button" class="${id === vista ? 'on' : ''}" data-vista="${id}">${t}</button>`).join('')}</div>
    <div id="corpo"></div>`;
  (vista === 'lista' ? renderLista : renderHistorico)($('corpo'));
}

/* ---- Exposições ---- */
function anos() {
  const s = new Set(EXPOS.map(e => dt(e.inicio).a));
  s.add(new Date().getFullYear());
  return [...s].sort((a, b) => b - a);
}
function renderLista(el) {
  const doAno = EXPOS.filter(e => dt(e.inicio).a === ano);
  const nPrem = doAno.reduce((s, e) => s + premiosDe(e.id).length, 0);
  el.innerHTML = `
    <div class="barra">
      <label class="campo">Ano <select id="selAno">${anos().map(a =>
        `<option ${a === ano ? 'selected' : ''}>${a}</option>`).join('')}</select></label>
      <span class="resumo"><b>${doAno.length}</b> exposição(ões) · <b>${nPrem}</b> premiação(ões)</span>
      <span class="dir">${souEditor ? '<button type="button" class="bt prim" id="btNova">Nova exposição</button>' : ''}</span>
    </div>
    ${doAno.length ? `<table class="t"><thead><tr><th>Período</th><th>Exposição</th><th>Local</th>
      <th>Status</th><th style="text-align:right">Premiações</th></tr></thead><tbody>
      ${doAno.map(e => `<tr class="clica${aberta && aberta.id === e.id ? ' sel' : ''}" data-expo="${e.id}">
        <td class="d">${esc(periodo(e))}</td><td>${esc(e.nome)}</td><td>${esc(e.local || '—')}</td>
        <td><span class="status ${esc(e.status)}">${esc(e.status)}</span></td>
        <td class="n">${premiosDe(e.id).length}</td></tr>`).join('')}
      </tbody></table>` : `<div class="vazio">Nenhuma exposição registrada em ${ano}.</div>`}
    <div id="ficha"></div>`;
  if (aberta) renderFicha($('ficha'));
}

function novaExposicao() {
  const hoje = new Date().toISOString().slice(0, 10);
  aberta = {id: null, nome: '', inicio: hoje, fim: hoje, so_mes: false, local: '', status: 'Próxima',
            observacao: '', premios: []};
}
function abre(id) {
  const e = EXPOS.find(x => x.id === id);
  aberta = {...e, premios: premiosDe(id).map(p => ({id: p.id, animal: p.animal, premio: p.premio}))};
}

function renderFicha(el) {
  const e = aberta;
  if (!souEditor) {
    el.innerHTML = `<div class="ficha"><h2>${esc(e.nome)} <span class="status ${esc(e.status)}">${esc(e.status)}</span></h2>
      <div class="resumo" style="color:var(--ink-2);margin-bottom:10px">${esc(periodo(e))} · ${esc(e.local || '—')}
      ${e.observacao ? ' · ' + esc(e.observacao) : ''}</div>
      ${e.premios.length ? `<table class="t"><thead><tr><th>Animal</th><th>Prêmio</th></tr></thead><tbody>
        ${e.premios.map(p => `<tr><td>${esc(p.animal)}</td><td>${esc(p.premio)}</td></tr>`).join('')}</tbody></table>`
        : '<div class="vazio">Sem premiações registradas.</div>'}</div>`;
    return;
  }
  const animais = [...new Set(PREM.map(p => p.animal))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  el.innerHTML = `<div class="ficha">
    <h2>${e.id ? 'Editar exposição' : 'Nova exposição'}</h2>
    <div class="grade">
      <label class="campo larga">Exposição <input type="text" id="fNome" value="${esc(e.nome)}"
        placeholder="Ex.: 43ª Exposição Nacional ABCCMM"></label>
      ${e.so_mes
        ? `<label class="campo">Mês <input type="month" id="fMes" value="${e.inicio.slice(0, 7)}"></label>`
        : `<label class="campo">Início <input type="date" id="fIni" value="${e.inicio}"></label>
           <label class="campo">Fim <input type="date" id="fFim" value="${e.fim}"></label>`}
      <label class="marca"><input type="checkbox" id="fSoMes" ${e.so_mes ? 'checked' : ''}> só o mês</label>
      <label class="campo">Local <input type="text" id="fLocal" value="${esc(e.local || '')}"></label>
      <label class="campo">Status <select id="fStatus">${STATUS.map(s =>
        `<option ${s === e.status ? 'selected' : ''}>${s}</option>`).join('')}</select></label>
      <label class="campo larga">Observação <textarea id="fObs">${esc(e.observacao || '')}</textarea></label>
    </div>
    <datalist id="animais">${animais.map(a => `<option value="${esc(a)}">`).join('')}</datalist>
    <table class="t premios"><thead><tr><th style="width:38%">Animal</th><th>Prêmio</th><th></th></tr></thead><tbody>
      ${e.premios.map((p, i) => `<tr>
        <td><input type="text" list="animais" data-i="${i}" data-c="animal" value="${esc(p.animal)}"></td>
        <td><input type="text" data-i="${i}" data-c="premio" value="${esc(p.premio)}"></td>
        <td class="n"><button type="button" class="x" data-tira="${i}" title="Remover">×</button></td></tr>`).join('')}
    </tbody></table>
    <div class="acoes">
      <button type="button" class="bt" id="btAddPremio">Adicionar premiação</button>
      <span class="msg" id="msg"></span>
      <span class="dir">
        ${e.id ? '<button type="button" class="bt perigo" id="btExcluir">Excluir exposição</button>' : ''}
        <button type="button" class="bt" id="btFechar">Cancelar</button>
        <button type="button" class="bt prim" id="btSalvar">Salvar</button>
      </span>
    </div></div>`;
}

/* lê o formulário de volta para `aberta`, antes de redesenhar ou salvar */
function leFicha() {
  const e = aberta, v = id => ($(id) || {}).value;
  e.nome = (v('fNome') || '').trim();
  e.local = (v('fLocal') || '').trim();
  e.status = v('fStatus') || e.status;
  e.observacao = (v('fObs') || '').trim();
  if ($('fMes')) {
    const m = v('fMes');
    if (m) { e.inicio = `${m}-01`; e.fim = fimDoMes(m); }
  } else {
    e.inicio = v('fIni') || e.inicio; e.fim = v('fFim') || e.fim;
  }
  document.querySelectorAll('.premios input[data-i]').forEach(inp => {
    e.premios[+inp.dataset.i][inp.dataset.c] = inp.value;
  });
}

function dizer(txt, erro) {
  const m = $('msg');
  if (m) { m.textContent = txt || ''; m.classList.toggle('erro', !!erro); }
}

async function salvar() {
  leFicha();
  const e = aberta, c = cliente();
  const premios = e.premios.map(p => ({...p, animal: p.animal.trim(), premio: p.premio.trim()}))
                           .filter(p => p.animal || p.premio);
  if (!e.nome) return dizer('Informe o nome da exposição.', true);
  if (!e.inicio || !e.fim || e.fim < e.inicio) return dizer('Confira as datas: o fim não pode vir antes do início.', true);
  if (premios.some(p => !p.animal || !p.premio)) return dizer('Cada premiação precisa de animal e prêmio.', true);
  const btn = $('btSalvar'); btn.disabled = true; dizer('salvando…');
  try {
    const linha = {nome: e.nome, inicio: e.inicio, fim: e.fim, so_mes: e.so_mes, local: e.local || null,
                   status: e.status, observacao: e.observacao || null};
    let id = e.id;
    if (id) {
      const r = await c.from('sgpg_exposicao').update(linha).eq('id', id);
      if (r.error) throw r.error;
    } else {
      const r = await c.from('sgpg_exposicao').insert(linha).select('id').single();
      if (r.error) throw r.error;
      id = r.data.id;
    }
    // premiações por diferença: nada é apagado em bloco antes de regravar
    const antes = premiosDe(id), fica = new Set(premios.filter(p => p.id).map(p => p.id));
    const tirar = antes.filter(p => !fica.has(p.id)).map(p => p.id);
    if (tirar.length) {
      const r = await c.from('sgpg_premiacao').delete().in('id', tirar);
      if (r.error) throw r.error;
    }
    for (const [ordem, p] of premios.entries()) {
      const r = p.id
        ? await c.from('sgpg_premiacao').update({animal: p.animal, premio: p.premio, ordem}).eq('id', p.id)
        : await c.from('sgpg_premiacao').insert({exposicao_id: id, animal: p.animal, premio: p.premio, ordem});
      if (r.error) throw r.error;
    }
    await carrega();
    ano = dt(e.inicio).a;
    abre(id);
    render();
    dizer('salvo.');
  } catch (err) {
    btn.disabled = false;
    dizer('Não deu pra salvar: ' + (err.message || err), true);
  }
}

async function excluir() {
  const e = aberta;
  if (!confirm(`Excluir "${e.nome}" e as ${e.premios.length} premiação(ões) dela? Não há como desfazer.`)) return;
  const r = await cliente().from('sgpg_exposicao').delete().eq('id', e.id);
  if (r.error) return dizer('Não deu pra excluir: ' + r.error.message, true);
  aberta = null;
  await carrega();
  render();
}

/* ---- Histórico ---- */
function renderHistorico(el) {
  const porId = Object.fromEntries(EXPOS.map(e => [e.id, e]));
  const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const linhas = PREM.map(p => ({...p, e: porId[p.exposicao_id]})).filter(p => p.e)
    .filter(p => !filtro.ano || String(dt(p.e.inicio).a) === filtro.ano)
    .filter(p => !filtro.expo || String(p.exposicao_id) === filtro.expo)
    .filter(p => !filtro.animal || norm(p.animal).includes(norm(filtro.animal)))
    .sort((a, b) => b.e.inicio.localeCompare(a.e.inicio) || a.ordem - b.ordem);
  const ranking = {};
  linhas.forEach(p => { ranking[p.animal] = (ranking[p.animal] || 0) + 1; });
  const top = Object.entries(ranking).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
  const animais = [...new Set(PREM.map(p => p.animal))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const expos = EXPOS.filter(e => !filtro.ano || String(dt(e.inicio).a) === filtro.ano);
  el.innerHTML = `
    <div class="barra">
      <label class="campo">Animal <input type="search" id="hAnimal" list="hAnimais" value="${esc(filtro.animal)}"
        placeholder="nome ou parte do nome"></label>
      <datalist id="hAnimais">${animais.map(a => `<option value="${esc(a)}">`).join('')}</datalist>
      <label class="campo">Ano <select id="hAno"><option value="">todos</option>${anos().map(a =>
        `<option ${String(a) === filtro.ano ? 'selected' : ''}>${a}</option>`).join('')}</select></label>
      <label class="campo">Exposição <select id="hExpo"><option value="">todas</option>${expos.map(e =>
        `<option value="${e.id}" ${String(e.id) === filtro.expo ? 'selected' : ''}>${esc(e.nome)}</option>`).join('')}</select></label>
      <span class="resumo"><b>${linhas.length}</b> premiação(ões) · <b>${top.length}</b> animal(is) ·
        <b>${new Set(linhas.map(p => p.exposicao_id)).size}</b> exposição(ões)</span>
    </div>
    <div class="lado">
      <div>${linhas.length ? `<table class="t"><thead><tr><th>Período</th><th>Exposição</th><th>Animal</th><th>Prêmio</th></tr></thead><tbody>
        ${linhas.map(p => `<tr><td class="d">${esc(periodo(p.e))}</td><td>${esc(p.e.nome)}</td>
          <td>${esc(p.animal)}</td><td>${esc(p.premio)}</td></tr>`).join('')}</tbody></table>`
        : '<div class="vazio">Nenhuma premiação com esses filtros.</div>'}</div>
      <div><h3>Premiações por animal</h3>${top.length ? `<table class="t"><tbody>
        ${top.map(([a, n]) => `<tr class="clica" data-animal="${esc(a)}"><td>${esc(a)}</td><td class="n">${n}</td></tr>`).join('')}
        </tbody></table>` : '<div class="vazio">—</div>'}</div>
    </div>`;
}

/* ---- eventos (delegados: o painel é redesenhado inteiro) ---- */
document.addEventListener('click', ev => {
  const t = ev.target.closest('button, tr.clica');
  if (!t) return;
  if (t.dataset.mod) { modulo = t.dataset.mod; render(); return; }
  if (t.dataset.vista) { vista = t.dataset.vista; render(); return; }
  if (t.dataset.expo) { abre(+t.dataset.expo); render(); $('ficha').scrollIntoView({block: 'nearest'}); return; }
  if (t.dataset.animal) { filtro = {animal: t.dataset.animal, ano: '', expo: ''}; render(); return; }
  if (t.dataset.tira) { leFicha(); aberta.premios.splice(+t.dataset.tira, 1); renderFicha($('ficha')); return; }
  switch (t.id) {
    case 'btNova': novaExposicao(); render(); $('ficha').scrollIntoView({block: 'nearest'}); break;
    case 'btAddPremio':
      leFicha(); aberta.premios.push({id: null, animal: '', premio: ''}); renderFicha($('ficha'));
      { const ins = document.querySelectorAll('.premios input[data-c="animal"]'); ins[ins.length - 1].focus(); }
      break;
    case 'btFechar': aberta = null; render(); break;
    case 'btSalvar': salvar(); break;
    case 'btExcluir': excluir(); break;
  }
});
document.addEventListener('change', ev => {
  const t = ev.target;
  if (t.id === 'selAno') { ano = +t.value; aberta = null; render(); }
  else if (t.id === 'fSoMes') {
    leFicha(); aberta.so_mes = t.checked;
    if (t.checked) { const m = aberta.inicio.slice(0, 7); aberta.inicio = `${m}-01`; aberta.fim = fimDoMes(m); }
    renderFicha($('ficha'));
  }
  else if (t.id === 'hAno') { filtro.ano = t.value; filtro.expo = ''; render(); }
  else if (t.id === 'hExpo') { filtro.expo = t.value; render(); }
});
document.addEventListener('input', ev => {
  if (ev.target.id !== 'hAnimal') return;
  filtro.animal = ev.target.value;
  const pos = ev.target.selectionStart;
  render();
  const i = $('hAnimal'); i.focus(); i.setSelectionRange(pos, pos);
});

/* ---- início ---- */
(async function () {
  if (!cliente()) {
    $('painel').innerHTML = '<div class="vazio">Abra o SGPG pelo hub: a sessão vem de lá.</div>';
    return;
  }
  try {
    await Promise.all([checaEditor(), carrega()]);
    render();
  } catch (err) {
    $('painel').innerHTML = `<div class="vazio">Não deu pra carregar: ${esc(err.message || err)}</div>`;
  }
})();
