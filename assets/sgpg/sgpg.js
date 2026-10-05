/* SGPG — Sistema de Gestão Pao Grande. Roda em iframe do hub (mesma origem).

   O sistema cresce módulo a módulo (MODULOS). O primeiro é Exposições e premiações,
   em duas abas:
     Registro   cadastro da exposição, com o que cada animal ganhou nela
     Histórico  consulta por exposição ou por animal; é dali que se corrige um registro
   O comitê mensal lê este registro (build_comite.py e deck.js, a partir de ago/2026).

   Sessão e cliente do Supabase vêm do hub (window.parent.HUB). Quem grava é decidido
   pelo banco (hub_sgpg_editor: admin e sgpg_editores); aqui só se escondem os botões
   que o banco recusaria. */
'use strict';

const STATUS = ['Próxima', 'Aguardando', 'Realizada', 'Cancelada'];
const CLS_STATUS = {'Próxima': 'st-prox', 'Aguardando': 'st-aguard', 'Realizada': 'st-real', 'Cancelada': 'st-canc'};
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho',
               'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const ABR = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
const MODULOS = [{id: 'exposicoes', titulo: 'Exposições e premiações'}];

const hub = () => { try { return window.parent.HUB || null; } catch (e) { return null; } };
const cliente = () => { const h = hub(); return h && h.sb; };
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g,
  c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const $ = id => document.getElementById(id);
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const plural = (n, um, varios) => `${n} ${n === 1 ? um : varios}`;
const lembra = (k, v) => { try { localStorage.setItem('sgpg.' + k, v); } catch (e) { /* sem storage */ } };
const lembrado = k => { try { return localStorage.getItem('sgpg.' + k); } catch (e) { return null; } };

let EXPOS = [], PREM = [], souEditor = false;
let aba = null;                                   // 'registro' | 'historico'
let visao = lembrado('visao') === 'animal' ? 'animal' : 'expo';
let filtro = {ano: String(new Date().getFullYear()), busca: ''};
let selExpo = null, selAnimal = null;             // selAnimal é a chave do animal
let rascunho = null;                              // formulário do Registro
let edicao = null, edicaoOriginal = '';           // registro sendo corrigido no Histórico
let registrada = null, salvoAgora = null;         // avisos depois de gravar

/* ---- datas ---- */
const dt = iso => { const [a, m, d] = iso.split('-').map(Number); return {a, m, d}; };
const dd = n => String(n).padStart(2, '0');
const hojeIso = () => { const h = new Date(); return `${h.getFullYear()}-${dd(h.getMonth() + 1)}-${dd(h.getDate())}`; };
function fimDoMes(aaaaMm) {
  const [a, m] = aaaaMm.split('-').map(Number);
  return `${aaaaMm}-${dd(new Date(a, m, 0).getDate())}`;
}
/* 'Junho/2026', '24/09/2026', '06 a 12/04/2026', '25/04 a 02/05/2026' — a mesma regra
   do periodo_sgpg do build_comite.py, que é como o comitê escreve */
function periodo(e) {
  const i = dt(e.inicio), f = dt(e.fim);
  if (e.so_mes) return `${MESES[i.m - 1]}/${i.a}`;
  if (e.inicio === e.fim) return `${dd(i.d)}/${dd(i.m)}/${i.a}`;
  if (i.a === f.a && i.m === f.m) return `${dd(i.d)} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
  if (i.a === f.a) return `${dd(i.d)}/${dd(i.m)} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
  return `${dd(i.d)}/${dd(i.m)}/${i.a} a ${dd(f.d)}/${dd(f.m)}/${f.a}`;
}
function seloHtml(e) {
  const i = dt(e.inicio), f = dt(e.fim);
  let b, s;
  if (e.so_mes) { b = ABR[i.m - 1]; s = 'MÊS'; }
  else if (e.inicio === e.fim) { b = dd(i.d); s = ABR[i.m - 1]; }
  else if (i.m === f.m && i.a === f.a) { b = `${dd(i.d)}–${dd(f.d)}`; s = ABR[i.m - 1]; }
  else { b = `${dd(i.d)}–${dd(f.d)}`; s = `${ABR[i.m - 1]}–${ABR[f.m - 1]}`; }
  return `<span class="selo"><b>${b}</b><span>${s}</span></span>`;
}
const statusHtml = s => `<span class="status ${CLS_STATUS[s] || ''}">${esc(s)}</span>`;

/* ---- animais ----
   O nome é como o haras escreveu, e o mesmo animal aparece de mais de um jeito ("da PG"
   e "da Pao Grande", com e sem acento, com e sem "do"). Na consulta, essas grafias são
   um animal só; o nome mostrado é a grafia mais usada. */
const chaveAnimal = s => norm(s).replace(/\bpao grande\b/g, 'pg')
  .replace(/\b(do|da|de|dos|das)\b/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
function animais(exps) {
  const porId = new Map(exps.map(e => [e.id, e])), mapa = new Map();
  for (const p of PREM) {
    const e = porId.get(p.exposicao_id);
    if (!e) continue;
    const k = chaveAnimal(p.animal);
    if (!mapa.has(k)) mapa.set(k, {chave: k, grafias: new Map(), ps: []});
    const a = mapa.get(k);
    const g = a.grafias.get(p.animal) || {n: 0, ultima: ''};
    g.n += 1; if (e.inicio > g.ultima) g.ultima = e.inicio;
    a.grafias.set(p.animal, g);
    a.ps.push({...p, e});
  }
  return [...mapa.values()].map(a => {
    const nomes = [...a.grafias.entries()].sort((x, y) => y[1].n - x[1].n || y[1].ultima.localeCompare(x[1].ultima));
    return {...a, nome: nomes[0][0], outras: nomes.slice(1).map(x => x[0]),
            nExp: new Set(a.ps.map(p => p.exposicao_id)).size};
  });
}
const nomesConhecidos = () => animais(EXPOS).map(a => a.nome).sort((a, b) => a.localeCompare(b, 'pt-BR'));

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
function agrupaPorAnimal(ps) {
  const out = [], idx = new Map();
  for (const p of ps) {
    if (!idx.has(p.animal)) { idx.set(p.animal, out.length); out.push({animal: p.animal, premios: []}); }
    out[idx.get(p.animal)].premios.push(p);
  }
  return out;
}

/* ---- formulário (o mesmo no Registro e na correção pelo Histórico) ---- */
const premioVazio = () => ({id: null, premio: ''});
const animalVazio = () => ({animal: '', premios: [premioVazio()]});
function novoForm() {
  const h = hojeIso();
  return {id: null, nome: '', inicio: h, fim: h, so_mes: false, local: '', status: 'Próxima',
          observacao: '', animais: [animalVazio()]};
}
function formDe(e) {
  const g = agrupaPorAnimal(premiosDe(e.id));
  return {id: e.id, nome: e.nome, inicio: e.inicio, fim: e.fim, so_mes: e.so_mes, local: e.local || '',
          status: e.status, observacao: e.observacao || '',
          animais: g.length ? g.map(x => ({animal: x.animal, premios: x.premios.map(p => ({id: p.id, premio: p.premio}))}))
                            : [animalVazio()]};
}
const formAtivo = qual => qual === 'edicao' ? edicao : rascunho;
function preenchido(f) {
  return !!(f.nome.trim() || f.local.trim() || f.observacao.trim()
    || f.animais.some(a => a.animal.trim() || a.premios.some(p => p.premio.trim())));
}
const edicaoMudou = () => !!edicao && JSON.stringify(edicao) !== edicaoOriginal;
const podeLargarEdicao = () => !edicaoMudou() || confirm('Descartar as alterações deste registro?');

function contagem(f) {
  const nPrem = f.animais.reduce((s, a) => s + (a.animal.trim() ? a.premios.filter(p => p.premio.trim()).length : 0), 0);
  const nAnim = f.animais.filter(a => a.animal.trim()).length;
  return nAnim ? `${plural(nPrem, 'prêmio', 'prêmios')} · ${plural(nAnim, 'animal', 'animais')}` : '';
}

const animalHtml = (a, ai, q) => `<div class="animal">
    <div class="animal-cab">
      <input type="text" id="${q}-a${ai}" list="lista-animais" data-ai="${ai}" data-c="animal"
        value="${esc(a.animal)}" placeholder="Nome do animal" autocomplete="off" aria-label="Animal">
      <button type="button" class="x" data-acao="rm-animal" data-ai="${ai}" title="Remover animal" aria-label="Remover animal">×</button>
    </div>
    ${a.premios.map((p, pi) => `<div class="premio">
      <input type="text" id="${q}-a${ai}-p${pi}" data-ai="${ai}" data-pi="${pi}" value="${esc(p.premio)}"
        placeholder="Prêmio, ex.: Campeã de Marcha" autocomplete="off" aria-label="Prêmio">
      <button type="button" class="x" data-acao="rm-premio" data-ai="${ai}" data-pi="${pi}" title="Remover prêmio" aria-label="Remover prêmio">×</button>
    </div>`).join('')}
    <button type="button" class="bt link" data-acao="add-premio" data-ai="${ai}">+ Prêmio</button>
  </div>`;

function formHtml(f, q) {
  const datas = f.so_mes
    ? `<input type="month" id="${q}-mes" data-f="mes" value="${f.inicio.slice(0, 7)}" aria-label="Mês">`
    : `<input type="date" id="${q}-ini" data-f="inicio" value="${f.inicio}" aria-label="Início">
       <span class="ate">a</span>
       <input type="date" id="${q}-fim" data-f="fim" value="${f.fim}" aria-label="Fim">`;
  const botoes = q === 'registro'
    ? `<button type="button" class="bt" data-acao="limpar">Limpar</button>
       <button type="button" class="bt prim" data-acao="salvar">Registrar exposição</button>`
    : `<button type="button" class="bt" data-acao="cancelar-edicao">Cancelar</button>
       <button type="button" class="bt prim" data-acao="salvar">Salvar alterações</button>`;
  return `<div class="registro" data-form="${q}">
    <section class="cartao">
      <div class="cartao-cab"><h2>Exposição</h2></div>
      <div class="campos">
        <label class="campo"><span class="rot">Nome</span>
          <input type="text" id="${q}-nome" data-f="nome" value="${esc(f.nome)}"
            placeholder="Ex.: 43ª Exposição Nacional ABCCMM" autocomplete="off"></label>
        <div class="campo"><span class="rot">Período</span>
          <div class="periodo">
            <div class="seg" role="group" aria-label="Tipo de período">
              <button type="button" class="${f.so_mes ? '' : 'on'}" data-acao="periodo" data-v="datas">Datas</button>
              <button type="button" class="${f.so_mes ? 'on' : ''}" data-acao="periodo" data-v="mes">Só o mês</button>
            </div>
            ${datas}
          </div></div>
        <label class="campo"><span class="rot">Local</span>
          <input type="text" id="${q}-local" data-f="local" list="lista-locais" value="${esc(f.local)}"
            placeholder="Cidade ou parque" autocomplete="off"></label>
        <div class="campo"><span class="rot">Status</span>
          <div class="seg status" role="group" aria-label="Status">${STATUS.map(s =>
            `<button type="button" class="${s === f.status ? 'on ' : ''}${CLS_STATUS[s]}" data-acao="status" data-v="${s}">${s}</button>`).join('')}</div></div>
        <label class="campo"><span class="rot">Observação</span>
          <textarea id="${q}-obs" data-f="observacao" rows="2">${esc(f.observacao)}</textarea></label>
      </div>
    </section>
    <section class="cartao">
      <div class="cartao-cab"><h2>Premiações</h2><span class="cont" data-cont>${contagem(f)}</span></div>
      <div class="animais">${f.animais.map((a, ai) => animalHtml(a, ai, q)).join('')}</div>
      <button type="button" class="bt" data-acao="add-animal">+ Animal</button>
    </section>
    <div class="acoes">
      ${q === 'edicao' ? '<button type="button" class="bt perigo" data-acao="excluir">Excluir registro</button>' : ''}
      <span class="msg" data-msg aria-live="polite"></span>
      <span class="dir">${botoes}</span>
    </div>
  </div>`;
}

const foca = id => {
  const x = document.getElementById(id);
  if (!x) return;
  x.focus();
  if (x.type === 'text') x.setSelectionRange(x.value.length, x.value.length);
};
function redesenhaForm(q) {
  const box = document.querySelector(`[data-form="${q}"]`);
  if (!box) return;
  const foco = document.activeElement && document.activeElement.id;
  box.outerHTML = formHtml(formAtivo(q), q);
  if (foco) foca(foco);
}

function escreve(f, t, box) {
  if (t.dataset.f === 'mes') {
    if (t.value) { f.inicio = `${t.value}-01`; f.fim = fimDoMes(t.value); }
  } else if (t.dataset.f) {
    f[t.dataset.f] = t.value;
    // fim antes do início não existe: acompanha o início
    if (t.dataset.f === 'inicio' && f.fim < f.inicio) {
      f.fim = f.inicio;
      const fim = box.querySelector('[data-f="fim"]');
      if (fim) fim.value = f.fim;
    }
  } else if (t.dataset.c === 'animal') {
    f.animais[+t.dataset.ai].animal = t.value;
  } else if (t.dataset.pi != null) {
    f.animais[+t.dataset.ai].premios[+t.dataset.pi].premio = t.value;
  }
}

function linhasDe(f) {
  const out = [];
  for (const a of f.animais) {
    const nome = a.animal.trim();
    const ps = a.premios.map(p => ({...p, premio: p.premio.trim()})).filter(p => p.premio);
    if (!nome && !ps.length) continue;
    if (!nome) return {erro: 'Tem prêmio sem animal. Informe o animal ou remova o prêmio.'};
    if (!ps.length) return {erro: `Informe ao menos um prêmio de ${nome}, ou remova o animal.`};
    ps.forEach(p => out.push({id: p.id, animal: nome, premio: p.premio}));
  }
  return {linhas: out};
}

async function salvar(q) {
  const f = formAtivo(q), box = document.querySelector(`[data-form="${q}"]`), c = cliente();
  const dizer = (txt, erro) => { const m = box.querySelector('[data-msg]'); m.textContent = txt; m.classList.toggle('erro', !!erro); };
  const nome = f.nome.trim();
  if (!nome) { foca(`${q}-nome`); return dizer('Informe o nome da exposição.', true); }
  if (!f.inicio || !f.fim) return dizer('Informe as datas da exposição.', true);
  if (f.fim < f.inicio) return dizer('O fim não pode ser antes do início.', true);
  const {linhas, erro} = linhasDe(f);
  if (erro) return dizer(erro, true);
  const btn = box.querySelector('[data-acao="salvar"]');
  btn.disabled = true; dizer('Salvando…');
  try {
    const linha = {nome, inicio: f.inicio, fim: f.fim, so_mes: f.so_mes, local: f.local.trim() || null,
                   status: f.status, observacao: f.observacao.trim() || null};
    let id = f.id;
    if (id) {
      const r = await c.from('sgpg_exposicao').update(linha).eq('id', id);
      if (r.error) throw r.error;
    } else {
      const r = await c.from('sgpg_exposicao').insert(linha).select('id').single();
      if (r.error) throw r.error;
      id = r.data.id;
    }
    // prêmios por diferença: nada é apagado em bloco antes de regravar
    const fica = new Set(linhas.filter(p => p.id).map(p => p.id));
    const tirar = premiosDe(id).filter(p => !fica.has(p.id)).map(p => p.id);
    if (tirar.length) {
      const r = await c.from('sgpg_premiacao').delete().in('id', tirar);
      if (r.error) throw r.error;
    }
    for (const [ordem, p] of linhas.entries()) {
      const r = p.id
        ? await c.from('sgpg_premiacao').update({animal: p.animal, premio: p.premio, ordem}).eq('id', p.id)
        : await c.from('sgpg_premiacao').insert({exposicao_id: id, animal: p.animal, premio: p.premio, ordem});
      if (r.error) throw r.error;
    }
    await carrega();
    if (q === 'registro') {
      registrada = {id, nome};
      rascunho = novoForm();
    } else {
      edicao = null; selExpo = id; salvoAgora = id;
    }
    render();
  } catch (err) {
    btn.disabled = false;
    dizer('Não foi possível salvar: ' + (err.message || err), true);
  }
}

async function excluir() {
  const e = EXPOS.find(x => x.id === edicao.id);
  const n = premiosDe(e.id).length;
  if (!confirm(`Excluir "${e.nome}"${n ? ` e ${plural(n, 'premiação', 'premiações')}` : ''}? Não dá para desfazer.`)) return;
  const r = await cliente().from('sgpg_exposicao').delete().eq('id', e.id);
  if (r.error) {
    const m = document.querySelector('[data-form="edicao"] [data-msg]');
    m.textContent = 'Não foi possível excluir: ' + r.error.message; m.classList.add('erro');
    return;
  }
  edicao = null; selExpo = null;
  await carrega();
  render();
}

/* ---- casca ---- */
function render() {
  if (!souEditor && aba === 'registro') aba = 'historico';
  const abas = (souEditor ? [['registro', 'Registro']] : []).concat([['historico', 'Histórico']]);
  $('topo').innerHTML = `<h1>${esc(MODULOS[0].titulo)}</h1>`;
  $('abas').innerHTML = abas.map(([id, t]) =>
    `<button type="button" class="${id === aba ? 'on' : ''}" data-acao="aba" data-v="${id}">${t}</button>`).join('');
  const locais = [...new Set(EXPOS.map(e => e.local).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  $('listas').innerHTML = `<datalist id="lista-animais">${nomesConhecidos().map(a => `<option value="${esc(a)}">`).join('')}</datalist>
    <datalist id="lista-locais">${locais.map(l => `<option value="${esc(l)}">`).join('')}</datalist>`;
  (aba === 'registro' ? renderRegistro : renderHistorico)($('painel'));
}

/* ---- Registro ---- */
function renderRegistro(el) {
  el.innerHTML = `${registrada ? `<div class="faixa-ok" role="status">
      <span><b>${esc(registrada.nome)}</b> registrada.</span>
      <button type="button" class="bt link" data-acao="ver-registro" data-id="${registrada.id}">Ver no histórico</button>
    </div>` : ''}${formHtml(rascunho, 'registro')}`;
}

/* ---- Histórico ---- */
const anos = () => [...new Set(EXPOS.map(e => dt(e.inicio).a).concat(new Date().getFullYear()))].sort((a, b) => b - a);
const escopo = () => EXPOS.filter(e => !filtro.ano || String(dt(e.inicio).a) === filtro.ano);

function renderHistorico(el) {
  const exps = escopo(), ids = new Set(exps.map(e => e.id));
  const prems = PREM.filter(p => ids.has(p.exposicao_id));
  const kpis = [['Exposições', exps.length], ['Realizadas', exps.filter(e => e.status === 'Realizada').length],
                ['Premiações', prems.length], ['Animais premiados', animais(exps).length]];
  el.innerHTML = `
    <div class="hist-barra">
      <div class="seg" role="group" aria-label="Ver por">
        <button type="button" class="${visao === 'expo' ? 'on' : ''}" data-acao="visao" data-v="expo">Por exposição</button>
        <button type="button" class="${visao === 'animal' ? 'on' : ''}" data-acao="visao" data-v="animal">Por animal</button>
      </div>
      <div class="anos" role="group" aria-label="Ano">
        <button type="button" class="${filtro.ano ? '' : 'on'}" data-acao="ano" data-v="">Todos</button>
        ${anos().map(a => `<button type="button" class="${String(a) === filtro.ano ? 'on' : ''}" data-acao="ano" data-v="${a}">${a}</button>`).join('')}
      </div>
      <div class="busca">
        <input type="search" id="busca" value="${esc(filtro.busca)}" autocomplete="off"
          placeholder="${visao === 'expo' ? 'Buscar exposição, local, animal ou prêmio' : 'Buscar animal ou prêmio'}" aria-label="Buscar">
        <button type="button" class="x" data-acao="limpa-busca" title="Limpar busca" aria-label="Limpar busca" ${filtro.busca ? '' : 'hidden'}>×</button>
      </div>
    </div>
    <div class="kpis">${kpis.map(([r, v]) => `<div class="kpi"><span class="rot">${r}</span><b>${v}</b></div>`).join('')}</div>
    <div id="hist-corpo"></div>`;
  renderCorpo();
}
function renderCorpo() {
  const el = $('hist-corpo');
  if (el) (visao === 'expo' ? corpoExpo : corpoAnimal)(el);
}

function corpoExpo(el) {
  const q = norm(filtro.busca);
  const exps = escopo().filter(e => !q || [e.nome, e.local, ...premiosDe(e.id).flatMap(p => [p.animal, p.premio])]
      .some(s => norm(s).includes(q)))
    .sort((a, b) => b.inicio.localeCompare(a.inicio));
  if (!exps.length) {
    el.innerHTML = `<div class="vazio">${q ? 'Nada encontrado com essa busca.'
      : `Nenhuma exposição registrada${filtro.ano ? ' em ' + filtro.ano : ''}.`}</div>`;
    return;
  }
  if (!exps.some(e => e.id === selExpo)) { selExpo = exps[0].id; edicao = null; }
  let ano = null, lista = '';
  for (const e of exps) {
    const a = dt(e.inicio).a, n = premiosDe(e.id).length;
    if (a !== ano) { lista += `<div class="lista-ano">${a}</div>`; ano = a; }
    lista += `<button type="button" class="it${e.id === selExpo ? ' sel' : ''}" data-acao="sel-expo" data-id="${e.id}">
      ${seloHtml(e)}
      <span class="it-txt"><span class="it-nome">${esc(e.nome)}</span>
        <span class="it-meta"><span class="ponto ${CLS_STATUS[e.status]}"></span>${esc(e.status)}${e.local ? ' · ' + esc(e.local) : ''}</span></span>
      <span class="it-n" title="${plural(n, 'premiação', 'premiações')}">${n || ''}</span>
    </button>`;
  }
  const e = EXPOS.find(x => x.id === selExpo);
  const det = edicao && edicao.id === e.id
    ? `<div class="edicao-cab"><h2>Corrigir registro</h2><span>${esc(e.nome)}</span></div>${formHtml(edicao, 'edicao')}`
    : detalheExpo(e);
  el.innerHTML = `<div class="md"><nav class="lista" aria-label="Exposições">${lista}</nav>
    <div class="painel-det">${det}</div></div>`;
}

function detalheExpo(e) {
  const grupos = agrupaPorAnimal(premiosDe(e.id));
  const n = premiosDe(e.id).length;
  const semPremio = {Realizada: 'Nenhuma premiação registrada.', Cancelada: 'Exposição cancelada.'}[e.status]
    || 'A exposição ainda não aconteceu.';
  const alt = e.alterado_em ? new Date(e.alterado_em).toLocaleDateString('pt-BR') : '';
  return `<article class="det">
    <header class="det-cab">
      <div><div class="det-kicker">${esc(periodo(e))}${e.local ? ' · ' + esc(e.local) : ''}</div>
        <h2>${esc(e.nome)}</h2></div>
      ${statusHtml(e.status)}
    </header>
    ${e.observacao ? `<p class="det-obs">${esc(e.observacao)}</p>` : ''}
    <section class="det-sec">
      <h3>Premiações ${n ? `<span class="cont">${n}</span>` : ''}</h3>
      ${grupos.length ? `<div class="grupos">${grupos.map(g => `<div class="grupo">
          <div class="grupo-cab">
            <button type="button" class="lnk" data-acao="ir-animal" data-animal="${esc(chaveAnimal(g.animal))}">${esc(g.animal)}</button>
            <span class="cont">${g.premios.length}</span>
          </div>
          <ul class="premios">${g.premios.map(p => `<li>${esc(p.premio)}</li>`).join('')}</ul>
        </div>`).join('')}</div>` : `<div class="vazio">${semPremio}</div>`}
    </section>
    <footer class="det-pe">
      ${alt ? `<span class="quando">Última alteração em ${alt}</span>` : ''}
      ${salvoAgora === e.id ? '<span class="msg ok">Alterações salvas.</span>' : ''}
      ${souEditor ? `<span class="dir"><button type="button" class="bt" data-acao="editar" data-id="${e.id}">Editar registro</button></span>` : ''}
    </footer>
  </article>`;
}

function corpoAnimal(el) {
  const q = norm(filtro.busca);
  const lista = animais(escopo())
    .filter(a => !q || norm(a.nome).includes(q) || a.outras.some(o => norm(o).includes(q))
      || a.ps.some(p => norm(p.premio).includes(q)))
    .sort((a, b) => b.ps.length - a.ps.length || a.nome.localeCompare(b.nome, 'pt-BR'));
  if (!lista.length) {
    el.innerHTML = `<div class="vazio">${q ? 'Nenhum animal encontrado com essa busca.'
      : `Nenhum animal premiado${filtro.ano ? ' em ' + filtro.ano : ''}.`}</div>`;
    return;
  }
  if (!lista.some(a => a.chave === selAnimal)) selAnimal = lista[0].chave;
  const max = lista[0].ps.length;
  const itens = lista.map(a => `<button type="button" class="it it-animal${a.chave === selAnimal ? ' sel' : ''}" data-acao="sel-animal" data-animal="${esc(a.chave)}">
      <span class="inicial" aria-hidden="true">${esc(a.nome.charAt(0).toUpperCase())}</span>
      <span class="it-txt"><span class="it-nome">${esc(a.nome)}</span>
        <span class="barra" aria-hidden="true"><i style="width:${Math.max(4, Math.round(100 * a.ps.length / max))}%"></i></span>
        <span class="it-meta">${plural(a.nExp, 'exposição', 'exposições')}</span></span>
      <span class="it-n" title="${plural(a.ps.length, 'premiação', 'premiações')}">${a.ps.length}</span>
    </button>`).join('');
  el.innerHTML = `<div class="md"><nav class="lista" aria-label="Animais">${itens}</nav>
    <div class="painel-det">${detalheAnimal(lista.find(a => a.chave === selAnimal))}</div></div>`;
}

function detalheAnimal(a) {
  const porExpo = new Map();
  for (const p of a.ps) {
    if (!porExpo.has(p.exposicao_id)) porExpo.set(p.exposicao_id, {e: p.e, premios: []});
    porExpo.get(p.exposicao_id).premios.push(p);
  }
  const itens = [...porExpo.values()].sort((x, y) => y.e.inicio.localeCompare(x.e.inicio));
  const ult = dt(itens[0].e.inicio);
  return `<article class="det">
    <header class="det-cab"><div>
      <div class="det-kicker">${filtro.ano ? 'Premiações em ' + filtro.ano : 'Todas as premiações'}</div>
      <h2>${esc(a.nome)}</h2>
      ${a.outras.length ? `<div class="grafias">Também registrado como ${a.outras.map(o => `<i>${esc(o)}</i>`).join(', ')}</div>` : ''}
    </div></header>
    <div class="mini-kpis">
      <div><b>${a.ps.length}</b><span>Premiações</span></div>
      <div><b>${a.nExp}</b><span>Exposições</span></div>
      <div><b>${MESES[ult.m - 1].slice(0, 3).toLowerCase()}/${ult.a}</b><span>Última premiação</span></div>
    </div>
    <section class="det-sec"><h3>Por exposição</h3>
      <ol class="tempo">${itens.map(({e, premios}) => `<li class="tempo-it">
        ${seloHtml(e)}
        <div>
          <button type="button" class="lnk" data-acao="ir-expo" data-id="${e.id}">${esc(e.nome)}</button>
          <div class="it-meta">${esc(periodo(e))}${e.local ? ' · ' + esc(e.local) : ''}</div>
          <ul class="premios">${premios.map(p => `<li>${esc(p.premio)}</li>`).join('')}</ul>
        </div></li>`).join('')}</ol>
    </section>
  </article>`;
}

const rolaParaDetalhe = () => {
  if (window.matchMedia('(max-width: 900px)').matches) {
    const d = document.querySelector('.painel-det');
    if (d) d.scrollIntoView({block: 'start'});
  }
};

/* ---- eventos (delegados: as telas são redesenhadas inteiras) ---- */
document.addEventListener('click', ev => {
  const t = ev.target.closest('[data-acao]');
  if (!t) return;
  const box = t.closest('[data-form]'), q = box && box.dataset.form, f = q && formAtivo(q);
  const ai = +t.dataset.ai, pi = +t.dataset.pi;
  switch (t.dataset.acao) {
    case 'aba':
      if (t.dataset.v === aba || !podeLargarEdicao()) return;
      aba = t.dataset.v; lembra('aba', aba); edicao = null; registrada = null; render(); break;
    case 'visao':
      if (t.dataset.v === visao || !podeLargarEdicao()) return;
      visao = t.dataset.v; lembra('visao', visao); edicao = null; filtro.busca = ''; render(); break;
    case 'ano':
      if (!podeLargarEdicao()) return;
      filtro.ano = t.dataset.v; edicao = null; render(); break;
    case 'limpa-busca':
      filtro.busca = ''; render(); foca('busca'); break;
    case 'sel-expo':
      if (+t.dataset.id === selExpo || !podeLargarEdicao()) return;
      selExpo = +t.dataset.id; edicao = null; salvoAgora = null; renderCorpo(); rolaParaDetalhe(); break;
    case 'sel-animal':
      selAnimal = t.dataset.animal; renderCorpo(); rolaParaDetalhe(); break;
    case 'ir-animal':
      if (!podeLargarEdicao()) return;
      visao = 'animal'; lembra('visao', visao); selAnimal = t.dataset.animal; filtro.busca = ''; edicao = null; render(); break;
    case 'ir-expo': {
      const e = EXPOS.find(x => x.id === +t.dataset.id);
      visao = 'expo'; lembra('visao', visao); selExpo = e.id; filtro.busca = ''; salvoAgora = null;
      if (filtro.ano && filtro.ano !== String(dt(e.inicio).a)) filtro.ano = String(dt(e.inicio).a);
      render(); break;
    }
    case 'ver-registro': {
      const e = EXPOS.find(x => x.id === +t.dataset.id);
      aba = 'historico'; lembra('aba', aba); visao = 'expo'; selExpo = e.id; filtro = {ano: String(dt(e.inicio).a), busca: ''};
      registrada = null; render(); break;
    }
    case 'editar':
      edicao = formDe(EXPOS.find(x => x.id === +t.dataset.id)); edicaoOriginal = JSON.stringify(edicao);
      salvoAgora = null; renderCorpo(); foca('edicao-nome'); break;
    case 'cancelar-edicao':
      if (!podeLargarEdicao()) return;
      edicao = null; renderCorpo(); break;
    case 'limpar':
      if (preenchido(rascunho) && !confirm('Apagar o que foi preenchido?')) return;
      rascunho = novoForm(); registrada = null; render(); break;
    case 'periodo':
      f.so_mes = t.dataset.v === 'mes';
      if (f.so_mes) { const m = f.inicio.slice(0, 7); f.inicio = `${m}-01`; f.fim = fimDoMes(m); }
      redesenhaForm(q); break;
    case 'status':
      f.status = t.dataset.v; redesenhaForm(q); break;
    case 'add-animal':
      f.animais.push(animalVazio()); redesenhaForm(q); foca(`${q}-a${f.animais.length - 1}`); break;
    case 'rm-animal':
      f.animais.splice(ai, 1);
      if (!f.animais.length) f.animais.push(animalVazio());
      redesenhaForm(q); break;
    case 'add-premio':
      f.animais[ai].premios.push(premioVazio()); redesenhaForm(q);
      foca(`${q}-a${ai}-p${f.animais[ai].premios.length - 1}`); break;
    case 'rm-premio':
      f.animais[ai].premios.splice(pi, 1);
      if (!f.animais[ai].premios.length) f.animais[ai].premios.push(premioVazio());
      redesenhaForm(q); break;
    case 'salvar': salvar(q); break;
    case 'excluir': excluir(); break;
  }
});

document.addEventListener('input', ev => {
  const t = ev.target, box = t.closest('[data-form]');
  if (box) {
    escreve(formAtivo(box.dataset.form), t, box);
    const c = box.querySelector('[data-cont]');
    if (c) c.textContent = contagem(formAtivo(box.dataset.form));
    return;
  }
  if (t.id === 'busca') {
    filtro.busca = t.value;
    const x = document.querySelector('[data-acao="limpa-busca"]');
    if (x) x.hidden = !t.value;
    renderCorpo();
  }
});

/* Enter no animal desce para o primeiro prêmio; Enter no último prêmio abre o próximo */
document.addEventListener('keydown', ev => {
  if (ev.key !== 'Enter') return;
  const t = ev.target, box = t.closest && t.closest('[data-form]');
  if (!box || t.tagName === 'TEXTAREA') return;
  const q = box.dataset.form, f = formAtivo(q);
  if (t.dataset.c === 'animal') {
    ev.preventDefault(); foca(`${q}-a${t.dataset.ai}-p0`);
  } else if (t.dataset.pi != null) {
    ev.preventDefault();
    const ai = +t.dataset.ai, pi = +t.dataset.pi, a = f.animais[ai];
    if (pi === a.premios.length - 1) {
      if (!t.value.trim()) return;
      a.premios.push(premioVazio()); redesenhaForm(q);
    }
    foca(`${q}-a${ai}-p${pi + 1}`);
  }
});

window.addEventListener('beforeunload', ev => {
  if (edicaoMudou() || (rascunho && preenchido(rascunho))) { ev.preventDefault(); ev.returnValue = ''; }
});

/* ---- início ---- */
(async function () {
  if (!cliente()) {
    $('painel').innerHTML = '<div class="vazio">Abra o SGPG pelo hub: a sessão vem de lá.</div>';
    return;
  }
  try {
    await Promise.all([checaEditor(), carrega()]);
    rascunho = novoForm();
    const a = lembrado('aba');
    aba = (a === 'registro' || a === 'historico') ? a : (souEditor ? 'registro' : 'historico');
    render();
  } catch (err) {
    $('painel').innerHTML = `<div class="vazio">Não foi possível carregar: ${esc(err.message || err)}</div>`;
  }
})();
