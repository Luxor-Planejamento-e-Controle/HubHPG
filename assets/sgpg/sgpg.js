/* SGPG — Sistema de Gestão Pao Grande. Roda em iframe do hub (mesma origem).

   O SGPG é o sistema da Pao Grande inteira, montado módulo a módulo: cada módulo é uma
   aba do sistema (MODULOS). O primeiro é Exposições e premiações — a consulta por
   exposição ou por animal, com o botão que abre a tela de registro. Corrigir um
   registro é pela ficha da exposição, na mesma tela.
   O comitê mensal lê este registro (build_comite.py e deck.js, a partir de ago/2026).

   O prêmio aponta para um animal do plantel (sgpg_animal, sincronizado do controle
   mensal do haras pelo tools/sync_sgpg_animais.py): o registro só oferece quem está no
   plantel; quem saiu continua no histórico com o que ganhou.

   Sessão e cliente do Supabase vêm do hub (window.parent.HUB). Quem grava é decidido
   pelo banco (hub_sgpg_editor: admin e sgpg_editores); aqui só se escondem os botões
   que o banco recusaria. */
'use strict';

const STATUS = ['Próxima', 'Aguardando', 'Realizada', 'Cancelada'];
const CLS_STATUS = {'Próxima': 'st-prox', 'Aguardando': 'st-aguard', 'Realizada': 'st-real', 'Cancelada': 'st-canc'};
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho',
               'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const ABR = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
// Módulo novo entra aqui e ganha uma função de render em RENDER_MODULO.
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

let EXPOS = [], PREM = [], ANIMAL = new Map(), PLANTEL = [], souEditor = false;
let modulo = MODULOS[0].id;
let visao = lembrado('visao') === 'animal' ? 'animal' : 'expo';
let filtro = {ano: String(new Date().getFullYear()), busca: ''};
let selExpo = null, selAnimal = null;             // selAnimal é a chave do grupo do animal
let tela = null, telaOriginal = '';               // {modo: 'novo'|'edicao', form}
let aviso = null;                                 // {id, texto}: recado na ficha depois de gravar

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

/* ---- animais ---- */
const categoria = c => c ? c.charAt(0) + c.slice(1).toLowerCase() : '';
// prêmio de antes do cadastro do plantel não tem animal_id: agrupa pelo nome escrito
const grupoAnimal = p => p.animal_id ? `id${p.animal_id}` : `nome:${norm(p.animal)}`;
const nomeAnimal = p => (p.animal_id && ANIMAL.has(p.animal_id)) ? ANIMAL.get(p.animal_id).nome : p.animal;
function animais(exps) {
  const porId = new Map(exps.map(e => [e.id, e])), mapa = new Map();
  for (const p of PREM) {
    const e = porId.get(p.exposicao_id);
    if (!e) continue;
    const k = grupoAnimal(p);
    if (!mapa.has(k)) mapa.set(k, {chave: k, nome: nomeAnimal(p), cad: ANIMAL.get(p.animal_id) || null, ps: []});
    mapa.get(k).ps.push({...p, e});
  }
  return [...mapa.values()].map(a => ({...a, nExp: new Set(a.ps.map(p => p.exposicao_id)).size}));
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
  const [e, p, a] = await Promise.all([
    c.from('sgpg_exposicao').select('*').order('inicio'),
    c.from('sgpg_premiacao').select('*').order('exposicao_id').order('ordem'),
    c.from('sgpg_animal').select('id,nome,categoria,no_plantel').order('nome'),
  ]);
  for (const r of [e, p, a]) if (r.error) throw r.error;
  EXPOS = e.data || []; PREM = p.data || [];
  ANIMAL = new Map((a.data || []).map(x => [x.id, x]));
  PLANTEL = (a.data || []).filter(x => x.no_plantel);
}
const premiosDe = id => PREM.filter(p => p.exposicao_id === id);
function agrupaPorAnimal(ps) {
  const out = [], idx = new Map();
  for (const p of ps) {
    const k = grupoAnimal(p);
    if (!idx.has(k)) { idx.set(k, out.length); out.push({chave: k, animal_id: p.animal_id, animal: nomeAnimal(p), premios: []}); }
    out[idx.get(k)].premios.push(p);
  }
  return out;
}

/* ---- tela de registro (nova exposição e correção) ---- */
const premioVazio = () => ({id: null, premio: ''});
const animalVazio = () => ({animal_id: null, animal: '', premios: [premioVazio()]});
function novoForm() {
  const h = hojeIso();
  return {id: null, nome: '', inicio: h, fim: h, so_mes: false, local: '', status: 'Próxima',
          observacao: '', animais: [animalVazio()]};
}
function formDe(e) {
  const g = agrupaPorAnimal(premiosDe(e.id));
  return {id: e.id, nome: e.nome, inicio: e.inicio, fim: e.fim, so_mes: e.so_mes, local: e.local || '',
          status: e.status, observacao: e.observacao || '',
          animais: g.length ? g.map(x => ({animal_id: x.animal_id, animal: x.animal,
                                           premios: x.premios.map(p => ({id: p.id, premio: p.premio}))}))
                            : [animalVazio()]};
}
const telaMudou = () => !!tela && JSON.stringify(tela.form) !== telaOriginal;
const podeFecharTela = () => !telaMudou()
  || confirm(tela.modo === 'novo' ? 'Descartar o que foi preenchido?' : 'Descartar as alterações deste registro?');
function abreTela(modo, e) {
  tela = {modo, form: modo === 'novo' ? novoForm() : formDe(e)};
  telaOriginal = JSON.stringify(tela.form);
  renderTela();
  foca('f-nome');
}
function fechaTela() {
  tela = null;
  renderTela();
}

function contagem(f) {
  const nPrem = f.animais.reduce((s, a) => s + (a.animal_id ? a.premios.filter(p => p.premio.trim()).length : 0), 0);
  const nAnim = f.animais.filter(a => a.animal_id).length;
  return nAnim ? `${plural(nPrem, 'prêmio', 'prêmios')} · ${plural(nAnim, 'animal', 'animais')}` : '';
}

const animalHtml = (a, ai) => `<div class="animal">
    <div class="animal-cab">
      <div class="combo">
        <input type="text" id="f-a${ai}" data-ai="${ai}" data-c="animal" value="${esc(a.animal)}"
          class="${a.animal.trim() && !a.animal_id ? 'invalido' : ''}" placeholder="Buscar animal do plantel"
          autocomplete="off" role="combobox" aria-autocomplete="list" aria-expanded="false"
          aria-controls="f-a${ai}-opcoes" aria-label="Animal">
        ${a.animal_id && ANIMAL.has(a.animal_id) ? `<span class="combo-tag">${esc(categoria(ANIMAL.get(a.animal_id).categoria))}</span>` : ''}
        <div class="opcoes" id="f-a${ai}-opcoes" role="listbox" hidden></div>
      </div>
      <button type="button" class="x" data-acao="rm-animal" data-ai="${ai}" title="Remover animal" aria-label="Remover animal">×</button>
    </div>
    ${a.premios.map((p, pi) => `<div class="premio">
      <input type="text" id="f-a${ai}-p${pi}" data-ai="${ai}" data-pi="${pi}" value="${esc(p.premio)}"
        placeholder="Prêmio, ex.: Campeã de Marcha" autocomplete="off" aria-label="Prêmio">
      <button type="button" class="x" data-acao="rm-premio" data-ai="${ai}" data-pi="${pi}" title="Remover prêmio" aria-label="Remover prêmio">×</button>
    </div>`).join('')}
    <button type="button" class="bt link" data-acao="add-premio" data-ai="${ai}">+ Prêmio</button>
  </div>`;

function formHtml(f, modo) {
  const datas = f.so_mes
    ? `<input type="month" id="f-mes" data-f="mes" value="${f.inicio.slice(0, 7)}" aria-label="Mês">`
    : `<input type="date" id="f-ini" data-f="inicio" value="${f.inicio}" aria-label="Início">
       <span class="ate">a</span>
       <input type="date" id="f-fim" data-f="fim" value="${f.fim}" aria-label="Fim">`;
  return `<div class="registro" data-form>
    <section class="cartao">
      <div class="cartao-cab"><h2>Exposição</h2></div>
      <div class="campos">
        <label class="campo"><span class="rot">Nome</span>
          <input type="text" id="f-nome" data-f="nome" value="${esc(f.nome)}"
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
          <input type="text" id="f-local" data-f="local" list="lista-locais" value="${esc(f.local)}"
            placeholder="Cidade ou parque" autocomplete="off"></label>
        <div class="campo"><span class="rot">Status</span>
          <div class="seg status" role="group" aria-label="Status">${STATUS.map(s =>
            `<button type="button" class="${s === f.status ? 'on ' : ''}${CLS_STATUS[s]}" data-acao="status" data-v="${s}">${s}</button>`).join('')}</div></div>
        <label class="campo"><span class="rot">Observação</span>
          <textarea id="f-obs" data-f="observacao" rows="2">${esc(f.observacao)}</textarea></label>
      </div>
    </section>
    <section class="cartao">
      <div class="cartao-cab"><h2>Premiações</h2><span class="cont" data-cont>${contagem(f)}</span></div>
      <div class="animais">${f.animais.map(animalHtml).join('')}</div>
      <button type="button" class="bt" data-acao="add-animal">+ Animal</button>
    </section>
    <div class="acoes">
      ${modo === 'edicao' ? '<button type="button" class="bt perigo" data-acao="excluir">Excluir registro</button>' : ''}
      <span class="msg" data-msg aria-live="polite"></span>
      <span class="dir">
        <button type="button" class="bt" data-acao="fecha-tela">Cancelar</button>
        <button type="button" class="bt prim" data-acao="salvar">${modo === 'novo' ? 'Registrar exposição' : 'Salvar alterações'}</button>
      </span>
    </div>
  </div>`;
}

function renderTela() {
  const el = $('tela');
  if (!tela) { el.hidden = true; el.innerHTML = ''; return; }
  el.hidden = false;
  el.innerHTML = `<div class="tela-cab">
      <div><div class="tela-kicker">${esc(MODULOS.find(m => m.id === modulo).titulo)}</div>
        <h2 id="tela-titulo">${tela.modo === 'novo' ? 'Registrar exposição' : 'Editar registro'}</h2></div>
      <button type="button" class="x grande" data-acao="fecha-tela" title="Fechar" aria-label="Fechar">×</button>
    </div>
    <div class="tela-corpo">${formHtml(tela.form, tela.modo)}</div>`;
}

const foca = id => {
  const x = document.getElementById(id);
  if (!x) return;
  x.focus();
  if (x.type === 'text') x.setSelectionRange(x.value.length, x.value.length);
};
function redesenhaForm() {
  const box = document.querySelector('[data-form]');
  if (!box) return;
  const foco = document.activeElement && document.activeElement.id;
  box.outerHTML = formHtml(tela.form, tela.modo);
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
    // escreveu o nome inteiro de um animal do plantel: vale como escolha
    const a = f.animais[+t.dataset.ai], exato = PLANTEL.find(x => norm(x.nome) === norm(t.value.trim()));
    a.animal = t.value;
    a.animal_id = exato ? exato.id : null;
  } else if (t.dataset.pi != null) {
    f.animais[+t.dataset.ai].premios[+t.dataset.pi].premio = t.value;
  }
}

/* ---- busca de animal do plantel ---- */
function opcoesAnimal(f, ai) {
  const a = f.animais[ai];
  const usados = new Set(f.animais.filter((x, i) => i !== ai && x.animal_id).map(x => x.animal_id));
  const termos = norm(a.animal).split(/\s+/).filter(Boolean);
  return PLANTEL.filter(x => !usados.has(x.id) && termos.every(t => norm(x.nome).includes(t))).slice(0, 8);
}
function abreOpcoes(input) {
  const ai = +input.dataset.ai, lista = $(`f-a${ai}-opcoes`);
  if (!lista) return;
  const ops = opcoesAnimal(tela.form, ai);
  lista.innerHTML = ops.length
    ? ops.map((x, i) => `<button type="button" class="opcao${i === 0 ? ' ativa' : ''}" role="option" tabindex="-1"
        data-acao="escolhe-animal" data-ai="${ai}" data-id="${x.id}"><span>${esc(x.nome)}</span><small>${esc(categoria(x.categoria))}</small></button>`).join('')
    : `<div class="opcao-vazia">${PLANTEL.length ? 'Nenhum animal do plantel com esse nome.' : 'A lista do plantel ainda não foi carregada.'}</div>`;
  lista.hidden = false;
  input.setAttribute('aria-expanded', 'true');
}
function fechaOpcoes(input) {
  const lista = $(`f-a${input.dataset.ai}-opcoes`);
  if (lista) lista.hidden = true;
  input.setAttribute('aria-expanded', 'false');
}
function escolheAnimal(ai, id) {
  const a = tela.form.animais[ai];
  a.animal_id = id; a.animal = ANIMAL.get(id).nome;
  redesenhaForm();
  foca(`f-a${ai}-p0`);
}

function linhasDe(f) {
  const out = [];
  for (const a of f.animais) {
    const texto = a.animal.trim();
    const ps = a.premios.map(p => ({...p, premio: p.premio.trim()})).filter(p => p.premio);
    if (!texto && !a.animal_id && !ps.length) continue;
    if (!a.animal_id) {
      return {erro: texto ? `"${texto}" não está no plantel. Escolha o animal na lista.`
                          : 'Tem prêmio sem animal. Escolha o animal ou remova o prêmio.'};
    }
    const nome = ANIMAL.get(a.animal_id).nome;
    if (!ps.length) return {erro: `Informe ao menos um prêmio de ${nome}, ou remova o animal.`};
    ps.forEach(p => out.push({id: p.id, animal_id: a.animal_id, animal: nome, premio: p.premio}));
  }
  return {linhas: out};
}

async function salvar() {
  const f = tela.form, box = document.querySelector('[data-form]'), c = cliente();
  const dizer = (txt, erro) => { const m = box.querySelector('[data-msg]'); m.textContent = txt; m.classList.toggle('erro', !!erro); };
  const nome = f.nome.trim();
  if (!nome) { foca('f-nome'); return dizer('Informe o nome da exposição.', true); }
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
      const campos = {animal_id: p.animal_id, animal: p.animal, premio: p.premio, ordem};
      const r = p.id
        ? await c.from('sgpg_premiacao').update(campos).eq('id', p.id)
        : await c.from('sgpg_premiacao').insert({exposicao_id: id, ...campos});
      if (r.error) throw r.error;
    }
    await carrega();
    aviso = {id, texto: tela.modo === 'novo' ? 'Exposição registrada.' : 'Alterações salvas.'};
    // a ficha da exposição gravada fica aberta, no ano dela
    visao = 'expo'; lembra('visao', visao); selExpo = id; filtro.busca = '';
    if (filtro.ano && filtro.ano !== f.inicio.slice(0, 4)) filtro.ano = f.inicio.slice(0, 4);
    fechaTela();
    render();
  } catch (err) {
    btn.disabled = false;
    dizer('Não foi possível salvar: ' + (err.message || err), true);
  }
}

async function excluir() {
  const e = EXPOS.find(x => x.id === tela.form.id);
  const n = premiosDe(e.id).length;
  if (!confirm(`Excluir "${e.nome}"${n ? ` e ${plural(n, 'premiação', 'premiações')}` : ''}? Não dá para desfazer.`)) return;
  const r = await cliente().from('sgpg_exposicao').delete().eq('id', e.id);
  if (r.error) {
    const m = document.querySelector('[data-form] [data-msg]');
    m.textContent = 'Não foi possível excluir: ' + r.error.message; m.classList.add('erro');
    return;
  }
  selExpo = null; aviso = null;
  await carrega();
  fechaTela();
  render();
}

/* ---- sistema: abas dos módulos ---- */
function render() {
  $('abas').innerHTML = MODULOS.map(m =>
    `<button type="button" class="${m.id === modulo ? 'on' : ''}" data-acao="modulo" data-v="${m.id}">${esc(m.titulo)}</button>`).join('');
  const locais = [...new Set(EXPOS.map(e => e.local).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  $('listas').innerHTML = `<datalist id="lista-locais">${locais.map(l => `<option value="${esc(l)}">`).join('')}</datalist>`;
  RENDER_MODULO[modulo]($('painel'));
}

/* ---- módulo Exposições e premiações ---- */
const anos = () => [...new Set(EXPOS.map(e => dt(e.inicio).a).concat(new Date().getFullYear()))].sort((a, b) => b - a);
const escopo = () => EXPOS.filter(e => !filtro.ano || String(dt(e.inicio).a) === filtro.ano);

function renderExposicoes(el) {
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
      ${souEditor ? '<button type="button" class="bt prim" data-acao="nova">+ Registrar exposição</button>' : ''}
    </div>
    <div class="kpis">${kpis.map(([r, v]) => `<div class="kpi"><span class="rot">${r}</span><b>${v}</b></div>`).join('')}</div>
    <div id="hist-corpo"></div>`;
  renderCorpo();
}
const RENDER_MODULO = {exposicoes: renderExposicoes};

function renderCorpo() {
  const el = $('hist-corpo');
  if (el) (visao === 'expo' ? corpoExpo : corpoAnimal)(el);
}

function corpoExpo(el) {
  const q = norm(filtro.busca);
  const exps = escopo().filter(e => !q || [e.nome, e.local, ...premiosDe(e.id).flatMap(p => [nomeAnimal(p), p.premio])]
      .some(s => norm(s).includes(q)))
    .sort((a, b) => b.inicio.localeCompare(a.inicio));
  if (!exps.length) {
    el.innerHTML = `<div class="vazio">${q ? 'Nada encontrado com essa busca.'
      : `Nenhuma exposição registrada${filtro.ano ? ' em ' + filtro.ano : ''}.`}</div>`;
    return;
  }
  if (!exps.some(e => e.id === selExpo)) selExpo = exps[0].id;
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
  el.innerHTML = `<div class="md"><nav class="lista" aria-label="Exposições">${lista}</nav>
    <div class="painel-det">${detalheExpo(EXPOS.find(x => x.id === selExpo))}</div></div>`;
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
            <button type="button" class="lnk" data-acao="ir-animal" data-animal="${esc(g.chave)}">${esc(g.animal)}</button>
            <span class="cont">${g.premios.length}</span>
          </div>
          <ul class="premios">${g.premios.map(p => `<li>${esc(p.premio)}</li>`).join('')}</ul>
        </div>`).join('')}</div>` : `<div class="vazio">${semPremio}</div>`}
    </section>
    <footer class="det-pe">
      ${alt ? `<span class="quando">Última alteração em ${alt}</span>` : ''}
      ${aviso && aviso.id === e.id ? `<span class="msg ok" role="status">${esc(aviso.texto)}</span>` : ''}
      ${souEditor ? `<span class="dir"><button type="button" class="bt" data-acao="editar" data-id="${e.id}">Editar registro</button></span>` : ''}
    </footer>
  </article>`;
}

function corpoAnimal(el) {
  const q = norm(filtro.busca);
  const lista = animais(escopo())
    .filter(a => !q || norm(a.nome).includes(q) || a.ps.some(p => norm(p.premio).includes(q)))
    .sort((a, b) => b.ps.length - a.ps.length || a.nome.localeCompare(b.nome, 'pt-BR'));
  if (!lista.length) {
    el.innerHTML = `<div class="vazio">${q ? 'Nenhum animal encontrado com essa busca.'
      : `Nenhum animal premiado${filtro.ano ? ' em ' + filtro.ano : ''}.`}</div>`;
    return;
  }
  if (!lista.some(a => a.chave === selAnimal)) selAnimal = lista[0].chave;
  const max = lista[0].ps.length;
  const itens = lista.map(a => `<button type="button" class="it it-animal${a.chave === selAnimal ? ' sel' : ''}" data-acao="sel-animal" data-animal="${esc(a.chave)}">
      <span class="inicial" aria-hidden="true">${esc(a.nome.normalize('NFD').charAt(0).toUpperCase())}</span>
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
      <div class="det-kicker">${a.cad ? `${esc(categoria(a.cad.categoria))} · ${a.cad.no_plantel ? 'no plantel' : 'fora do plantel'}` : 'fora do plantel'}</div>
      <h2>${esc(a.nome)}</h2>
    </div></header>
    <div class="mini-kpis">
      <div><b>${a.ps.length}</b><span>Premiações${filtro.ano ? ' em ' + filtro.ano : ''}</span></div>
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
document.addEventListener('mousedown', ev => {
  // escolher na lista não pode tirar o foco do campo (o blur fecharia a lista antes do clique)
  if (ev.target.closest('.opcao')) ev.preventDefault();
});

document.addEventListener('click', ev => {
  const t = ev.target.closest('[data-acao]');
  if (!t) return;
  const f = tela && tela.form, ai = +t.dataset.ai, pi = +t.dataset.pi;
  switch (t.dataset.acao) {
    case 'modulo':
      modulo = t.dataset.v; render(); break;
    case 'visao':
      if (t.dataset.v === visao) return;
      visao = t.dataset.v; lembra('visao', visao); filtro.busca = ''; render(); break;
    case 'ano':
      filtro.ano = t.dataset.v; render(); break;
    case 'limpa-busca':
      filtro.busca = ''; render(); foca('busca'); break;
    case 'sel-expo':
      selExpo = +t.dataset.id; aviso = null; renderCorpo(); rolaParaDetalhe(); break;
    case 'sel-animal':
      selAnimal = t.dataset.animal; renderCorpo(); rolaParaDetalhe(); break;
    case 'ir-animal':
      visao = 'animal'; lembra('visao', visao); selAnimal = t.dataset.animal; filtro.busca = ''; render(); break;
    case 'ir-expo': {
      const e = EXPOS.find(x => x.id === +t.dataset.id);
      visao = 'expo'; lembra('visao', visao); selExpo = e.id; filtro.busca = ''; aviso = null;
      if (filtro.ano && filtro.ano !== String(dt(e.inicio).a)) filtro.ano = String(dt(e.inicio).a);
      render(); break;
    }
    case 'nova': abreTela('novo'); break;
    case 'editar': abreTela('edicao', EXPOS.find(x => x.id === +t.dataset.id)); break;
    case 'fecha-tela': if (podeFecharTela()) fechaTela(); break;
    case 'periodo':
      f.so_mes = t.dataset.v === 'mes';
      if (f.so_mes) { const m = f.inicio.slice(0, 7); f.inicio = `${m}-01`; f.fim = fimDoMes(m); }
      redesenhaForm(); break;
    case 'status':
      f.status = t.dataset.v; redesenhaForm(); break;
    case 'add-animal':
      f.animais.push(animalVazio()); redesenhaForm(); foca(`f-a${f.animais.length - 1}`); break;
    case 'rm-animal':
      f.animais.splice(ai, 1);
      if (!f.animais.length) f.animais.push(animalVazio());
      redesenhaForm(); break;
    case 'add-premio':
      f.animais[ai].premios.push(premioVazio()); redesenhaForm();
      foca(`f-a${ai}-p${f.animais[ai].premios.length - 1}`); break;
    case 'rm-premio':
      f.animais[ai].premios.splice(pi, 1);
      if (!f.animais[ai].premios.length) f.animais[ai].premios.push(premioVazio());
      redesenhaForm(); break;
    case 'escolhe-animal': escolheAnimal(ai, +t.dataset.id); break;
    case 'salvar': salvar(); break;
    case 'excluir': excluir(); break;
  }
});

document.addEventListener('input', ev => {
  const t = ev.target, box = t.closest('[data-form]');
  if (box) {
    escreve(tela.form, t, box);
    if (t.dataset.c === 'animal') {
      t.classList.remove('invalido');
      const tag = t.parentElement.querySelector('.combo-tag');
      if (tag) tag.remove();
      abreOpcoes(t);
    }
    const c = box.querySelector('[data-cont]');
    if (c) c.textContent = contagem(tela.form);
    return;
  }
  if (t.id === 'busca') {
    filtro.busca = t.value;
    const x = document.querySelector('[data-acao="limpa-busca"]');
    if (x) x.hidden = !t.value;
    renderCorpo();
  }
});

document.addEventListener('focusin', ev => {
  if (ev.target.dataset && ev.target.dataset.c === 'animal' && tela) abreOpcoes(ev.target);
});
document.addEventListener('focusout', ev => {
  const t = ev.target;
  if (!(t.dataset && t.dataset.c === 'animal') || !tela) return;
  fechaOpcoes(t);
  const a = tela.form.animais[+t.dataset.ai];
  if (a) t.classList.toggle('invalido', !!a.animal.trim() && !a.animal_id);
});

document.addEventListener('keydown', ev => {
  const t = ev.target;
  if (ev.key === 'Escape') {
    if (t.dataset && t.dataset.c === 'animal' && !$(`f-a${t.dataset.ai}-opcoes`).hidden) { fechaOpcoes(t); return; }
    if (tela && podeFecharTela()) fechaTela();
    return;
  }
  if (!tela || !t.closest || !t.closest('[data-form]') || t.tagName === 'TEXTAREA') return;
  const f = tela.form;
  if (t.dataset.c === 'animal') {
    const ai = +t.dataset.ai, lista = $(`f-a${ai}-opcoes`);
    const ops = [...lista.querySelectorAll('.opcao')], i = ops.findIndex(o => o.classList.contains('ativa'));
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
      ev.preventDefault();
      if (lista.hidden) { abreOpcoes(t); return; }
      if (!ops.length) return;
      const j = ev.key === 'ArrowDown' ? Math.min(ops.length - 1, i + 1) : Math.max(0, i - 1);
      ops.forEach((o, k) => o.classList.toggle('ativa', k === j));
      ops[j].scrollIntoView({block: 'nearest'});
    } else if (ev.key === 'Enter') {
      ev.preventDefault();
      if (!lista.hidden && i >= 0) escolheAnimal(ai, +ops[i].dataset.id);
      else if (f.animais[ai].animal_id) foca(`f-a${ai}-p0`);
    }
  } else if (ev.key === 'Enter' && t.dataset.pi != null) {
    // Enter no último prêmio abre o próximo
    ev.preventDefault();
    const ai = +t.dataset.ai, pi = +t.dataset.pi, a = f.animais[ai];
    if (pi === a.premios.length - 1) {
      if (!t.value.trim()) return;
      a.premios.push(premioVazio()); redesenhaForm();
    }
    foca(`f-a${ai}-p${pi + 1}`);
  }
});

window.addEventListener('beforeunload', ev => {
  if (telaMudou()) { ev.preventDefault(); ev.returnValue = ''; }
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
    $('painel').innerHTML = `<div class="vazio">Não foi possível carregar: ${esc(err.message || err)}</div>`;
  }
})();
