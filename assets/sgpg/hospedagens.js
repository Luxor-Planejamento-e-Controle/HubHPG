/* SGPG — módulo Hospedagens: as estadias na Casa da Fazenda Pao Grande.

   Carregado depois do sgpg.js e usa o que ele já tem (MODULOS, RENDER_MODULO, render,
   esc, cliente, souEditor, periodo, seloHtml...). Fica em arquivo próprio para não
   misturar com o módulo de exposições: as ações daqui começam com "h-", o formulário
   é [data-hform] e a tela é `htela` — nada disso cruza com o que o sgpg.js trata.

   A acomodação é uma só, a Casa: não há quarto, suíte ou chalé a escolher.

   A consulta tem duas visões: Agenda (o mapa do mês da Casa e as próximas chegadas — é
   a agenda futura) e Lista (histórico por ano, lista e ficha, igual às exposições).
   Registro e correção abrem por botão, numa tela por cima.

   Tabela: sgpg_hospedagem (migration 20261119100000). */
'use strict';

const H_TIPOS = ['Família', 'Locação', 'Visita', 'Outro'];
/* cores dos tipos, validadas com dataviz/validate_palette.js no modo escuro sobre #0A3050
   (todas acima de 3:1); a barra da agenda leva o nome escrito, então a cor não é o único
   sinal */
const H_COR = {'Família': '#B88600', 'Locação': '#3D8BD9', 'Visita': '#199e70', 'Outro': '#9085e9'};
const H_STATUS = ['Prevista', 'Confirmada', 'Realizada', 'Cancelada'];
const H_CLS = {'Prevista': 'st-prox', 'Confirmada': 'st-aguard', 'Realizada': 'st-real', 'Cancelada': 'st-canc'};
const DIA_SEM = ['D', 'S', 'T', 'Q', 'Q', 'S', 'S'];

let HOSP = [], hCarregado = false, hErro = null;
let hVisao = lembrado('h.visao') === 'lista' ? 'lista' : 'agenda';
let hMes = hojeIso().slice(0, 7);
let hFiltro = {ano: String(new Date().getFullYear()), busca: ''};
let hSel = null, hAviso = null;
let htela = null, htelaOriginal = '';

MODULOS.push({id: 'hospedagens', titulo: 'Hospedagens'});
RENDER_MODULO.hospedagens = renderHospedagens;

/* ---- datas ---- */
const diaUTC = iso => { const {a, m, d} = dt(iso); return Date.UTC(a, m - 1, d); };
const isoDe = ms => { const x = new Date(ms); return `${x.getUTCFullYear()}-${dd(x.getUTCMonth() + 1)}-${dd(x.getUTCDate())}`; };
const somaDias = (iso, n) => isoDe(diaUTC(iso) + n * 864e5);
const noites = h => Math.round((diaUTC(h.saida) - diaUTC(h.entrada)) / 864e5);
const pessoas = h => (h.adultos || 0) + (h.criancas || 0);
const diasNoMes = aaaaMm => { const [a, m] = aaaaMm.split('-').map(Number); return new Date(a, m, 0).getDate(); };
const mesNome = aaaaMm => { const [a, m] = aaaaMm.split('-').map(Number); return `${MESES[m - 1]} ${a}`; };
const movMes = (aaaaMm, n) => { const [a, m] = aaaaMm.split('-').map(Number); const x = new Date(a, m - 1 + n, 1);
  return `${x.getFullYear()}-${dd(x.getMonth() + 1)}`; };
const hPeriodo = h => periodo({inicio: h.entrada, fim: h.saida, so_mes: false});
const hSelo = h => seloHtml({inicio: h.entrada, fim: h.saida, so_mes: false});
const txtNoites = n => n === 0 ? 'day use' : plural(n, 'noite', 'noites');
const txtPessoas = h => {
  const p = [];
  if (h.adultos) p.push(plural(h.adultos, 'adulto', 'adultos'));
  if (h.criancas) p.push(plural(h.criancas, 'criança', 'crianças'));
  return p.join(' e ');
};
const reais = v => v == null ? '' : 'R$ ' + Number(v).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const tipoHtml = t => `<span class="tipo-tag" style="--cor:${H_COR[t] || H_COR.Outro}">${esc(t)}</span>`;
const hStatusHtml = s => `<span class="status ${H_CLS[s] || ''}">${esc(s)}</span>`;
const ativa = h => h.status !== 'Cancelada';

/* noites de `h` que caem entre `ini` e `fim` (inclusive; a noite é a do dia da entrada) */
function noitesEntre(h, ini, fim) {
  const a = Math.max(diaUTC(h.entrada), diaUTC(ini)), b = Math.min(diaUTC(h.saida), diaUTC(fim) + 864e5);
  return Math.max(0, Math.round((b - a) / 864e5));
}
const cruza = (h, ini, fim) => h.entrada <= fim && h.saida >= ini;

/* ---- dados ---- */
async function hCarrega() {
  const c = cliente();
  if (!c) return;
  const r = await c.from('sgpg_hospedagem').select('*').order('entrada');
  hErro = r.error ? (r.error.message || String(r.error)) : null;
  HOSP = r.error ? [] : (r.data || []);
  hCarregado = true;
}

/* ---- consulta ---- */
function renderHospedagens(el) {
  if (!hCarregado) { el.innerHTML = '<div class="vazio">Carregando as hospedagens…</div>'; return; }
  if (hErro) {
    el.innerHTML = `<div class="vazio">Não foi possível carregar as hospedagens: ${esc(hErro)}</div>`;
    return;
  }
  const lista = hVisao === 'lista';
  const anos = [...new Set(HOSP.map(h => dt(h.entrada).a).concat(new Date().getFullYear()))].sort((a, b) => b - a);
  // os números acompanham o que está na tela: o mês da agenda ou o ano da lista
  const ini = lista ? (hFiltro.ano ? `${hFiltro.ano}-01-01` : '0000-01-01') : `${hMes}-01`;
  const fim = lista ? (hFiltro.ano ? `${hFiltro.ano}-12-31` : '9999-12-31') : fimDoMes(hMes);
  const doPer = HOSP.filter(h => ativa(h) && cruza(h, ini, fim));
  const nNoites = doPer.reduce((s, h) => s + noitesEntre(h, ini, fim), 0);
  const nPessoas = doPer.reduce((s, h) => s + pessoas(h), 0);
  let quarto;
  if (lista) {
    const prox = HOSP.filter(h => ativa(h) && h.entrada >= hojeIso()).sort((a, b) => a.entrada.localeCompare(b.entrada))[0];
    quarto = ['Próxima chegada', prox ? `${dd(dt(prox.entrada).d)}/${dd(dt(prox.entrada).m)}` : '—'];
  } else {
    // dias do mês com alguém na casa (pernoite)
    const ocup = new Set();
    for (const h of doPer) {
      for (let d = Math.max(diaUTC(h.entrada), diaUTC(ini)); d < Math.min(diaUTC(h.saida), diaUTC(fim) + 864e5); d += 864e5) ocup.add(d);
    }
    quarto = ['Ocupação', `${Math.round(ocup.size / diasNoMes(hMes) * 100)}%`];
  }
  const kpis = [['Hospedagens', doPer.length], ['Noites', nNoites], ['Hóspedes', nPessoas], quarto];

  el.innerHTML = `
    <div class="hist-barra">
      <div class="seg" role="group" aria-label="Ver">
        <button type="button" class="${lista ? '' : 'on'}" data-acao="h-visao" data-v="agenda">Agenda</button>
        <button type="button" class="${lista ? 'on' : ''}" data-acao="h-visao" data-v="lista">Lista</button>
      </div>
      ${lista ? `<div class="anos" role="group" aria-label="Ano">
          <button type="button" class="${hFiltro.ano ? '' : 'on'}" data-acao="h-ano" data-v="">Todos</button>
          ${anos.map(a => `<button type="button" class="${String(a) === hFiltro.ano ? 'on' : ''}" data-acao="h-ano" data-v="${a}">${a}</button>`).join('')}
        </div>
        <div class="busca">
          <input type="search" id="h-busca" value="${esc(hFiltro.busca)}" autocomplete="off"
            placeholder="Buscar hóspede ou observação" aria-label="Buscar">
          <button type="button" class="x" data-acao="h-limpa-busca" title="Limpar busca" aria-label="Limpar busca" ${hFiltro.busca ? '' : 'hidden'}>×</button>
        </div>`
      : `<div class="h-mesnav" role="group" aria-label="Mês">
          <button type="button" class="bt" data-acao="h-mes" data-v="-1" aria-label="Mês anterior">‹</button>
          <span class="mes">${mesNome(hMes)}</span>
          <button type="button" class="bt" data-acao="h-mes" data-v="1" aria-label="Próximo mês">›</button>
          ${hMes !== hojeIso().slice(0, 7) ? '<button type="button" class="bt" data-acao="h-mes" data-v="0">Hoje</button>' : ''}
        </div>`}
      ${souEditor ? '<button type="button" class="bt prim h-nova" data-acao="h-nova">+ Registrar hospedagem</button>' : ''}
    </div>
    <div class="kpis">${kpis.map(([r, v]) => `<div class="kpi"><span class="rot">${r}</span><b>${v}</b></div>`).join('')}</div>
    <div id="h-corpo"></div>`;
  hRenderCorpo();
}

function hRenderCorpo() {
  const el = $('h-corpo');
  if (el) (hVisao === 'lista' ? hCorpoLista : hCorpoAgenda)(el);
}

/* faixas: estadias que se sobrepõem na Casa vão uma embaixo da outra */
function faixas(hs) {
  const lanes = [];
  for (const h of hs.slice().sort((a, b) => a.entrada.localeCompare(b.entrada))) {
    let i = lanes.findIndex(l => l[l.length - 1].saida < h.entrada);
    if (i < 0) { lanes.push([]); i = lanes.length - 1; }
    lanes[i].push(h);
  }
  return lanes;
}

function hCorpoAgenda(el) {
  const n = diasNoMes(hMes), ini = `${hMes}-01`, fim = fimDoMes(hMes), hoje = hojeIso();
  const doMes = HOSP.filter(h => ativa(h) && cruza(h, ini, fim));
  const dias = Array.from({length: n}, (_, i) => {
    const iso = `${hMes}-${dd(i + 1)}`, w = new Date(diaUTC(iso)).getUTCDay();
    return {iso, d: i + 1, w, fds: w === 0 || w === 6, hoje: iso === hoje};
  });
  const cab = dias.map(x => `<div class="ag-dia${x.fds ? ' fds' : ''}${x.hoje ? ' hoje' : ''}"><span>${DIA_SEM[x.w]}</span><b>${x.d}</b></div>`).join('');
  const fundo = dias.map(x => `<i class="${x.fds ? 'fds' : ''}${x.hoje ? ' hoje' : ''}"></i>`).join('');
  const linhas = (() => {
    const ls = faixas(doMes);
    const barras = ls.flatMap((l, li) => l.map(h => {
      const a = Math.max(0, Math.round((diaUTC(h.entrada) - diaUTC(ini)) / 864e5));
      const b = Math.min(n - 1, Math.round((diaUTC(h.saida) - diaUTC(ini)) / 864e5));
      const corta = [h.entrada < ini ? 'corta-ini' : '', h.saida > fim ? 'corta-fim' : ''].join(' ');
      const titulo = `${h.hospede} · ${hPeriodo(h)} · ${txtNoites(noites(h))}${pessoas(h) ? ' · ' + txtPessoas(h) : ''} · ${h.tipo} · ${h.status}`;
      return `<button type="button" class="ag-barra ${h.status === 'Prevista' ? 'prevista' : ''} ${corta}" data-acao="h-abre" data-id="${h.id}"
        style="--cor:${H_COR[h.tipo] || H_COR.Outro};left:${a / n * 100}%;width:${(b - a + 1) / n * 100}%;top:${6 + li * 30}px"
        title="${esc(titulo)}"><span class="nm">${esc(h.hospede)}</span>${pessoas(h) ? `<span class="n">${pessoas(h)}</span>` : ''}</button>`;
    })).join('');
    return `<div class="ag-linha"><div class="ag-rot">Casa</div>
      <div class="ag-faixa" style="height:${Math.max(1, ls.length) * 30 + 10}px"><div class="ag-fundo">${fundo}</div>${barras}</div></div>`;
  })();

  const hoje0 = hojeIso();
  const prox = HOSP.filter(h => ativa(h) && h.saida >= hoje0).sort((a, b) => a.entrada.localeCompare(b.entrada)).slice(0, 8);
  const itProx = h => `<button type="button" class="it" data-acao="h-abre" data-id="${h.id}">
      ${hSelo(h)}
      <span class="it-txt"><span class="it-nome">${esc(h.hospede)}</span>
        <span class="it-meta">${tipoHtml(h.tipo)} · ${esc(h.status)}${h.entrada <= hoje0 ? ' · na casa agora' : ''}</span></span>
      <span class="it-n" title="${txtNoites(noites(h))}">${noites(h) || ''}</span>
    </button>`;

  el.innerHTML = `
    <section class="agenda" style="--n:${n}">
      <div class="ag-grade">
        <div class="ag-linha ag-cab"><div class="ag-rot"></div><div class="ag-dias">${cab}</div></div>
        ${linhas}
      </div>
      ${doMes.length ? '' : `<div class="ag-vazio">Nenhuma hospedagem em ${mesNome(hMes).toLowerCase()}.</div>`}
      <div class="h-leg">${H_TIPOS.map(t => `<span><i style="--cor:${H_COR[t]}"></i>${t}</span>`).join('')}
        <span><i class="prev"></i>Prevista (a confirmar)</span></div>
    </section>
    <section class="h-prox">
      <h3 class="h-tit">Próximas hospedagens</h3>
      ${prox.length ? `<div class="lista">${prox.map(itProx).join('')}</div>`
        : `<div class="vazio">Nenhuma hospedagem marcada daqui pra frente.${souEditor ? ' Use <b>+ Registrar hospedagem</b>.' : ''}</div>`}
    </section>`;
}

function hEscopo() {
  return HOSP.filter(h => !hFiltro.ano || String(dt(h.entrada).a) === hFiltro.ano
    || String(dt(h.saida).a) === hFiltro.ano);
}

function hCorpoLista(el) {
  const q = norm(hFiltro.busca);
  const hs = hEscopo().filter(h => !q || [h.hospede, h.observacao, h.tipo].some(s => norm(s).includes(q)))
    .sort((a, b) => b.entrada.localeCompare(a.entrada));
  if (!hs.length) {
    el.innerHTML = `<div class="vazio">${q ? 'Nada encontrado com essa busca.'
      : `Nenhuma hospedagem registrada${hFiltro.ano ? ' em ' + hFiltro.ano : ''}.`}</div>`;
    return;
  }
  if (!hs.some(h => h.id === hSel)) hSel = hs[0].id;
  let mes = null, lista = '';
  for (const h of hs) {
    const m = h.entrada.slice(0, 7);
    if (m !== mes) { lista += `<div class="lista-ano">${mesNome(m).toUpperCase()}</div>`; mes = m; }
    lista += `<button type="button" class="it${h.id === hSel ? ' sel' : ''}" data-acao="h-sel" data-id="${h.id}">
      ${hSelo(h)}
      <span class="it-txt"><span class="it-nome">${esc(h.hospede)}</span>
        <span class="it-meta"><span class="ponto ${H_CLS[h.status]}"></span>${esc(h.status)} · ${tipoHtml(h.tipo)}</span></span>
      <span class="it-n" title="${txtNoites(noites(h))}">${noites(h) || ''}</span>
    </button>`;
  }
  el.innerHTML = `<div class="md"><nav class="lista" aria-label="Hospedagens">${lista}</nav>
    <div class="painel-det">${hDetalhe(HOSP.find(x => x.id === hSel))}</div></div>`;
}

function hDetalhe(h) {
  const alt = h.alterado_em ? new Date(h.alterado_em).toLocaleDateString('pt-BR') : '';
  const kpis = [['Noites', noites(h) || 'day use'], ['Hóspedes', pessoas(h) || '—'], ['Tipo', esc(h.tipo)]];
  if (h.valor != null) kpis.push(['Valor', reais(h.valor)]);
  return `<article class="det">
    <header class="det-cab">
      <div><div class="det-kicker">${esc(hPeriodo(h))}</div>
        <h2>${esc(h.hospede)}</h2></div>
      ${hStatusHtml(h.status)}
    </header>
    <div class="mini-kpis">${kpis.map(([r, v]) => `<div><b>${v}</b><span>${r}</span></div>`).join('')}</div>
    ${pessoas(h) ? `<p class="det-obs">${esc(txtPessoas(h))}</p>` : ''}
    ${h.observacao ? `<p class="det-obs">${esc(h.observacao)}</p>` : ''}
    <footer class="det-pe">
      ${alt ? `<span class="quando">Última alteração em ${alt}</span>` : ''}
      ${hAviso && hAviso.id === h.id ? `<span class="msg ok" role="status">${esc(hAviso.texto)}</span>` : ''}
      ${souEditor ? `<span class="dir"><button type="button" class="bt" data-acao="h-editar" data-id="${h.id}">Editar registro</button></span>` : ''}
    </footer>
  </article>`;
}

/* ---- tela de registro ---- */
function hNovoForm() {
  const h = hojeIso();
  return {id: null, tipo: 'Família', hospede: '', entrada: h, saida: somaDias(h, 1), adultos: '', criancas: '',
          status: 'Prevista', valor: '', observacao: ''};
}
function hFormDe(h) {
  return {id: h.id, tipo: h.tipo, hospede: h.hospede, entrada: h.entrada, saida: h.saida,
          adultos: h.adultos ? String(h.adultos) : '', criancas: h.criancas ? String(h.criancas) : '',
          status: h.status, valor: h.valor == null ? '' : String(h.valor),
          observacao: h.observacao || ''};
}
const htelaMudou = () => !!htela && JSON.stringify(htela.form) !== htelaOriginal;
const hPodeFechar = () => !htelaMudou()
  || confirm(htela.modo === 'novo' ? 'Descartar o que foi preenchido?' : 'Descartar as alterações deste registro?');
function hAbreTela(modo, h) {
  htela = {modo, form: modo === 'novo' ? hNovoForm() : hFormDe(h)};
  htelaOriginal = JSON.stringify(htela.form);
  hRenderTela();
  foca('hf-hospede');
}
function hFechaTela() {
  htela = null;
  const el = $('tela');
  el.hidden = true; el.innerHTML = '';
}

function hFormHtml(f, modo) {
  const nN = f.entrada && f.saida && f.saida >= f.entrada ? noites({entrada: f.entrada, saida: f.saida}) : null;
  const sug = campo => [...new Set(HOSP.map(h => (h[campo] || '').trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, 'pt-BR')).map(v => `<option value="${esc(v)}">`).join('');
  return `<div class="h-registro" data-hform>
    <section class="cartao">
      <div class="cartao-cab"><h2>Hospedagem</h2>${nN != null ? `<span class="cont">${txtNoites(nN)}</span>` : ''}</div>
      <div class="campos">
        <div class="campo"><span class="rot">Tipo</span>
          <div class="seg tipo" role="group" aria-label="Tipo">${H_TIPOS.map(t =>
            `<button type="button" class="${t === f.tipo ? 'on' : ''}" style="--cor:${H_COR[t]}" data-acao="h-tipo" data-v="${t}">${t}</button>`).join('')}</div></div>
        <label class="campo"><span class="rot">Hóspede</span>
          <input type="text" id="hf-hospede" data-hf="hospede" value="${esc(f.hospede)}" list="h-lista-hospedes"
            placeholder="Pessoa, núcleo da família ou grupo" autocomplete="off"></label>
        <div class="campo"><span class="rot">Estadia</span>
          <div class="periodo">
            <input type="date" id="hf-entrada" data-hf="entrada" value="${f.entrada}" aria-label="Entrada">
            <span class="ate">a</span>
            <input type="date" id="hf-saida" data-hf="saida" value="${f.saida}" aria-label="Saída">
          </div></div>
        <div class="h-linha">
          <label class="campo"><span class="rot">Adultos</span>
            <input type="number" min="0" step="1" id="hf-adultos" data-hf="adultos" value="${esc(f.adultos)}" inputmode="numeric"></label>
          <label class="campo"><span class="rot">Crianças</span>
            <input type="number" min="0" step="1" id="hf-criancas" data-hf="criancas" value="${esc(f.criancas)}" inputmode="numeric"></label>
        </div>
        <div class="campo"><span class="rot">Status</span>
          <div class="seg status" role="group" aria-label="Status">${H_STATUS.map(s =>
            `<button type="button" class="${s === f.status ? 'on ' : ''}${H_CLS[s]}" data-acao="h-status" data-v="${s}">${s}</button>`).join('')}</div></div>
        ${f.tipo === 'Locação' ? `<label class="campo h-valor"><span class="rot">Valor da locação (R$)</span>
          <input type="number" min="0" step="0.01" id="hf-valor" data-hf="valor" value="${esc(f.valor)}" inputmode="decimal"></label>` : ''}
        <label class="campo"><span class="rot">Observação</span>
          <textarea id="hf-obs" data-hf="observacao" rows="2">${esc(f.observacao)}</textarea></label>
      </div>
      <datalist id="h-lista-hospedes">${sug('hospede')}</datalist>
    </section>
    <div class="acoes">
      ${modo === 'edicao' ? '<button type="button" class="bt perigo" data-acao="h-excluir">Excluir registro</button>' : ''}
      <span class="msg" data-hmsg aria-live="polite"></span>
      <span class="dir">
        <button type="button" class="bt" data-acao="h-fecha-tela">Cancelar</button>
        <button type="button" class="bt prim" data-acao="h-salvar">${modo === 'novo' ? 'Registrar hospedagem' : 'Salvar alterações'}</button>
      </span>
    </div>
  </div>`;
}

function hRenderTela() {
  const el = $('tela');
  if (!htela) { el.hidden = true; el.innerHTML = ''; return; }
  el.hidden = false;
  el.innerHTML = `<div class="tela-cab">
      <div><div class="tela-kicker">Hospedagens</div>
        <h2 id="tela-titulo">${htela.modo === 'novo' ? 'Registrar hospedagem' : 'Editar registro'}</h2></div>
      <button type="button" class="x grande" data-acao="h-fecha-tela" title="Fechar" aria-label="Fechar">×</button>
    </div>
    <div class="tela-corpo">${hFormHtml(htela.form, htela.modo)}</div>`;
}
function hRedesenhaForm() {
  const box = document.querySelector('[data-hform]');
  if (!box) return;
  const foco = document.activeElement && document.activeElement.id;
  box.outerHTML = hFormHtml(htela.form, htela.modo);
  if (foco) foca(foco);
}

async function hSalvar() {
  const f = htela.form, box = document.querySelector('[data-hform]'), c = cliente();
  const dizer = (txt, erro) => { const m = box.querySelector('[data-hmsg]'); m.textContent = txt; m.classList.toggle('erro', !!erro); };
  const hospede = f.hospede.trim();
  if (!hospede) { foca('hf-hospede'); return dizer('Informe o hóspede.', true); }
  if (!f.entrada || !f.saida) return dizer('Informe a entrada e a saída.', true);
  if (f.saida < f.entrada) return dizer('A saída não pode ser antes da entrada.', true);
  const inteiro = v => v === '' ? 0 : Number(v);
  const adultos = inteiro(f.adultos), criancas = inteiro(f.criancas);
  if (![adultos, criancas].every(v => Number.isInteger(v) && v >= 0)) return dizer('Adultos e crianças: números inteiros.', true);
  const valor = f.tipo === 'Locação' && f.valor !== '' ? Number(f.valor) : null;
  if (valor != null && !(valor >= 0)) return dizer('Valor da locação inválido.', true);
  const btn = box.querySelector('[data-acao="h-salvar"]');
  btn.disabled = true; dizer('Salvando…');
  try {
    const linha = {tipo: f.tipo, hospede, entrada: f.entrada, saida: f.saida, adultos, criancas,
                   status: f.status, valor,
                   observacao: f.observacao.trim() || null};
    let id = f.id;
    if (id) {
      const r = await c.from('sgpg_hospedagem').update(linha).eq('id', id);
      if (r.error) throw r.error;
    } else {
      const r = await c.from('sgpg_hospedagem').insert(linha).select('id').single();
      if (r.error) throw r.error;
      id = r.data.id;
    }
    await hCarrega();
    hAviso = {id, texto: htela.modo === 'novo' ? 'Hospedagem registrada.' : 'Alterações salvas.'};
    // a ficha da hospedagem gravada fica aberta, na lista do ano dela
    hVisao = 'lista'; lembra('h.visao', hVisao); hSel = id; hFiltro.busca = '';
    if (hFiltro.ano && hFiltro.ano !== f.entrada.slice(0, 4)) hFiltro.ano = f.entrada.slice(0, 4);
    hFechaTela();
    render();
  } catch (err) {
    btn.disabled = false;
    dizer('Não foi possível salvar: ' + (err.message || err), true);
  }
}

async function hExcluir() {
  const h = HOSP.find(x => x.id === htela.form.id);
  if (!confirm(`Excluir a hospedagem de "${h.hospede}" (${hPeriodo(h)})? Não dá para desfazer.`)) return;
  const r = await cliente().from('sgpg_hospedagem').delete().eq('id', h.id);
  if (r.error) {
    const m = document.querySelector('[data-hform] [data-hmsg]');
    m.textContent = 'Não foi possível excluir: ' + r.error.message; m.classList.add('erro');
    return;
  }
  hSel = null; hAviso = null;
  await hCarrega();
  hFechaTela();
  render();
}

/* ---- eventos do módulo (só as ações "h-"; o resto é do sgpg.js) ---- */
document.addEventListener('click', ev => {
  const t = ev.target.closest('[data-acao^="h-"]');
  if (!t) return;
  const f = htela && htela.form;
  switch (t.dataset.acao) {
    case 'h-visao':
      if (t.dataset.v === hVisao) return;
      hVisao = t.dataset.v; lembra('h.visao', hVisao); render(); break;
    case 'h-ano': hFiltro.ano = t.dataset.v; render(); break;
    case 'h-mes': hMes = +t.dataset.v ? movMes(hMes, +t.dataset.v) : hojeIso().slice(0, 7); render(); break;
    case 'h-limpa-busca': hFiltro.busca = ''; render(); foca('h-busca'); break;
    case 'h-sel': hSel = +t.dataset.id; hAviso = null; hRenderCorpo(); rolaParaDetalhe(); break;
    case 'h-abre': {
      // da agenda para a ficha: lista do ano da estadia
      const h = HOSP.find(x => x.id === +t.dataset.id);
      hVisao = 'lista'; lembra('h.visao', hVisao); hSel = h.id; hFiltro.busca = ''; hAviso = null;
      if (hFiltro.ano && hFiltro.ano !== h.entrada.slice(0, 4)) hFiltro.ano = h.entrada.slice(0, 4);
      render(); rolaParaDetalhe(); break;
    }
    case 'h-nova': hAbreTela('novo'); break;
    case 'h-editar': hAbreTela('edicao', HOSP.find(x => x.id === +t.dataset.id)); break;
    case 'h-fecha-tela': if (hPodeFechar()) hFechaTela(); break;
    case 'h-tipo':
      f.tipo = t.dataset.v;
      if (f.tipo !== 'Locação') f.valor = '';
      hRedesenhaForm(); break;
    case 'h-status': f.status = t.dataset.v; hRedesenhaForm(); break;
    case 'h-salvar': hSalvar(); break;
    case 'h-excluir': hExcluir(); break;
  }
});

document.addEventListener('input', ev => {
  const t = ev.target, box = t.closest && t.closest('[data-hform]');
  if (box && htela && t.dataset.hf) {
    const f = htela.form;
    f[t.dataset.hf] = t.value;
    // saída antes da entrada não existe: acompanha a entrada (uma noite)
    if (t.dataset.hf === 'entrada' && f.saida < f.entrada) {
      f.saida = somaDias(f.entrada, 1);
      const s = box.querySelector('[data-hf="saida"]');
      if (s) s.value = f.saida;
    }
    if (t.dataset.hf === 'entrada' || t.dataset.hf === 'saida') {
      const c = box.querySelector('.cartao-cab .cont'), ok = f.entrada && f.saida && f.saida >= f.entrada;
      if (c) c.textContent = ok ? txtNoites(noites({entrada: f.entrada, saida: f.saida})) : '';
    }
    return;
  }
  if (t.id === 'h-busca') {
    hFiltro.busca = t.value;
    const x = document.querySelector('[data-acao="h-limpa-busca"]');
    if (x) x.hidden = !t.value;
    hRenderCorpo();
  }
});

document.addEventListener('keydown', ev => {
  if (ev.key === 'Escape' && htela && hPodeFechar()) hFechaTela();
});
window.addEventListener('beforeunload', ev => {
  if (htelaMudou()) { ev.preventDefault(); ev.returnValue = ''; }
});

/* carrega junto com o resto do SGPG; se o módulo já estiver na tela, redesenha */
hCarrega().then(() => { if (modulo === 'hospedagens') render(); })
  .catch(err => { hErro = err.message || String(err); hCarregado = true; if (modulo === 'hospedagens') render(); });
