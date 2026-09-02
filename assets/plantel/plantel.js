/* Aba Plantel / Movimentação — mini-sistema do fechamento mensal.

   O artefato (assets/plantel/spec.js, gerado por tools/build_plantel_mov.py) é a
   ÚNICA fonte de número desta tela. Nada é calculado aqui: o que o build não
   soube explicar aparece na FILA para uma pessoa decidir, e a decisão vai pro
   Supabase — nunca some num JSON local nem fica só na cabeça de quem fechou.

   Cinco painéis:
     Resumo   — cascata mensal no formato do Resumo Contábil que era feito à mão
     Fila     — o que precisa de decisão humana: efeito sem causa e log sem efeito
     Registro — todo movimento do ano com sua abertura (antes/depois + log)
     Abertura — caminhada anual por animal: valor inicial + causas = valor final
     Checks   — nossa base contra os totais da planilha do haras

   O mês ABERTO é o único classificável. Mês fechado é histórico: aparece inteiro,
   mas sem botão — fechado é fechado. */
'use strict';

const SPEC = window.PLANTEL_MOV;
const ROT = {
  compras: '(+) Compras', producao_embrioes: '(+) Produção / embriões',
  baixa_vendas: '(−) Baixa vendas', baixa_mortes_doacoes: '(−) Baixa mortes e doações',
  reavaliacoes: '(+/−) Reavaliações', sem_classificacao: '(?) Sem classificação',
};
const CLASSES = ['venda', 'morte', 'doacao', 'compra', 'producao', 'transferencia', 'manter'];
const MESES_PT = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const rs = v => (v == null || isNaN(v)) ? '—'
  : (v < 0 ? '−' : '') + 'R$ ' + Math.abs(v).toLocaleString('pt-BR', {maximumFractionDigits: 0});
const cls = v => !v ? 'zero' : v > 0 ? 'pos' : 'neg';
const rotMes = m => { const [a, mm] = m.split('-'); return `${MESES_PT[+mm - 1]}/${a.slice(2)}`; };
const pct = v => v == null ? '—' : (v * 100).toLocaleString('pt-BR', {maximumFractionDigits: 2}) + '%';

/* ---- sessão do hub (mesma origem: o hub monta window.HUB antes do iframe) ---- */
function hub(){ try { return window.parent.HUB || null; } catch (e) { return null; } }
function sb(){ const h = hub(); return h && h.sb; }
function euSou(){ const h = hub(); return (h && h.email) || null; }

/* decisões já tomadas, por chave "mes|k" */
let decisoes = {};

async function carregaDecisoes(){
  const c = sb();
  if (!c) return;
  try {
    const { data, error } = await c.from('plantel_mov_classificacao')
      .select('mes,chave,classe,nota,autor,criado_em').eq('mes', SPEC.mes_aberto);
    if (error) throw error;
    (data || []).forEach(r => { decisoes[`${r.mes}|${r.chave}`] = r; });
  } catch (e) {
    console.warn('[plantel] sem classificações salvas', e.message || e);
  }
}

async function salvaDecisao(mes, chave, classe, nota, nome){
  const c = sb();
  if (!c) { alert('Sem sessão do hub — a decisão não pode ser gravada.'); return false; }
  const linha = {mes, chave, classe, nota: nota || null, nome: nome || null,
                 autor: euSou(), criado_em: new Date().toISOString()};
  const { error } = await c.from('plantel_mov_classificacao')
    .upsert(linha, {onConflict: 'mes,chave'});
  if (error) { alert('Falha ao gravar: ' + error.message); return false; }
  decisoes[`${mes}|${chave}`] = linha;
  return true;
}

/* ================= Resumo ================= */
function painelResumo(){
  const meses = SPEC.meses.filter(m => SPEC.resumo[m]);
  const r = SPEC.resumo[SPEC.mes_aberto] || {};
  const erroMax = Math.max(...meses.map(m => Math.abs(SPEC.resumo[m].erro_identidade)));
  const kpis = [
    {v: rs(r.saldo_fim), l: `saldo em ${rotMes(SPEC.mes_aberto)}`,
     s: `${r.cabecas || 0} cabeças no escopo`},
    {v: rs(r.saldo_fim - r.saldo_ini), l: 'variação do mês',
     s: `partiu de ${rs(r.saldo_ini)}`},
    {v: rs(erroMax), l: 'erro de identidade', cls: erroMax < 0.01 ? 'ok' : 'alerta',
     s: 'Σ causas − Δ saldo, pior mês do ano'},
    {v: String(SPEC.pendencias.length + SPEC.divergencias.length), l: 'na fila',
     cls: (SPEC.pendencias.length + SPEC.divergencias.length) ? 'alerta' : 'ok',
     s: 'itens sem explicação automática'},
  ];
  const linhas = Object.keys(ROT);
  return `
    <div class="nota">Escopo contábil: sufixo <b>${SPEC.escopo.join('</b> e <b>')}</b> — o plantel da
      Carla. Os <b>DA PAO GRANDE - E xx%</b> são parceria e ficam fora, que é o corte com que este
      saldo reproduz o Resumo Contábil. A comissão de 8,5% de alguns animais fica em coluna própria
      (${rs(r.comissao_fim)} no mês), e é por ela que o mapa do haras fecha um pouco acima.</div>
    <div class="grade">${kpis.map(k => `<div class="kpi ${k.cls || ''}">
      <div class="v">${k.v}</div><div class="l">${esc(k.l)}</div><div class="s">${esc(k.s)}</div></div>`).join('')}</div>
    <div class="rolagem"><table class="t">
      <thead><tr><th class="l">Título</th>${meses.map(m => `<th>${rotMes(m)}${SPEC.resumo[m].fechado ? '' : ' •'}</th>`).join('')}<th>Ano</th></tr></thead>
      <tbody>
        <tr class="tot"><td>Saldo inicial</td>${meses.map(m => `<td>${rs(SPEC.resumo[m].saldo_ini)}</td>`).join('')}
          <td>${rs(SPEC.resumo[meses[0]].saldo_ini)}</td></tr>
        ${linhas.map(rot => {
          const vals = meses.map(m => SPEC.resumo[m].linhas[rot] || 0);
          const soma = vals.reduce((a, b) => a + b, 0);
          if (!soma && !vals.some(v => v)) return '';
          return `<tr><td>${ROT[rot]}</td>${vals.map(v => `<td class="${cls(v)}">${v ? rs(v) : '—'}</td>`).join('')}
            <td class="${cls(soma)}">${rs(soma)}</td></tr>`;
        }).join('')}
        <tr class="tot"><td>Saldo final</td>${meses.map(m => `<td>${rs(SPEC.resumo[m].saldo_fim)}</td>`).join('')}
          <td>${rs(SPEC.resumo[meses[meses.length - 1]].saldo_fim)}</td></tr>
        <tr><td>Identidade (Σ causas − Δ saldo)</td>${meses.map(m => {
          const e = SPEC.resumo[m].erro_identidade;
          return `<td class="${Math.abs(e) < 0.01 ? 'pos' : 'neg'}">${Math.abs(e) < 0.01 ? 'ok' : rs(e)}</td>`;
        }).join('')}<td></td></tr>
        <tr><td>Cabeças</td>${meses.map(m => `<td>${SPEC.resumo[m].cabecas}</td>`).join('')}<td></td></tr>
      </tbody></table></div>
    <p style="color:var(--ink-3);font-size:11.5px;margin-top:10px">• mês aberto — ainda pode mudar.</p>`;
}

/* ================= Fila ================= */
function cartaoPendencia(p, i){
  const chave = `${p.mes}|${p.k}`;
  const d = decisoes[chave];
  const logs = (p.log || []);
  return `
    <div class="item ${d ? 'resolvido' : ''}" data-item="${i}">
      <span class="valor ${cls(p.efeito)}">${rs(p.efeito)}</span>
      <h3>${esc(p.nome)}</h3>
      <div class="meta">${esc(p.categoria || '—')} · ${esc(p.sufixo)} · efeito sem causa identificada
        ${d ? ` · <b style="color:var(--amber)">classificado como ${esc(d.classe)}</b>${d.autor ? ' por ' + esc(d.autor) : ''}` : ''}</div>
      <div class="antes-depois">
        <div><div class="rot">Cota</div><div class="val">${pct(p.cotas_ant)} → ${pct(p.cotas_atual)}</div></div>
        <div><div class="rot">Valor 100%</div><div class="val">${rs(p.valor_ant)} → ${rs(p.valor_atual)}</div></div>
        <div><div class="rot">Status</div><div class="val">${esc(p.status_ant)} → ${esc(p.status_atual)}</div></div>
      </div>
      ${logs.length
        ? `<ul class="logs">${logs.map(l => `<li><b>${esc(l.data)}</b> · <span class="tag t${l.tipo[0]}">${esc(l.tipo)}</span>
            ${esc(l.ocorrencia)}</li>`).join('')}</ul>`
        : `<div class="vazio-log">Nada no log do haras para este animal neste mês — a decisão é toda humana.</div>`}
      <div class="acoes" data-chave="${esc(chave)}" data-nome="${esc(p.nome)}">
        ${CLASSES.map(c => `<button type="button" data-classe="${c}" class="${d && d.classe === c ? 'on' : ''}">${c}</button>`).join('')}
        <input placeholder="por que (opcional)" value="${esc(d && d.nota || '')}">
      </div>
    </div>`;
}

function cartaoDivergencia(x, i){
  const chave = `${SPEC.mes_aberto}|LOG:${x.produto}|${x.data}`;
  const d = decisoes[chave];
  const e = x.estado_no_plantel;
  return `
    <div class="item ${d ? 'resolvido' : ''}">
      <h3>${esc(x.produto)}</h3>
      <div class="meta"><span class="tag t3">3-financeira / ${esc(x.subtipo)}</span> registrada em
        ${esc(x.data)}, sem efeito no patrimônio
        ${d ? ` · <b style="color:var(--amber)">${esc(d.classe)}</b>${d.autor ? ' por ' + esc(d.autor) : ''}` : ''}</div>
      <ul class="logs"><li>${esc(x.ocorrencia)}</li></ul>
      ${e ? `<div class="antes-depois">
          <div><div class="rot">Cota hoje</div><div class="val">${pct(e.cota)}</div></div>
          <div><div class="rot">Patrimônio</div><div class="val">${rs(e.patrimonio)}</div></div>
          <div><div class="rot">Status</div><div class="val">${esc(e.status)}</div></div>
          <div><div class="rot">Sufixo</div><div class="val">${esc(e.sufixo)}</div></div>
        </div>`
        : `<div class="vazio-log">Animal não encontrado no plantel deste mês.</div>`}
      <div class="acoes" data-chave="${esc(chave)}" data-nome="${esc(x.produto)}">
        ${['baixar', 'manter', 'sem_efeito', 'corrigir_planilha'].map(c =>
          `<button type="button" data-classe="${c}" class="${d && d.classe === c ? 'on' : ''}">${c}</button>`).join('')}
        <input placeholder="por que (opcional)" value="${esc(d && d.nota || '')}">
      </div>
    </div>`;
}

function painelFila(){
  const semSessao = !sb();
  const p = SPEC.pendencias, dv = SPEC.divergencias;
  if (!p.length && !dv.length) {
    return `<div class="limpo">Nada na fila: todo efeito do mês tem causa identificada, e todo
      lançamento financeiro do log tem efeito no patrimônio.</div>`;
  }
  return `
    ${semSessao ? `<div class="nota"><b>Somente leitura.</b> Sem sessão do hub, a classificação não
      pode ser gravada — abra pelo hub autenticado para decidir.</div>` : ''}
    <div class="nota">Mês aberto: <b>${rotMes(SPEC.mes_aberto)}</b>. Mês fechado não entra na fila —
      fechado é fechado. Duas naturezas aqui: <b>efeito sem causa</b> (o patrimônio mudou e o log não
      explica) e <b>log sem efeito</b> (o haras registrou a movimentação e o plantel não mexeu).</div>
    ${p.length ? `<h2 style="font-size:14px;margin:18px 0 10px">Efeito sem causa (${p.length})</h2>
      ${p.map(cartaoPendencia).join('')}` : ''}
    ${dv.length ? `<h2 style="font-size:14px;margin:22px 0 10px">Log sem efeito no patrimônio (${dv.length})</h2>
      ${dv.map(cartaoDivergencia).join('')}` : ''}`;
}

/* ================= Registro ================= */
function painelRegistro(){
  const tipos = [...new Set(SPEC.efeitos.map(e => e.tipo))].sort();
  return `
    <div class="filtros">
      <select id="fMes"><option value="">todos os meses</option>
        ${SPEC.meses.map(m => `<option value="${m}"${m === SPEC.mes_aberto ? ' selected' : ''}>${rotMes(m)}</option>`).join('')}</select>
      <select id="fTipo"><option value="">todos os tipos</option>
        ${tipos.map(t => `<option value="${t}">${t}</option>`).join('')}</select>
      <select id="fEscopo"><option value="1">só escopo Carla</option><option value="">tudo</option></select>
      <input id="fTexto" placeholder="filtrar por nome do animal">
    </div>
    <div class="rolagem"><table class="t" id="tabRegistro"></table></div>`;
}

function desenhaRegistro(){
  const mes = document.getElementById('fMes').value;
  const tipo = document.getElementById('fTipo').value;
  const soEscopo = document.getElementById('fEscopo').value === '1';
  const txt = document.getElementById('fTexto').value.trim().toUpperCase();
  const linhas = SPEC.efeitos.filter(e =>
    (!mes || e.mes === mes) && (!tipo || e.tipo === tipo) && (!soEscopo || e.no_escopo) &&
    (!txt || (e.nome || '').toUpperCase().includes(txt)));
  const soma = linhas.reduce((a, e) => a + e.efeito, 0);
  document.getElementById('tabRegistro').innerHTML = `
    <thead><tr><th class="l">Mês</th><th class="l">Animal</th><th class="l">Tipo</th>
      <th>Cota</th><th>Valor 100%</th><th class="l">Status</th><th>Efeito</th><th class="l">Log do haras</th></tr></thead>
    <tbody>${linhas.map(e => `<tr>
      <td class="l">${rotMes(e.mes)}</td>
      <td class="l nome" title="${esc(e.nome)}">${esc(e.nome)}</td>
      <td class="l">${esc(e.tipo)}</td>
      <td>${pct(e.cotas_ant)} → ${pct(e.cotas_atual)}</td>
      <td>${rs(e.valor_ant)} → ${rs(e.valor_atual)}</td>
      <td class="l">${esc(e.status_ant)} → ${esc(e.status_atual)}</td>
      <td class="${cls(e.efeito)}">${rs(e.efeito)}</td>
      <td class="l nome">${(e.log || []).map(l => `<span class="tag t${l.tipo[0]}">${esc(l.subtipo || l.tipo)}</span>`).join(' ') || '<span class="zero">—</span>'}</td>
    </tr>`).join('')}
    <tr class="tot"><td colspan="6">${linhas.length} movimento(s)</td>
      <td class="${cls(soma)}">${rs(soma)}</td><td></td></tr></tbody>`;
}

/* ================= Abertura ================= */
function painelAbertura(){
  const cols = Object.keys(ROT);
  const ab = SPEC.aberturas.slice().sort((a, b) => b.valor_fim - a.valor_fim);
  const t = k => ab.reduce((a, x) => a + (x[k] || 0), 0);
  return `
    <div class="nota">Caminhada do ano por animal, no mesmo formato do mapa que o haras mantém:
      <b>valor inicial + causas = valor final</b>. A coluna <b>fecha?</b> marca a linha em que a conta
      não fecha — é dado a conferir, não número a usar.</div>
    <div class="filtros"><input id="fAb" placeholder="filtrar por nome do animal"></div>
    <div class="rolagem"><table class="t" id="tabAb">
      <thead><tr><th class="l">Animal</th><th class="l">Categoria</th><th class="l">Status</th><th>Cota</th>
        <th>Valor inicial</th>${cols.map(c => `<th>${ROT[c].replace(/^\([^)]*\)\s*/, '')}</th>`).join('')}
        <th>Valor final</th><th>fecha?</th></tr></thead>
      <tbody>${ab.map(a => `<tr data-nome="${esc((a.nome || '').toUpperCase())}">
        <td class="l nome" title="${esc(a.nome)}">${esc(a.nome)}</td>
        <td class="l">${esc(a.categoria || '—')}</td><td class="l">${esc(a.status || '—')}</td>
        <td>${pct(a.cota)}</td><td>${rs(a.valor_ini)}</td>
        ${cols.map(c => `<td class="${cls(a[c])}">${a[c] ? rs(a[c]) : '—'}</td>`).join('')}
        <td>${rs(a.valor_fim)}</td>
        <td class="${Math.abs(a.erro) < 0.01 ? 'pos' : 'neg'}">${Math.abs(a.erro) < 0.01 ? 'ok' : rs(a.erro)}</td>
      </tr>`).join('')}
      <tr class="tot"><td colspan="4">${ab.length} animais</td><td>${rs(t('valor_ini'))}</td>
        ${cols.map(c => `<td class="${cls(t(c))}">${rs(t(c))}</td>`).join('')}
        <td>${rs(t('valor_fim'))}</td><td></td></tr></tbody></table></div>`;
}

/* ================= Checks ================= */
function painelChecks(){
  const c = SPEC.checks.slice().reverse();
  const ruins = c.filter(x => !x.ok);
  return `
    <div class="nota">Confere a <b>nossa base</b> contra os <b>totais da própria planilha do haras</b>,
      mês a mês: cotas, valor e comissão. Divergência aqui invalida tudo o que vem depois, então é o
      primeiro lugar a olhar quando um número parecer estranho.</div>
    <div class="grade">
      <div class="kpi ${ruins.length ? 'alerta' : 'ok'}"><div class="v">${c.length - ruins.length}/${c.length}</div>
        <div class="l">meses conferidos</div><div class="s">${ruins.length ? ruins.length + ' com divergência' : 'nenhuma divergência'}</div></div>
    </div>
    <div class="rolagem"><table class="t">
      <thead><tr><th class="l">Mês</th><th>Cotas base</th><th>Cotas planilha</th><th>Dif</th>
        <th>Valor base</th><th>Valor planilha</th><th>Dif</th><th>Comissão base</th><th>Comissão planilha</th>
        <th>Dif</th><th class="l">Situação</th></tr></thead>
      <tbody>${c.map(x => `<tr>
        <td class="l">${rotMes(x.mes)}</td>
        <td>${(x.cotas_base ?? 0).toFixed(2)}</td><td>${x.cotas_orig == null ? '—' : x.cotas_orig.toFixed(2)}</td>
        <td class="${cls(x.cotas_dif)}">${x.cotas_dif == null ? '—' : x.cotas_dif.toFixed(2)}</td>
        <td>${rs(x.valor_base)}</td><td>${rs(x.valor_orig)}</td><td class="${cls(x.valor_dif)}">${x.valor_dif == null ? '—' : rs(x.valor_dif)}</td>
        <td>${rs(x.comissao_base)}</td><td>${rs(x.comissao_orig)}</td><td class="${cls(x.comissao_dif)}">${x.comissao_dif == null ? '—' : rs(x.comissao_dif)}</td>
        <td class="l">${x.ok ? '<span class="tag t2">ok</span>' : '<span class="tag t3">' + esc(x.motivo || 'divergência') + '</span>'}</td>
      </tr>`).join('')}</tbody></table></div>`;
}

/* ================= casca ================= */
const PAINEIS = [
  {id: 'resumo', t: 'Resumo contábil', render: painelResumo},
  {id: 'fila', t: 'Fila', render: painelFila, cont: () => SPEC.pendencias.length + SPEC.divergencias.length},
  {id: 'registro', t: 'Registro', render: painelRegistro, depois: desenhaRegistro,
   cont: () => SPEC.efeitos.filter(e => e.mes === SPEC.mes_aberto).length},
  {id: 'abertura', t: 'Abertura por animal', render: painelAbertura, cont: () => SPEC.aberturas.length},
  {id: 'checks', t: 'Checks', render: painelChecks},
];

function mostra(id){
  for (const p of PAINEIS) {
    const el = document.getElementById('p-' + p.id);
    el.hidden = p.id !== id;
    document.getElementById('b-' + p.id).classList.toggle('on', p.id === id);
    if (p.id === id && !el.dataset.pronto) {
      el.innerHTML = p.render();
      el.dataset.pronto = '1';
      if (p.depois) p.depois();
    }
  }
  location.hash = '#' + id;
}

function boot(){
  const r = SPEC.resumo[SPEC.mes_aberto] || {};
  document.getElementById('topo').innerHTML = `
    <h1>Plantel / Movimentação</h1>
    <span class="mes">${rotMes(SPEC.mes_aberto)} <span class="tag aberto">aberto</span></span>
    <span class="tag fechado">fechado até ${rotMes(SPEC.fechado_ate)}</span>
    <span class="fonte">${esc(SPEC.fonte.controle)}<br>
      gerado em ${esc(SPEC.gerado_em.replace('T', ' '))}
      ${SPEC.fonte.revertidos_pos_mes ? ` · ${SPEC.fonte.revertidos_pos_mes} movimentação(ões) posterior(es) ao mês revertida(s)` : ''}</span>`;
  document.getElementById('abas').innerHTML = PAINEIS.map(p =>
    `<button type="button" id="b-${p.id}">${p.t}${p.cont ? `<span class="cont">${p.cont()}</span>` : ''}</button>`).join('');
  document.getElementById('paineis').innerHTML = PAINEIS.map(p =>
    `<div class="painel" id="p-${p.id}" hidden></div>`).join('');
  PAINEIS.forEach(p => { document.getElementById('b-' + p.id).onclick = () => mostra(p.id); });

  // filtros e classificação por delegação: os painéis são redesenhados inteiros
  document.getElementById('paineis').addEventListener('input', e => {
    if (['fMes', 'fTipo', 'fEscopo', 'fTexto'].includes(e.target.id)) desenhaRegistro();
    if (e.target.id === 'fAb') {
      const t = e.target.value.trim().toUpperCase();
      document.querySelectorAll('#tabAb tbody tr[data-nome]').forEach(tr => {
        tr.hidden = !!t && !tr.dataset.nome.includes(t);
      });
    }
  });
  document.getElementById('paineis').addEventListener('change', e => {
    if (['fMes', 'fTipo', 'fEscopo'].includes(e.target.id)) desenhaRegistro();
  });
  document.getElementById('paineis').addEventListener('click', async e => {
    const btn = e.target.closest('.acoes button');
    if (!btn) return;
    const caixa = btn.closest('.acoes');
    const nota = caixa.querySelector('input').value.trim();
    const ok = await salvaDecisao(SPEC.mes_aberto, caixa.dataset.chave.split('|').slice(1).join('|'),
                                  btn.dataset.classe, nota, caixa.dataset.nome);
    if (!ok) return;
    caixa.querySelectorAll('button').forEach(b => b.classList.toggle('on', b === btn));
    btn.closest('.item').classList.add('resolvido');
  });

  const inicial = (location.hash || '').replace('#', '');
  mostra(PAINEIS.some(p => p.id === inicial) ? inicial : 'resumo');
}

if (!SPEC) {
  document.getElementById('topo').innerHTML =
    `<h1>Plantel / Movimentação</h1><span class="fonte">sem artefato: rode
     <code>python tools/build_plantel_mov.py</code></span>`;
} else {
  carregaDecisoes().then(boot);
}
