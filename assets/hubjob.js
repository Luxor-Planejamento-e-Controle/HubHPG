/* Barra "Atualizar dados" — compartilhada pela Atualização Semanal (dentro do
   app.js) e pelo Comitê (dentro do deck, que roda num iframe).

   O hub é site estático e o pipeline lê o Google Drive montado em G:, que só
   existe na máquina de quem fecha. Então clicar aqui NÃO executa nada: grava um
   pedido em `hub_job` e fica acompanhando. Quem executa é tools/agente_hub.py,
   de pé nessa máquina e olhando a fila a cada 5 s — com ela desligada, o
   pedido fica "na fila", e é isso que a barra mostra em vez de fingir progresso.

   Por que um arquivo só: no comitê a barra vive dentro do iframe (junto dos
   outros botões do deck) e no semanal vive na casca. Duplicar daria duas
   versões da mesma regra de estado. */
'use strict';

(function () {
  const POLL_MS = 5000;
  // Rótulo de cada estado, na voz de quem está esperando — não é o valor cru do
  // banco. 'fila' é o estado que mais confunde: sem a máquina do agente ligada,
  // o pedido não anda.
  const ESTADO = {
    fila:    ['na fila — aguardando o agente', 'espera'],
    rodando: ['atualizando…', 'espera'],
    ok:      ['atualizado', 'ok'],
    erro:    ['falhou', 'erro'],
  };

  // O deck roda em iframe; a sessão do Supabase vive na casca do hub.
  function sb() {
    if (window.HUB && window.HUB.sb) return window.HUB.sb;
    try { return window.parent.HUB && window.parent.HUB.sb; } catch (e) { return null; }
  }

  // Data e hora completas: "às 09:27" sozinho não dizia de que dia — e o último
  // pedido pode ser de uma semana atrás, que é justamente o que se quer saber.
  const quandoTxt = t => {
    try {
      const d = new Date(t);
      const dia = d.toLocaleDateString('pt-BR', {day: '2-digit', month: '2-digit', year: 'numeric'});
      const hora = d.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'});
      return `em ${dia} às ${hora}`;
    } catch (e) { return ''; }
  };

  /* Janela do log. Era alert(log.slice(-4000)): o navegador corta texto longo
     com "…" e não deixa selecionar — em 09/10/2026 o motivo da falha (última
     linha, "fonte mais velha que a janela") nunca apareceu. Aqui vai o log
     inteiro, copiável, já rolado até o fim, que é onde o pipeline diz por que
     parou.

     Abre na casca do hub: no semanal a barra mora num iframe que cresce com o
     conteúdo, e um position:fixed lá dentro cairia no meio de uma página de
     milhares de px; no comitê o script roda dentro do deck. */
  const LOG_CSS = `
.hj-janela{width:min(900px,94vw);max-height:86vh;padding:0;border:1px solid var(--line,#1B486B);
  border-radius:12px;background:var(--bg-2,#072B49);color:var(--ink,#EAF0F4);
  font-family:var(--font-sans,"Segoe UI",system-ui,sans-serif)}
.hj-janela::backdrop{background:rgba(2,16,28,.78)}
.hj-janela .hj-cab{display:flex;align-items:center;gap:10px;padding:12px 16px;
  border-bottom:1px solid var(--line,#1B486B);font-size:13px}
.hj-janela .hj-cab b{margin-right:auto;font-weight:600}
.hj-janela button{background:none;border:1px solid var(--line,#1B486B);color:var(--ink-3,#93AABC);
  border-radius:7px;padding:5px 10px;font-size:12px;cursor:pointer}
.hj-janela button:hover{color:var(--ink,#EAF0F4);border-color:var(--amber,#CA9703)}
.hj-janela pre{margin:0;padding:14px 16px;max-height:calc(86vh - 56px);overflow:auto;
  font:12px/1.55 ui-monospace,Consolas,monospace;white-space:pre-wrap;word-break:break-word;
  user-select:text}
.hj-janela .ruim{color:var(--neg,#F07A7A);font-weight:600}`;
  // o que marca o motivo da parada; o resto do log é contexto
  const RUIM = /ABORTADO|FALHOU|Traceback|Error\b|!!|recusou/;

  function mostraLog(j) {
    let d;
    try { d = window.top.document; } catch (e) { d = document; }
    if (!d.getElementById('hj-log-estilo')) {
      const st = d.createElement('style');
      st.id = 'hj-log-estilo'; st.textContent = LOG_CSS;
      d.head.appendChild(st);
    }
    const dlg = d.createElement('dialog');
    dlg.className = 'hj-janela';
    const quando = j.terminado_em ? ' ' + quandoTxt(j.terminado_em) : '';
    dlg.innerHTML = `<div class="hj-cab"><b></b>
      <button type="button" data-a="copia">copiar</button>
      <button type="button" data-a="fecha">fechar</button></div><pre></pre>`;
    dlg.querySelector('b').textContent = `Pedido #${j.id} — ${(ESTADO[j.status] || [j.status])[0]}${quando}`;
    const pre = dlg.querySelector('pre');
    j.log.split('\n').forEach(l => {
      const s = d.createElement('span');
      if (RUIM.test(l)) s.className = 'ruim';
      s.textContent = l + '\n';
      pre.appendChild(s);
    });
    dlg.addEventListener('click', ev => {
      const a = ev.target.dataset && ev.target.dataset.a;
      if (a === 'fecha' || ev.target === dlg) dlg.close();
      if (a === 'copia') {
        d.defaultView.navigator.clipboard.writeText(j.log).then(
          () => { ev.target.textContent = 'copiado'; },
          () => { ev.target.textContent = 'não deu — selecione e Ctrl+C'; });
      }
    });
    dlg.addEventListener('close', () => dlg.remove());
    d.body.appendChild(dlg);
    dlg.showModal();
    pre.scrollTop = pre.scrollHeight;
  }

  async function ultimo(tipo) {
    const c = sb();
    if (!c) return null;
    const { data } = await c.from('hub_job').select('*').eq('tipo', tipo)
      .order('pedido_em', {ascending: false}).limit(1);
    return (data && data[0]) || null;
  }

  // `detalhe` escolhe entre os pedidos do mesmo painel (o comitê tem dois); o
  // agente só aceita as ações que conhece.
  async function pedir(tipo, detalhe) {
    const c = sb();
    if (!c) throw new Error('sem sessão do hub');
    // status/pedido_por são postos por trigger — mandar daqui não adiantaria
    const { error } = await c.from('hub_job').insert(detalhe ? {tipo, detalhe} : {tipo});
    if (error) throw error;
  }

  /* Monta a barra em `el`. `opcoes.acoes` troca o botão único por um por ação —
     {rotulo, detalhe, bloqueio}; `bloqueio` é o motivo de o botão estar apagado.
     As ações dividem o estado: o banco aceita um pedido ativo por tipo.
     Devolve um objeto com destroy(), pra quem troca de aba não deixar timer
     rodando em painel que saiu da tela. */
  function barra(el, tipo, opcoes) {
    const o = opcoes || {};
    const acoes = o.acoes || [{rotulo: 'Atualizar dados'}];
    const raiz = document.createElement('div');
    raiz.className = 'hj-barra';
    raiz.innerHTML = `
      <span class="hj-estado"></span>
      <button type="button" class="hj-log" hidden>ver o que aconteceu</button>
      ${o.extra || ''}`;
    const btns = acoes.map((a, i) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'hj-btn' + (i ? ' sec' : '');
      b.textContent = a.rotulo;
      return b;
    });
    raiz.prepend(...btns);
    el.appendChild(raiz);

    const est = raiz.querySelector('.hj-estado');
    const verLog = raiz.querySelector('.hj-log');
    let timer = null, job = null;

    const libera = ativo => btns.forEach((b, i) => {
      b.disabled = ativo || !!acoes[i].bloqueio;
      b.title = acoes[i].bloqueio || '';
    });

    const pinta = () => {
      if (!job) { est.textContent = ''; est.className = 'hj-estado'; verLog.hidden = true; libera(false); return; }
      const [txt, cls] = ESTADO[job.status] || ['', ''];
      const quando = job.terminado_em || job.pedido_em;
      est.textContent = txt + (job.status === 'ok' && quando ? ` ${quandoTxt(quando)}` : '');
      est.className = 'hj-estado ' + cls;
      verLog.hidden = !(job.log && job.status === 'erro');
      libera(job.status === 'fila' || job.status === 'rodando');
    };

    const olha = async () => {
      const antes = job && job.status;
      job = await ultimo(tipo);
      pinta();
      const ativo = job && (job.status === 'fila' || job.status === 'rodando');
      if (ativo && !timer) timer = setInterval(olha, POLL_MS);
      if (!ativo && timer) { clearInterval(timer); timer = null; }
      // terminou agora: quem embute avisa (o painel precisa recarregar o dado)
      if (antes && antes !== job.status && job.status === 'ok' && o.aoTerminar) o.aoTerminar(job);
    };

    btns.forEach((b, i) => b.onclick = async () => {
      libera(true);
      est.textContent = 'pedindo…'; est.className = 'hj-estado espera';
      try { await pedir(tipo, acoes[i].detalhe); } catch (e) {
        est.textContent = 'não deu pra pedir: ' + (e.message || e);
        est.className = 'hj-estado erro';
        libera(false);
        return;
      }
      olha();
    });
    verLog.onclick = () => { if (job && job.log) mostraLog(job); };

    olha();
    return {destroy: () => { if (timer) clearInterval(timer); raiz.remove(); }, raiz, recarrega: olha};
  }

  window.HubJob = {barra, ultimo, pedir};
})();
