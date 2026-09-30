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
    verLog.onclick = () => { if (job && job.log) alert(job.log.slice(-4000)); };

    olha();
    return {destroy: () => { if (timer) clearInterval(timer); raiz.remove(); }, raiz, recarrega: olha};
  }

  window.HubJob = {barra, ultimo, pedir};
})();
