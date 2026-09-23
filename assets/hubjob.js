/* Barra "Atualizar dados" — compartilhada pela Atualização Semanal (dentro do
   app.js) e pelo Comitê (dentro do deck, que roda num iframe).

   O hub é site estático e o pipeline lê o Google Drive montado em G:, que só
   existe na máquina de quem fecha. Então clicar aqui NÃO executa nada: grava um
   pedido em `hub_job` e fica acompanhando. Quem executa é tools/agente_hub.py,
   agendado nessa máquina — por isso o estado "na fila" pode durar alguns
   minutos, e é ele que a barra mostra em vez de fingir progresso.

   Por que um arquivo só: no comitê a barra vive dentro do iframe (junto dos
   outros botões do deck) e no semanal vive na casca. Duplicar daria duas
   versões da mesma regra de estado. */
'use strict';

(function () {
  const POLL_MS = 5000;
  // Rótulo de cada estado, na voz de quem está esperando — não é o valor cru do
  // banco. 'fila' é o estado que mais confunde: o agente pode levar minutos.
  const ESTADO = {
    fila:    ['na fila — o agente pega em alguns minutos', 'espera'],
    rodando: ['atualizando…', 'espera'],
    ok:      ['atualizado', 'ok'],
    erro:    ['falhou', 'erro'],
  };

  // O deck roda em iframe; a sessão do Supabase vive na casca do hub.
  function sb() {
    if (window.HUB && window.HUB.sb) return window.HUB.sb;
    try { return window.parent.HUB && window.parent.HUB.sb; } catch (e) { return null; }
  }

  const hhmm = t => {
    try { return new Date(t).toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'}); }
    catch (e) { return ''; }
  };

  async function ultimo(tipo) {
    const c = sb();
    if (!c) return null;
    const { data } = await c.from('hub_job').select('*').eq('tipo', tipo)
      .order('pedido_em', {ascending: false}).limit(1);
    return (data && data[0]) || null;
  }

  async function pedir(tipo) {
    const c = sb();
    if (!c) throw new Error('sem sessão do hub');
    // status/pedido_por são postos por trigger — mandar daqui não adiantaria
    const { error } = await c.from('hub_job').insert({tipo});
    if (error) throw error;
  }

  /* Monta a barra em `el`. Devolve um objeto com destroy(), pra quem troca de
     aba não deixar timer rodando em painel que saiu da tela. */
  function barra(el, tipo, opcoes) {
    const o = opcoes || {};
    const raiz = document.createElement('div');
    raiz.className = 'hj-barra';
    raiz.innerHTML = `
      <button type="button" class="hj-btn">Atualizar dados</button>
      <span class="hj-estado"></span>
      <button type="button" class="hj-log" hidden>ver o que aconteceu</button>
      ${o.extra || ''}`;
    el.appendChild(raiz);

    const btn = raiz.querySelector('.hj-btn');
    const est = raiz.querySelector('.hj-estado');
    const verLog = raiz.querySelector('.hj-log');
    let timer = null, job = null;

    const pinta = () => {
      if (!job) { est.textContent = ''; est.className = 'hj-estado'; verLog.hidden = true; btn.disabled = false; return; }
      const [txt, cls] = ESTADO[job.status] || ['', ''];
      const quando = job.terminado_em || job.pedido_em;
      est.textContent = txt + (job.status === 'ok' && quando ? ` às ${hhmm(quando)}` : '');
      est.className = 'hj-estado ' + cls;
      verLog.hidden = !(job.log && job.status === 'erro');
      btn.disabled = job.status === 'fila' || job.status === 'rodando';
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

    btn.onclick = async () => {
      btn.disabled = true;
      est.textContent = 'pedindo…'; est.className = 'hj-estado espera';
      try { await pedir(tipo); } catch (e) {
        est.textContent = 'não deu pra pedir: ' + (e.message || e);
        est.className = 'hj-estado erro';
        btn.disabled = false;
        return;
      }
      olha();
    };
    verLog.onclick = () => { if (job && job.log) alert(job.log.slice(-4000)); };

    olha();
    return {destroy: () => { if (timer) clearInterval(timer); raiz.remove(); }, raiz, recarrega: olha};
  }

  window.HubJob = {barra, ultimo, pedir};
})();
