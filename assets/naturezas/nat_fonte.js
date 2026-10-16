/* De onde o cadastro de naturezas lê e onde ele grava. Abre de dentro da aba Gastos
 * ("Editar cadastro de naturezas"), em iframe do hub (mesma origem): o
 * cliente e a sessão do Supabase vêm de window.parent.HUB, como no SGPG.
 *
 * DOIS PAINÉIS usam esta tela, e a URL diz qual: `naturezas.html` é o de Gastos do Haras
 * (o que está descrito abaixo) e `naturezas.html?painel=casa` é o de Gastos da Casa —
 * documento `naturezas_casa`, pendentes em `gastos_casa.pendentes.json`, sem responsável,
 * e os MESMOS editores (gastos_editores). Ver PAINEIS.
 *
 *   cadastro    tabela `gastos_cadastro`, linha id 'naturezas' — o que a tela edita. É
 *               também o que a rotina `publicar_gastos_haras` do Function App lê
 *               (service_role) para montar o painel de Gastos.
 *
 *   pendentes   bucket privado `hpg-data`, arquivo `gastos.pendentes.json`, publicado pela
 *               mesma rotina: os códigos do Bimer que apareceram nos relatórios e que o
 *               cadastro não conhece. Forma: { gerado_em, pendentes: [{bimer, nome, qtd,
 *               valor, meses}] }. O prefixo antes do primeiro ponto é 'gastos', que é o
 *               que a policy do bucket confere (hub_can(split_part(name,'.',1))).
 *
 *   editor      admin, ou linha em `gastos_editores`. Só esconde os botões: quem decide é a
 *               policy do banco (hub_gastos_editor).
 *
 * A CONSEQUÊNCIA, que a tela precisa dizer: gravar aqui NÃO atualiza o painel na hora. O
 * painel é publicado pela rotina (dias úteis de manhã, ou "Atualizar dados" na aba), e é
 * nele que a edição passa a valer.
 */
window.NAT_FONTE = (function () {
  'use strict';

  /* O que muda de um painel para o outro. Sem `?painel=` (ou com um desconhecido) é o
   * Haras, exatamente como era antes de a Casa existir. */
  var PAINEIS = {
    haras: {
      id: 'naturezas', arquivo: 'gastos.pendentes.json', migration: 'gastos_naturezas',
      painel: 'painel de Gastos', quando: 'dias úteis, 08:30', comResponsavel: true,
      exemploGrupo: 'Custos e Despesas > Custos > Sanidade', exemploCod: '10.4.2.03.21'
    },
    casa: {
      id: 'naturezas_casa', arquivo: 'gastos_casa.pendentes.json', migration: 'gastos_casa_cadastro',
      painel: 'painel de Gastos da Casa', quando: 'dias úteis, 08:45', comResponsavel: false,
      exemploGrupo: 'Despesas - FPG > Manutenção', exemploCod: '10.2.2.04.04'
    }
  };
  var PAINEL = (function () {
    var p = null;
    try { p = new URLSearchParams(location.search).get('painel'); } catch (e) { /* sem URL */ }
    return PAINEIS[p] || PAINEIS.haras;
  })();

  var ID = PAINEL.id;
  var TABELA = 'gastos_cadastro';
  var ARQUIVO = PAINEL.arquivo;
  var BUCKET = 'hpg-data';

  /* `updated_at` da linha quando ESTA tela carregou. É o que impede a gravação de apagar
   * em silêncio o que outra pessoa gravou no intervalo: o UPDATE só casa se a linha ainda
   * está na versão que a tela leu. O cadastro sobe INTEIRO a cada gravação, então sem esta
   * guarda a segunda pessoa a salvar desfaz o trabalho da primeira sem erro nenhum. */
  var _versao = null;

  function hub() {
    try { return window.parent.HUB || null; } catch (e) { return null; }
  }

  function cliente() {
    var h = hub();
    if (!h || !h.sb) throw new Error('sessão do hub indisponível — abra o cadastro pela aba Gastos do hub');
    return h.sb;
  }

  async function ehEditor(sb) {
    var h = hub();
    if (h && h.role === 'admin') return true;
    if (!h || !h.email) return false;
    var r = await sb.from('gastos_editores').select('email').eq('email', h.email).maybeSingle();
    return !!r.data;
  }

  async function lerEstado(sb) {
    var r = await sb.from(TABELA).select('data, updated_at').eq('id', ID).maybeSingle();
    if (r.error) throw new Error(TABELA + ': ' + r.error.message);
    if (!r.data) throw new Error('o cadastro ainda não foi criado no banco (rode a migration ' +
                                 PAINEL.migration + ')');
    return { doc: r.data.data || {}, versao: r.data.updated_at };
  }

  async function carregar() {
    var sb = cliente();

    /* O bucket pode não ter o arquivo ainda (a rotina não publicou). Não é erro: a tela
     * abre sem a lista de pendentes e diz por quê. */
    var pendentes = null;
    try {
      var dl = await sb.storage.from(BUCKET).download(ARQUIVO);
      if (dl.data) pendentes = JSON.parse(await dl.data.text());
    } catch (e) {
      console.warn('[naturezas] pendentes ainda não publicados:', e.message);
    }

    var lido = await lerEstado(sb);
    _versao = lido.versao;
    return {
      doc: NAT_DADOS.normalizar(lido.doc),
      pendentes: pendentes,
      editor: await ehEditor(sb),
      painel: PAINEL
    };
  }

  /* Campos que contam como "conteúdo". `atualizado_*` são consequência da mudança. */
  var CAMPOS_NAT = ['ordem', 'cod', 'alt', 'nome', 'resp', 'bimer', 'path'];
  var CAMPOS_IGN = ['nome'];

  function iguais(a, b, campos) {
    return campos.every(function (c) {
      return JSON.stringify(a[c] === undefined ? null : a[c]) ===
             JSON.stringify(b[c] === undefined ? null : b[c]);
    });
  }

  /* Carimba só o que MUDOU: `atualizado_por` existe para responder "quem mexeu nesta
   * natureza e quando", e carimbar tudo a cada gravação o tornaria inútil. */
  function carimbar(novos, antigos, chave, campos, agora, email) {
    var idx = {};
    antigos.forEach(function (x) { idx[chave(x)] = x; });
    return novos.map(function (x) {
      var velho = idx[chave(x)];
      if (velho && iguais(velho, x, campos)) return velho;
      return Object.assign({}, x, { atualizado_em: agora, atualizado_por: email });
    });
  }

  /* Grava o cadastro inteiro. Lê antes porque o documento é guardado por inteiro: escrever
   * só `naturezas` apagaria qualquer outra chave que ele venha a ter. */
  async function salvar(cadastro) {
    var sb = cliente();
    var lido = await lerEstado(sb);
    var estado = NAT_DADOS.normalizar(lido.doc);
    var h = hub();
    var email = (h && h.email) || null;
    var agora = new Date().toISOString();

    var naturezas = carimbar(cadastro.naturezas, estado.naturezas,
      function (n) { return n.cod; }, CAMPOS_NAT, agora, email);
    var ignoradas = carimbar(cadastro.ignoradas, estado.ignoradas,
      function (g) { return String(g.bimer); }, CAMPOS_IGN, agora, email);

    var novo = Object.assign({}, lido.doc, {
      versao: 1, naturezas: naturezas, ignoradas: ignoradas
    });

    /* `_versao` é a do carregamento da tela, não a da releitura logo acima: é a releitura
     * que enxergaria a gravação do outro e faria a guarda passar sempre. */
    var esperada = _versao || lido.versao;
    var r = await sb.from(TABELA)
      .update({ data: novo })
      .eq('id', ID)
      .eq('updated_at', esperada)
      .select('id, updated_at');
    if (r.error) throw new Error('não foi possível gravar: ' + r.error.message);
    if (!r.data || !r.data.length) {
      /* Zero linhas casadas tem DOIS motivos: alguém gravou no meio (a versão mudou) ou a
       * policy recusou (não é editor) — o PostgREST responde igual nos dois. O texto não
       * afirma um só: quem não é editor não chega aqui, a tela já está em consulta. */
      throw new Error('outra pessoa gravou o cadastro enquanto esta tela estava aberta (ou ' +
                      'você perdeu a permissão de editar). Nada foi sobrescrito — recarregue ' +
                      'a aba e refaça a alteração.');
    }
    /* A versão vem da resposta: quem decide o valor gravado é o banco (trigger). */
    _versao = r.data[0].updated_at;
    return { naturezas: naturezas, ignoradas: ignoradas };
  }

  return { carregar: carregar, salvar: salvar, painel: PAINEL };
})();
