/* Comitê Mensal HPG — renderiza o spec como deck e exporta o mesmo spec em PPTX.

   Regra do desenho: o spec (assets/comite/spec.js, gerado por
   hub/tools/build_comite.py) é a ÚNICA fonte. HTML e PPTX são duas saídas do
   mesmo objeto — se um slide muda, muda nos dois, porque nenhum dos dois guarda
   conteúdo próprio. Slide sem fonte de dado vira `pendente` e diz isso na tela;
   número não se inventa pra encher slide.

   O spec traz um deck por MÊS (`SPEC.decks`), e o seletor da barra troca de mês
   sem recarregar a página. */
'use strict';

const SPEC = window.COMITE_SPEC;
/* O desenho de cada slide mora em layout.js: ele devolve a lista de primitivas
   (retângulo, texto, imagem) na geometria do relatório da Ana, e é a MESMA
   lista que vira HTML aqui e PPTX na exportação. Este arquivo cuida do resto —
   navegação, conteúdo ao vivo, editor, versões, vídeo e exportações. */
const esc = s => String(s == null ? '' : s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));

/* ---- render HTML ---- */
function slideHTML(s){
  let corpo = LAYOUT.html(LAYOUT.prims(s, {linkVideo}));
  // slide sem conteúdo: o cartão (com o botão de escrever, pra quem edita) fica
  // por cima do corpo — só na tela; no PPTX vai o texto da pendência
  if (s.t === 'pendente') {
    corpo += `<div class="vazio-sobre">${editorDe(s)
      ? vazioEditavel(s)
      : vazio('◇', 'slide em aberto', s.titulo, s.sub || '', s.fonte, s.motivo)}</div>`;
  }
  if (s.oculto) corpo += `<div class="selo-oculto" title="Existe no arquivo, mas não entra na apresentação — como slide oculto no PowerPoint">oculto na apresentação</div>`;
  return corpo;
}

/* Slide de conteúdo humano ainda vazio. O cartão antigo dizia "base que vai
   alimentar: _docs/comite_conteudo.json → comentarios" e "por que ainda não tem:
   escreva o conteúdo desse mês pelo hub" — caminho de arquivo e instrução
   genérica, nenhum dos dois acionável de dentro do deck. Aqui o slide diz o que
   falta e abre o editor no clique; quem não é editor vê só o aviso. */
const NOMES_EDITOR = {
  comentarios: 'os comentários',
  manejo: 'os pontos de manejo e decisões',
  exposicoes: 'as exposições',
  fotos: 'as fotos',
  pendencias: 'as pendências da apresentação anterior',
};

function vazioEditavel(s){
  const oque = NOMES_EDITOR[editorDe(s)] || 'o conteúdo';
  return `<div class="vazio"><div class="box">
    <div class="glyph">✎</div>
    <span class="tag">a escrever</span>
    <h2>${esc(s.titulo)}</h2>
    <p>Ainda não há ${esc(oque)} deste mês.</p>
    ${souEditor
      ? `<button type="button" class="vazio-btn" data-abrir-editor="1">Escrever agora</button>`
      : `<p class="vazio-nota">Quem edita o comitê pode preencher direto por aqui.</p>`}
  </div></div>`;
}

/* Estado vazio comum aos dois casos — cartão centralizado, e não um bloco de
   texto solto no meio do slide, que era como ficava antes. */
function vazio(glyph, tag, titulo, texto, fonte, motivo){
  return `<div class="vazio"><div class="box">
    <div class="glyph">${glyph}</div>
    <span class="tag">${esc(tag)}</span>
    <h2>${esc(titulo)}</h2>
    ${texto ? `<p>${esc(texto)}</p>` : ''}
    ${fonte ? `<div class="campos">
      <div><div class="l">Base que vai alimentar</div><div class="v">${esc(fonte)}</div></div>
      <div><div class="l">Por que ainda não tem</div><div class="v">${esc(motivo)}</div></div>
    </div>` : ''}
  </div></div>`;
}

/* ==================== conteúdo ao vivo (Supabase) ====================
   Comentários, exposições, manejo e fotos são "sem fonte" (ver cabeçalho de
   tools/build_comite.py) — editados pelo hub desde 31/08/2026, não mais só
   no _docs/comite_conteudo.json. O SPEC (spec.js) é o que o PIPELINE gerou
   da última vez que alguém rodou — pode estar atrasado em relação ao
   Supabase, que é a fonte principal agora. Pra editor E visualizador verem
   sempre o mais recente sem depender de alguém lembrar de rodar o build de
   novo, o deck busca o mês corrente no Supabase toda vez que troca de mês
   (ou salva algo) e RECONSTRÓI esses 4 tipos de slide na hora — mesma regra
   de slide_comentarios/slides_exposicoes/slide_manejo/_fotos_grupo_por_tema
   do build_comite.py, só que em JS. Se o Supabase não responder ou o mês
   não tiver linha lá, o SPEC (baked) continua valendo — não quebra nada. */
/* ---- vídeo no slide de fotos ----
   Vídeo entra na MESMA grade das fotos: o que aparece no slide é o primeiro
   frame (poster), e o clique abre o vídeo. O poster é gerado no navegador de
   quem sobe e guardado ao lado do arquivo (`<path>.poster.jpg`), por dois
   motivos: o deck não precisa decodificar vídeo pra desenhar o slide, e o PPTX
   tem uma imagem pronta pra colar — o pptxgen não extrai frame nenhum.

   O arquivo continua sendo um path no bucket privado, então o formato gravado
   em `comite_conteudo.fotos` não muda: segue um array de strings. Vídeo se
   reconhece pela extensão. */
const EXT_VIDEO = /\.(mp4|webm|mov|m4v|ogv)$/i;
const ehVideo = p => EXT_VIDEO.test(String(p || ''));
const posterDe = p => `${p}.poster.jpg`;
/* Link que o PPTX e o PDF carregam: abre o deck no vídeo, que lá dentro pede uma
   URL assinada nova. Não dá pra pôr a URL assinada do bucket no arquivo — ela
   vence em 1h e o .pptx circula por semanas. */
const linkVideo = p => `${location.origin}${location.pathname}?video=${encodeURIComponent(p)}`;
/* Teto por arquivo do Supabase Storage no plano atual. Medido em 23/09/2026
   subindo blobs de tamanho crescente: 50 MB passa, 51 devolve 413
   EntityTooLarge. Não é config do bucket (que está sem limite próprio) — é do
   plano, então não adianta mexer no painel. */
const LIMITE_UPLOAD = 50 * 1024 * 1024;
const ALVO_COMPRESSAO = 45 * 1024 * 1024;   // folga pro container do webm/mp4
const mb = n => (n / 1024 / 1024).toFixed(0);
/* Item da grade, normalizado: o spec baked traz string (data URI da foto) e o
   conteúdo ao vivo traz {img, video}. Os três consumidores — slide, PPTX e
   export — passam por aqui pra não divergirem. */
const itemFoto = f => (typeof f === 'string' ? {img: f, video: null} : (f || {img: null, video: null}));

const FOTOS_POR_SLIDE = 6;
const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const ABR_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const TIPOS_EDITAVEIS = new Set(['comentarios', 'manejo', 'fotos', 'resultados', 'pendencias']);
const ehExposicaoProg = s => s.t === 'tabela' && s.n === 23;

/* Qual editor abre este slide, ou null se não é de conteúdo humano.

   Slide vazio vem como t="pendente" e antes caía fora daqui — o Editar ficava
   apagado justamente no slide que precisa ser escrito. Agora o pendente carrega
   `edita` (posto pelo build) dizendo que editor ele quer. */
const editorDe = s => {
  if (!s) return null;
  if (s.t === 'pendente') return s.edita || null;
  if (s.t === 'resultados') return 'exposicoes';
  if (ehExposicaoProg(s)) return 'exposicoes';
  return TIPOS_EDITAVEIS.has(s.t) ? s.t : null;
};
const ehEditavel = s => !!editorDe(s);

function hubSb(){ try { return window.parent.HUB && window.parent.HUB.sb; } catch (e) { return null; } }
function hubEmail(){ try { return window.parent.HUB && window.parent.HUB.email; } catch (e) { return null; } }
function hubIsAdmin(){ try { return window.parent.HUB && window.parent.HUB.role === 'admin'; } catch (e) { return false; } }

let souEditor = false;
async function checaEditor(){
  if (hubIsAdmin()) {
    souEditor = true;
  } else {
    const sb = hubSb(), email = hubEmail();
    if (!sb || !email) {
      souEditor = false;
    } else {
      try {
        const { data } = await sb.from('comite_editores').select('email').eq('email', email).maybeSingle();
        souEditor = !!data;
      } catch (e) { souEditor = false; }
    }
  }
  const btn = document.getElementById('editar');
  if (btn) btn.hidden = !souEditor;
  // histórico e atualização andam com a edição: quem não escreve não precisa
  // ver versão nem disparar pipeline
  const bv = document.getElementById('versoes');
  if (bv) bv.hidden = !souEditor;
  if (souEditor) montaBarraJob();
  /* o slide vazio mostra o botao 'Escrever agora' so pra quem edita, e quem
     edita so se sabe DEPOIS desta consulta — entao repinta o slide atual */
  if (souEditor && typeof render === 'function' && slides.length) render();
}

const conteudoCache = {};   // {mes: linha do Supabase | null} — invalidado ao salvar
async function buscaConteudoAoVivo(mes){
  if (mes in conteudoCache) return conteudoCache[mes];
  const sb = hubSb();
  if (!sb) { conteudoCache[mes] = null; return null; }
  try {
    const { data, error } = await sb.from('comite_conteudo').select('*').eq('mes', mes).maybeSingle();
    conteudoCache[mes] = error ? null : data;
  } catch (e) { conteudoCache[mes] = null; }
  return conteudoCache[mes];
}

/* foto do bucket privado -> data URI, mesmo formato que o resto do arquivo já
   usa (ver dataURI() mais abaixo, reaproveitada aqui) — assim a exportação
   PPTX continua funcionando sem mudar nada nela (ela espera base64 embutido,
   não URL remota). */
async function urlAssinada(path, seg){
  const sb = hubSb();
  if (!sb) return null;
  try {
    const { data, error } = await sb.storage.from('comite-fotos').createSignedUrl(path, seg || 3600);
    return (error || !data) ? null : data.signedUrl;
  } catch (e) { return null; }
}
async function fotoDataUri(path){
  const url = await urlAssinada(path);
  return url ? await dataURI(url) : null;
}

/* Primeiro frame do vídeo, como Blob JPEG — o poster que vai pro slide, pro
   PPTX e pro PDF. Roda no navegador de quem sobe, uma vez.

   Procura o frame em ~0.6s, não em 0: o começo costuma ser preto ou tremido, e
   um quadro preto na grade parece foto que não carregou. Se o navegador não
   decodifica o formato (.mov do iPhone é o caso comum fora do Safari), devolve
   null e o item fica com a marca de play sem imagem — o vídeo continua salvo e
   tocável, só não tem capa. */
/* O navegador consegue DECODIFICAR este vídeo?

   Vale a pergunta antes de subir: se a resposta é não, não é só a capa que
   falha — o vídeo não toca no deck, e subir 40 MB de algo que ninguém vê é o
   pior desfecho. O caso comum é o .mov do iPhone gravado em HEVC: o Safari
   toca, o Chrome no Windows só com o decodificador do sistema. Testar de
   verdade (carregar e pedir um frame) é mais confiável que `canPlayType`, que
   responde "maybe" para praticamente tudo. */
function decodifica(file){
  return new Promise(resolve => {
    const url = URL.createObjectURL(file);
    const v = document.createElement('video');
    // 'auto' e não 'metadata': o veredito depende de decodificar um frame de
    // verdade, e com metadata alguns navegadores param antes disso
    v.muted = true; v.playsInline = true; v.preload = 'auto'; v.src = url;
    const fim = ok => { URL.revokeObjectURL(url); v.remove(); resolve(ok); };
    const prazo = setTimeout(() => fim(false), 15000);
    v.onerror = () => { clearTimeout(prazo); fim(false); };
    v.onloadeddata = () => { clearTimeout(prazo); fim(!!(v.videoWidth && v.videoHeight)); };
  });
}

/* Re-encoda pra caber no teto do Storage: desenha o vídeo num canvas menor e
   grava a saída com o MediaRecorder, junto com a faixa de áudio do original.
   Canvas em vez de `video.captureStream()` direto porque o stream do elemento
   sai na resolução original — e é justamente a resolução que precisa cair.

   Roda em tempo real: um vídeo de 1 minuto leva ~1 minuto. Só funciona se o
   navegador decodificar o original, então HEVC no Chrome continua de fora — o
   `decodifica()` acima é quem barra esse caso, antes de chegar aqui. */
const MIMES_SAIDA = ['video/mp4;codecs=h264,aac', 'video/webm;codecs=vp9,opus', 'video/webm;codecs=vp8,opus', 'video/webm'];

function comprimeVideo(file, aoProgredir, ladoMax, tetoBps){
  return new Promise(async resolve => {
    const mime = MIMES_SAIDA.find(m => window.MediaRecorder && MediaRecorder.isTypeSupported(m));
    if (!mime) return resolve(null);
    const url = URL.createObjectURL(file);
    const v = document.createElement('video');
    v.muted = true; v.playsInline = true; v.preload = 'auto'; v.src = url;
    const limpa = () => { URL.revokeObjectURL(url); v.remove(); };
    v.onerror = () => { limpa(); resolve(null); };
    v.onloadedmetadata = async () => {
      try {
        const dur = v.duration;
        if (!dur || !isFinite(dur)) { limpa(); return resolve(null); }
        // 720p no maior lado, mantendo proporção e lado par (exigência de codec)
        const escala = Math.min(1, (ladoMax || 1280) / Math.max(v.videoWidth, v.videoHeight));
        const par = n => Math.max(2, Math.round(n * escala / 2) * 2);
        const c = document.createElement('canvas');
        c.width = par(v.videoWidth); c.height = par(v.videoHeight);
        const ctx = c.getContext('2d');
        // bitrate que cabe no alvo, com teto de 2,5 Mbps — acima disso 720p
        // não melhora o suficiente pra justificar o tamanho
        const audioBps = 96000;
        const videoBps = Math.max(250000, Math.min(tetoBps || 2500000, (ALVO_COMPRESSAO * 8) / dur - audioBps));
        const saida = c.captureStream(30);
        const doVideo = v.captureStream ? v.captureStream() : (v.mozCaptureStream ? v.mozCaptureStream() : null);
        (doVideo ? doVideo.getAudioTracks() : []).forEach(t => saida.addTrack(t));
        const rec = new MediaRecorder(saida, {mimeType: mime, videoBitsPerSecond: videoBps, audioBitsPerSecond: audioBps});
        const pedacos = [];
        rec.ondataavailable = e => { if (e.data && e.data.size) pedacos.push(e.data); };
        rec.onstop = () => { limpa(); resolve(new Blob(pedacos, {type: mime.split(';')[0]})); };
        let vivo = true;
        const pinta = () => {
          if (!vivo) return;
          ctx.drawImage(v, 0, 0, c.width, c.height);
          if (aoProgredir && dur) aoProgredir(Math.min(1, v.currentTime / dur));
          requestAnimationFrame(pinta);
        };
        v.onended = () => { vivo = false; if (rec.state !== 'inactive') rec.stop(); };
        rec.start(1000);
        await v.play();
        pinta();
      } catch (e) { limpa(); resolve(null); }
    };
  });
}

/* Escada de qualidade: tenta 720p, e só desce se o resultado ainda não coube.
   Sem Supabase Pro o teto de 50 MB é fixo, então vale mais entregar um vídeo
   de 480p do que recusar um de 4 minutos — no slide ele aparece em 1/6 da tela
   e a diferença mal se nota. Cada degrau custa o tempo de duração do vídeo,
   por isso o primeiro chute já é calculado pelo bitrate que caberia. */
const DEGRAUS = [[1280, 2500000], [854, 1200000], [640, 700000]];

async function comprimeAteCaber(file, diz){
  let ultimo = null;
  for (const [lado, teto] of DEGRAUS) {
    const rot = lado >= 1280 ? '720p' : (lado >= 854 ? '480p' : '360p');
    const saida = await comprimeVideo(file, f => diz(`convertendo em ${rot} — ${Math.round(f * 100)}%`), lado, teto);
    if (!saida) return ultimo;               // navegador não deu conta: para por aqui
    ultimo = saida;
    if (saida.size <= LIMITE_UPLOAD) return saida;
    diz(`${rot} ficou em ${mb(saida.size)} MB — tentando menor…`);
  }
  return ultimo;
}

function capturaPoster(file){
  return new Promise(resolve => {
    const url = URL.createObjectURL(file);
    const v = document.createElement('video');
    v.muted = true; v.playsInline = true; v.preload = 'auto'; v.src = url;
    const fim = blob => { URL.revokeObjectURL(url); v.remove(); resolve(blob); };
    const erro = () => fim(null);
    const prazo = setTimeout(erro, 15000);
    v.onerror = erro;
    v.onloadeddata = () => { v.currentTime = Math.min(0.6, (v.duration || 1) / 2); };
    v.onseeked = () => {
      clearTimeout(prazo);
      try {
        const c = document.createElement('canvas');
        c.width = v.videoWidth; c.height = v.videoHeight;
        if (!c.width || !c.height) return erro();
        c.getContext('2d').drawImage(v, 0, 0, c.width, c.height);
        c.toBlob(b => fim(b), 'image/jpeg', 0.85);
      } catch (e) { erro(); }
    };
  });
}

/* monta os slides sem fonte a partir do conteúdo ao vivo — mesma regra do
   Python. Devolve {comentarios:[...], exposicoes:[...], manejo:[...],
   fotos:[...]} só com as chaves que têm conteúdo; ausente = spec baked vale. */
/* tamanho da imagem, pra o layout pôr a foto inteira na proporção dela */
function medeImagem(uri){
  return new Promise(resolve => {
    if (!uri) return resolve({w: null, h: null});
    const im = new Image();
    im.onload = () => resolve({w: im.naturalWidth, h: im.naturalHeight});
    im.onerror = () => resolve({w: null, h: null});
    im.src = uri;
  });
}

async function montaSlidesAoVivo(mes){
  const c = await buscaConteudoAoVivo(mes);
  if (!c) return null;
  const [ano, mNum] = mes.split('-').map(Number);
  const out = {};
  const MES = MESES_PT[mNum - 1];

  // mesmos títulos e subtítulos do build_comite.py — o slide ao vivo substitui
  // o do spec e não pode trocar de cara no meio da apresentação
  if (c.comentarios && c.comentarios.length) {
    out.comentarios = [{t:'comentarios', n:8, titulo:`COMENTÁRIOS — ${MES.toUpperCase()} ${ano}`,
      sub:'DRE 2026 | HPG  ·  Principais destaques do mês por categoria', itens: c.comentarios}];
  }

  const pend = (c.pendencias || []).filter(x => String(x || '').trim());
  if (pend.length) {
    out.pendencias = [{t:'pendencias', n:3,
      titulo:`PENDÊNCIAS DA APRESENTAÇÃO DE ${MESES_PT[(mNum + 10) % 12].toUpperCase()}`, itens: pend}];
  }

  const exp = c.exposicoes || {}, prog = exp.programacao || [], res = exp.resultados || [];
  if (prog.length || res.length) {
    const s = [];
    if (prog.length) s.push({t:'tabela', n:23, titulo:`EXPOSIÇÕES ${ano} — PROGRAMAÇÃO`,
      sub:'Calendário de participações previstas', cols:['EVENTO','DATA','LOCAL','STATUS'], rows: prog});
    res.forEach((r, k) => s.push({t:'resultados', n:24+k, titulo:r.titulo, sub:r.sub || '', animais:r.animais}));
    out.exposicoes = s;
  }

  if (c.manejo && c.manejo.length) {
    out.manejo = [{t:'manejo', n:38, titulo:'MANEJO — PONTOS DE MELHORIA E DECISÕES',
      sub:`Histórico de intervenções Jan–${ABR_PT[mNum-1]} ${ano}`, itens: c.manejo, atual: ABR_PT[mNum-1]}];
  }

  if (c.fotos && c.fotos.length) {
    const grupos = typeof c.fotos[0] === 'string' ? [{tema:'', arquivos:c.fotos}] : c.fotos;
    const s = [];
    for (const g of grupos) {
      /* Foto vira data URI (é o que o PPTX cola). Vídeo vira {img: poster,
         video: path}: o arquivo em si nunca é embutido — o slide mostra a capa
         e o clique busca uma URL assinada na hora. Vídeo sem poster entra
         assim mesmo, com a marca de play e sem capa. */
      const itens = (await Promise.all((g.arquivos || []).map(async a => {
        const img = await fotoDataUri(ehVideo(a) ? posterDe(a) : a);
        if (!img && !ehVideo(a)) return null;
        return Object.assign({img, video: ehVideo(a) ? a : null}, await medeImagem(img));
      }))).filter(Boolean);
      if (!itens.length) continue;
      const n = Math.ceil(itens.length / FOTOS_POR_SLIDE);
      for (let k = 0; k < n; k++) {
        const bloco = itens.slice(k * FOTOS_POR_SLIDE, (k + 1) * FOTOS_POR_SLIDE);
        let sub = g.tema ? `Obras e melhorias realizadas  ·  ${g.tema}` : `Registros de ${MES} ${ano}`;
        if (n > 1) sub += ` (${k+1}/${n})`;
        s.push({t:'fotos', n:39, titulo:'MANEJO — FOTOS E REGISTROS', sub, fotos:bloco});
      }
    }
    if (s.length) out.fotos = s;
  }
  return out;
}

/* troca, na lista `slides` já carregada, os slides de um tipo pelos novos —
   por tipo (comentarios/manejo/fotos) ou pelo par tabela(n=23)+resultados
   (exposições, que são dois tipos de slide pra 1 conteúdo só). */
/* Troca os slides de um editor pelos recem-salvos, na posicao em que ja
   estavam. O casamento e por `editorDe`, nao por `s.t`: o slide AINDA VAZIO vem
   como t='pendente' (com `edita`), entao comparar `s.t === chave` nao o
   reconhecia — ele sobrevivia ao filtro, `primeiro` ficava -1 e o slide escrito
   ia parar no FIM do deck. Da tela, salvar nao mudava nada: o lugar de sempre
   continuava dizendo "a escrever". */
function substituiSlidesDoTipo(chave, novos){
  let primeiro = -1;
  slides = slides.filter((s, i) => {
    const bate = editorDe(s) === chave;
    if (bate && primeiro === -1) primeiro = i;
    return !bate;
  });
  if (primeiro === -1) primeiro = slides.length;
  slides.splice(primeiro, 0, ...novos);
}

async function aplicaConteudoAoVivo(mes){
  const vivo = await montaSlidesAoVivo(mes);
  if (mes !== mesAtual || !vivo) return;   // usuário já trocou de mês, ou nada pra aplicar
  for (const [chave, novos] of Object.entries(vivo)) substituiSlidesDoTipo(chave, novos);
  listaSlides();
  go(Math.min(idx, slides.length - 1));
}

/* ==================== editor ====================
   Um painel só, corpo diferente por tipo. `estado` guarda os dados sendo
   editados (array de linhas, ou objeto — depende do tipo); os inputs
   escrevem direto em `estado` via listener delegado (sem re-render a cada
   tecla), e só reconstrói o HTML quando uma linha é adicionada/removida. */
let estado = null, tipoAtual = null;

function painelEditor(){
  let el = document.getElementById('editorOverlay');
  if (el) return el;
  el = document.createElement('div');
  el.id = 'editorOverlay';
  el.innerHTML = `<div id="editorPanel">
    <div class="ed-head"><h3 id="edTitulo"></h3>
      <button type="button" id="edFechar" aria-label="Fechar">✕</button></div>
    <div id="edCorpo"></div>
    <div class="ed-rodape">
      <span id="edStatus"></span>
      <button type="button" id="edCancelar">Cancelar</button>
      <button type="button" id="edSalvar" class="primary">Salvar</button>
    </div>
  </div>`;
  document.body.appendChild(el);
  el.addEventListener('click', e => { if (e.target === el) fechaEditor(); });
  document.getElementById('edFechar').onclick = fechaEditor;
  document.getElementById('edCancelar').onclick = fechaEditor;
  document.getElementById('edSalvar').onclick = salvaEditor;
  // delegado: linhas de comentários (data-i/data-f) escrevem direto em
  // `estado` sem precisar de handler próprio por input. Os outros tipos
  // (manejo/exposições/fotos) usam atributos diferentes (data-mi, data-pi,
  // data-gi...) com oninput própria montada em cada render — não passam
  // por aqui, o filtro abaixo já ignora.
  document.getElementById('edCorpo').addEventListener('input', e => {
    const t = e.target, i = t.dataset.i, f = t.dataset.f;
    if (i == null) return;
    estado[+i][f] = t.value;
  });
  return el;
}
function fechaEditor(){
  const el = document.getElementById('editorOverlay');
  if (el) el.style.display = 'none';
}

async function abreEditor(s){
  if (!souEditor || !ehEditavel(s)) return;
  const el = painelEditor();
  el.style.display = 'flex';
  document.getElementById('edStatus').textContent = 'carregando…';
  document.getElementById('edSalvar').disabled = true;
  const c = (await buscaConteudoAoVivo(mesAtual)) || {};

  const alvo = editorDe(s);
  if (alvo === 'comentarios') {
    tipoAtual = 'comentarios';
    estado = JSON.parse(JSON.stringify(c.comentarios || []));
    document.getElementById('edTitulo').textContent = `Comentários — ${SPEC.labels[mesAtual]}`;
    renderComentarios();
  } else if (alvo === 'manejo') {
    tipoAtual = 'manejo';
    estado = JSON.parse(JSON.stringify(c.manejo || []));
    document.getElementById('edTitulo').textContent = `Manejo — ${SPEC.labels[mesAtual]}`;
    renderManejo();
  } else if (alvo === 'exposicoes') {
    tipoAtual = 'exposicoes';
    const exp = c.exposicoes || {};
    estado = {programacao: JSON.parse(JSON.stringify(exp.programacao || [])),
              resultados: JSON.parse(JSON.stringify(exp.resultados || []))};
    document.getElementById('edTitulo').textContent = `Exposições — ${SPEC.labels[mesAtual]}`;
    renderExposicoes();
  } else if (alvo === 'pendencias') {
    tipoAtual = 'pendencias';
    estado = JSON.parse(JSON.stringify(c.pendencias || []));
    document.getElementById('edTitulo').textContent = `Pendências da apresentação anterior — ${SPEC.labels[mesAtual]}`;
    renderPendencias();
  } else if (alvo === 'fotos') {
    tipoAtual = 'fotos';
    const fotos = c.fotos || [];
    estado = JSON.parse(JSON.stringify(
      typeof fotos[0] === 'string' ? [{tema:'', arquivos: fotos}] : fotos));
    document.getElementById('edTitulo').textContent = `Fotos — ${SPEC.labels[mesAtual]}`;
    await renderFotos();
  }
  document.getElementById('edStatus').textContent = '';
  document.getElementById('edSalvar').disabled = false;
}

/* ---- comentários: linhas {cat, txt, delta} ---- */
function renderComentarios(){
  const corpo = document.getElementById('edCorpo');
  corpo.innerHTML = estado.map((r, i) => `
    <div class="ed-linha">
      <input data-i="${i}" data-f="cat" value="${escAttr(r.cat)}" placeholder="Categoria">
      <textarea data-i="${i}" data-f="txt" placeholder="Texto do comentário" rows="2">${esc(r.txt || '')}</textarea>
      <input data-i="${i}" data-f="delta" value="${escAttr(r.delta)}" placeholder="+R$Xk" class="ed-curta">
      <button type="button" class="ed-rm" data-rm="${i}">✕</button>
    </div>`).join('') + `<button type="button" id="edAdd" class="ed-add">+ categoria</button>`;
  corpo.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rm, 1); renderComentarios(); });
  document.getElementById('edAdd').onclick = () => { estado.push({cat:'', txt:'', delta:''}); renderComentarios(); };
}

/* ---- pendências: uma frase por item, na ordem em que vão pro slide ---- */
function renderPendencias(){
  const corpo = document.getElementById('edCorpo');
  corpo.innerHTML = estado.map((t, i) => `
    <div class="ed-linha ed-linha2">
      <textarea data-pi="${i}" rows="2" placeholder="O que ficou combinado na apresentação anterior">${esc(t || '')}</textarea>
      <button type="button" class="ed-rm" data-rm="${i}">✕</button>
    </div>`).join('') + `<button type="button" id="edAdd" class="ed-add">+ pendência</button>`;
  corpo.querySelectorAll('[data-pi]').forEach(inp => inp.oninput = () => { estado[+inp.dataset.pi] = inp.value; });
  corpo.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rm, 1); renderPendencias(); });
  document.getElementById('edAdd').onclick = () => { estado.push(''); renderPendencias(); };
}

/* ---- manejo: linhas [mes, texto] ---- */
function renderManejo(){
  const corpo = document.getElementById('edCorpo');
  corpo.innerHTML = estado.map((r, i) => `
    <div class="ed-linha">
      <input data-mi="${i}" data-mf="0" value="${escAttr(r[0])}" placeholder="Mês" class="ed-curta">
      <textarea data-mi="${i}" data-mf="1" placeholder="Intervenções e decisões" rows="2">${esc(r[1] || '')}</textarea>
      <button type="button" class="ed-rm" data-rm="${i}">✕</button>
    </div>`).join('') + `<button type="button" id="edAdd" class="ed-add">+ mês</button>`;
  corpo.querySelectorAll('[data-mi]').forEach(inp => inp.oninput = () => {
    estado[+inp.dataset.mi][+inp.dataset.mf] = inp.value;
  });
  corpo.querySelectorAll('[data-rm]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rm, 1); renderManejo(); });
  document.getElementById('edAdd').onclick = () => { estado.push(['', '']); renderManejo(); };
}

/* ---- exposições: programação (linhas [evento,data,local,status]) +
   resultados ({titulo,sub,animais}, animais editado como texto "Nome:
   prêmio1; prêmio2" por linha — mais simples que formulário aninhado) ---- */
function animaisParaTexto(animais){
  return (animais || []).map(a => `${a.nome}: ${(a.premios || []).join('; ')}`).join('\n');
}
function textoParaAnimais(txt){
  return txt.split('\n').map(l => l.trim()).filter(Boolean).map(l => {
    const [nome, resto] = l.split(':');
    return {nome: (nome || '').trim(), premios: (resto || '').split(';').map(p => p.trim()).filter(Boolean)};
  });
}
function renderExposicoes(){
  const corpo = document.getElementById('edCorpo');
  const prog = estado.programacao.map((r, i) => `
    <div class="ed-linha ed-linha4">
      <input data-pi="${i}" data-pf="0" value="${escAttr(r[0])}" placeholder="Evento">
      <input data-pi="${i}" data-pf="1" value="${escAttr(r[1])}" placeholder="Data">
      <input data-pi="${i}" data-pf="2" value="${escAttr(r[2])}" placeholder="Local">
      <input data-pi="${i}" data-pf="3" value="${escAttr(r[3])}" placeholder="Status">
      <button type="button" class="ed-rm" data-rmp="${i}">✕</button>
    </div>`).join('');
  const res = estado.resultados.map((r, i) => `
    <div class="ed-bloco">
      <input data-ri="${i}" data-rf="titulo" value="${escAttr(r.titulo)}" placeholder="Título do slide">
      <input data-ri="${i}" data-rf="sub" value="${escAttr(r.sub)}" placeholder="Subtítulo">
      <textarea data-ri="${i}" data-rf="animais" rows="4" placeholder="Um animal por linha: Nome: prêmio 1; prêmio 2">${esc(animaisParaTexto(r.animais))}</textarea>
      <button type="button" class="ed-rm" data-rmr="${i}">✕ remover resultado</button>
    </div>`).join('');
  corpo.innerHTML = `<h4>Programação</h4>${prog}
    <button type="button" id="edAddP" class="ed-add">+ evento</button>
    <h4>Resultados</h4>${res}
    <button type="button" id="edAddR" class="ed-add">+ resultado</button>`;
  corpo.querySelectorAll('[data-pi]').forEach(inp => inp.oninput = () => {
    estado.programacao[+inp.dataset.pi][+inp.dataset.pf] = inp.value;
  });
  corpo.querySelectorAll('[data-rmp]').forEach(b => b.onclick = () => { estado.programacao.splice(+b.dataset.rmp, 1); renderExposicoes(); });
  corpo.querySelectorAll('[data-ri]').forEach(inp => inp.oninput = () => {
    const r = estado.resultados[+inp.dataset.ri], f = inp.dataset.rf;
    r[f === 'animais' ? '_txt' : f] = inp.value;
  });
  corpo.querySelectorAll('[data-rmr]').forEach(b => b.onclick = () => { estado.resultados.splice(+b.dataset.rmr, 1); renderExposicoes(); });
  document.getElementById('edAddP').onclick = () => { estado.programacao.push(['', '', '', '']); renderExposicoes(); };
  document.getElementById('edAddR').onclick = () => { estado.resultados.push({titulo:'', sub:'', animais:[]}); renderExposicoes(); };
}

/* ---- fotos: grupos {tema, arquivos[]} — upload direto no bucket ao
   escolher arquivo (não espera o Salvar geral, senão perde o arquivo se
   fechar sem salvar); remover/reordenar só mexe no array em memória até
   Salvar gravar a lista final. ---- */
async function renderFotos(){
  const corpo = document.getElementById('edCorpo');
  const blocos = await Promise.all(estado.map(async (g, gi) => {
    const thumbs = await Promise.all((g.arquivos || []).map(async (p, ai) => {
      // vídeo mostra o poster que subiu com ele; sem poster, fica só a marca
      const v = ehVideo(p);
      const uri = await fotoDataUri(v ? posterDe(p) : p);
      return `<div class="ed-foto${v ? ' ed-video' : ''}">
        <img src="${uri || ''}" alt="">
        ${v ? '<span class="ed-play" aria-hidden="true"></span>' : ''}
        <div class="ed-foto-acoes">
          <button type="button" data-mv="${gi}:${ai}:-1" ${ai === 0 ? 'disabled' : ''}>↑</button>
          <button type="button" data-mv="${gi}:${ai}:1" ${ai === g.arquivos.length - 1 ? 'disabled' : ''}>↓</button>
          ${v ? `<label class="ed-capa" title="Trocar a capa do vídeo">capa<input type="file"
             accept="image/*" data-capa="${escAttr(p)}" hidden></label>` : ''}
          <button type="button" data-rmf="${gi}:${ai}">✕</button>
        </div></div>`;
    }));
    return `<div class="ed-grupo">
      <div class="ed-grupo-head">
        <input data-gi="${gi}" value="${escAttr(g.tema)}" placeholder="Tema (opcional)">
        <label class="ed-upload">+ foto/vídeo<input type="file" accept="image/*,video/*" multiple data-up="${gi}" hidden></label>
        <button type="button" class="ed-rm" data-rmg="${gi}">✕ remover grupo</button>
      </div>
      <div class="ed-fotos-grade">${thumbs.join('') || '<span class="ed-vazio">sem foto neste grupo</span>'}</div>
    </div>`;
  }));
  corpo.innerHTML = blocos.join('') + `<button type="button" id="edAddG" class="ed-add">+ grupo (tema)</button>`;

  corpo.querySelectorAll('[data-gi]').forEach(inp => inp.oninput = () => { estado[+inp.dataset.gi].tema = inp.value; });
  corpo.querySelectorAll('[data-rmg]').forEach(b => b.onclick = () => { estado.splice(+b.dataset.rmg, 1); renderFotos(); });
  corpo.querySelectorAll('[data-rmf]').forEach(b => b.onclick = () => {
    const [gi, ai] = b.dataset.rmf.split(':').map(Number);
    estado[gi].arquivos.splice(ai, 1); renderFotos();
  });
  corpo.querySelectorAll('[data-mv]').forEach(b => b.onclick = () => {
    const [gi, ai, dir] = b.dataset.mv.split(':').map(Number);
    const arr = estado[gi].arquivos, novo = ai + dir;
    if (novo < 0 || novo >= arr.length) return;
    [arr[ai], arr[novo]] = [arr[novo], arr[ai]];
    renderFotos();
  });
  corpo.querySelectorAll('[data-up]').forEach(inp => inp.onchange = async () => {
    const gi = +inp.dataset.up;
    // vídeo demora: o status conta quantos faltam, senão parece travado
    const fila = [...inp.files];
    const st = document.getElementById('edStatus');
    for (const [k, file] of fila.entries()) {
      const pre = fila.length > 1 ? `${k + 1}/${fila.length} · ` : '';
      const path = await sobeFoto(file, t => { st.textContent = pre + t; });
      if (path) estado[gi].arquivos.push(path);
    }
    st.textContent = '';
    renderFotos();
  });
  /* capa manual: grava por cima do poster automático, mesmo nome */
  corpo.querySelectorAll('[data-capa]').forEach(inp => inp.onchange = async () => {
    const st = document.getElementById('edStatus');
    st.textContent = 'trocando a capa…';
    await trocaCapa(inp.dataset.capa, inp.files[0]);
    st.textContent = '';
    renderFotos();
  });
  document.getElementById('edAddG').onclick = () => { estado.push({tema:'', arquivos:[]}); renderFotos(); };
}
async function sobeArquivo(path, file, contentType){
  const sb = hubSb();
  if (!sb) return false;
  try {
    const { error } = await sb.storage.from('comite-fotos')
      .upload(path, file, {contentType: contentType || file.type, upsert: true});
    if (error) { alert('Falha no upload: ' + error.message); return false; }
    return true;
  } catch (e) { alert('Falha no upload: ' + e.message); return false; }
}

/* Sobe foto ou vídeo. No vídeo vai junto o poster, com o nome do próprio
   arquivo + '.poster.jpg' — assim quem lê sabe onde procurar sem guardar mais
   nada no JSON.

   Antes de gastar upload, dois portões: o navegador precisa decodificar o
   arquivo (senão nem capa nem playback existem) e ele precisa caber no teto do
   Storage — passando disso, tenta encolher na hora. Quem não passa é recusado
   com o motivo, não com um erro de rede no meio do envio. */
async function sobeFoto(file, aviso){
  const diz = t => { if (aviso) aviso(t); };
  const video = (file.type || '').startsWith('video/') || ehVideo(file.name);
  let dados = file, nomeBase = file.name, poster = null;

  if (video) {
    diz('conferindo o vídeo…');
    if (!await decodifica(file)) {
      alert(`"${file.name}" não abre neste navegador — provavelmente HEVC, o formato padrão do iPhone.\n\n`
        + 'Três saídas, da mais simples pra menos:\n'
        + '1. No iPhone: Ajustes › Câmera › Formatos › "Mais Compatível". Grava em H.264 e o problema '
        + 'não volta mais.\n'
        + '2. Subir este mesmo vídeo pelo Safari (Mac ou iPhone), que decodifica HEVC — eu converto '
        + 'sozinho na hora do upload.\n'
        + '3. Converter o arquivo pra MP4/H.264 antes de subir.\n\n'
        + 'Não subi o arquivo: do jeito que está, ele também não tocaria no deck.');
      return null;
    }
    /* Re-encoda também quando o formato não é universal, mesmo cabendo no
       limite. É o caso do .mov/HEVC subido pelo Safari (que decodifica): sem
       isso ele passa aqui e falha na hora de assistir, no Chrome de outra
       pessoa — o pior lugar pra descobrir. Re-encodar entrega H.264 ou VP9,
       que tocam em todo lugar. */
    const formatoArriscado = /\.mov$/i.test(file.name) || (file.type || '') === 'video/quicktime';
    if (file.size > LIMITE_UPLOAD || formatoArriscado) {
      diz(formatoArriscado && file.size <= LIMITE_UPLOAD
        ? 'convertendo pra formato compatível (leva o tempo do vídeo)…'
        : `comprimindo (${mb(file.size)} MB — leva o tempo do vídeo)…`);
      const menor = await comprimeAteCaber(file, diz);
      // formato arriscado que já cabia: se a conversão falhar, sobe o original
      // mesmo — toca pra quem decodifica, e é melhor que perder o vídeo
      if (!menor && formatoArriscado && file.size <= LIMITE_UPLOAD) {
        diz('não deu pra converter — subindo o original');
      } else if (!menor || menor.size > LIMITE_UPLOAD) {
        alert(`"${file.name}" tem ${mb(file.size)} MB e o limite por arquivo é ${mb(LIMITE_UPLOAD)} MB.\n\n`
          + (menor ? `Mesmo em 360p ficou em ${mb(menor.size)} MB. ` : 'A conversão automática não funcionou aqui. ')
          + 'Corte um trecho mais curto — no slide o vídeo ocupa 1/6 da tela, então 20 ou 30 segundos '
          + 'costumam bastar.');
        return null;
      } else {
        dados = menor;
        // a extensão tem de casar com o formato de saída, senão o deck não
        // reconhece o item como vídeo (o tipo sai do nome do arquivo)
        const ext = (menor.type || '').includes('mp4') ? 'mp4' : 'webm';
        nomeBase = file.name.replace(/\.[^.]+$/, '') + '.' + ext;
        diz(`convertido: ${mb(file.size)} MB → ${mb(menor.size)} MB`);
      }
    }
    poster = await capturaPoster(dados) || await capturaPoster(file);
  }

  const nome = `${Date.now()}_${nomeBase}`.replace(/[^\w.-]/g, '_');
  const path = `${mesAtual}/${nome}`;
  diz('enviando…');
  if (!await sobeArquivo(path, dados, dados.type || file.type)) return null;
  if (poster) await sobeArquivo(posterDe(path), poster, 'image/jpeg');
  return path;
}

/* Capa escolhida à mão — para quando o frame automático sai preto, tremido ou
   simplesmente feio. Grava no mesmo nome que o automático usaria, então o
   slide, o PPTX e o build pegam a nova sem saber a diferença. */
async function trocaCapa(pathVideo, file){
  if (!file || !file.type.startsWith('image/')) return false;
  return await sobeArquivo(posterDe(pathVideo), file, file.type);
}

async function salvaEditor(){
  const sb = hubSb();
  if (!sb) { alert('Sem sessão do hub — não dá pra salvar.'); return; }
  document.getElementById('edStatus').textContent = 'salvando…';
  document.getElementById('edSalvar').disabled = true;
  let coluna, valor;
  if (tipoAtual === 'comentarios') { coluna = 'comentarios'; valor = estado; }
  else if (tipoAtual === 'manejo') { coluna = 'manejo'; valor = estado; }
  else if (tipoAtual === 'fotos') { coluna = 'fotos'; valor = estado; }
  else if (tipoAtual === 'pendencias') { coluna = 'pendencias'; valor = estado.map(x => String(x || '').trim()).filter(Boolean); }
  else if (tipoAtual === 'exposicoes') {
    coluna = 'exposicoes';
    valor = {
      programacao: estado.programacao,
      resultados: estado.resultados.map(r => ({titulo: r.titulo, sub: r.sub, animais: textoParaAnimais(r._txt ?? animaisParaTexto(r.animais))})),
    };
  }
  const { error } = await sb.from('comite_conteudo').upsert({mes: mesAtual, [coluna]: valor});
  document.getElementById('edSalvar').disabled = false;
  if (error) { document.getElementById('edStatus').textContent = ''; alert('Falha ao salvar: ' + error.message); return; }
  delete conteudoCache[mesAtual];
  fechaEditor();
  aplicaConteudoAoVivo(mesAtual);
}

function escAttr(s){ return esc(s).replace(/"/g, '&quot;'); }

/* ---- navegação ---- */
let mesAtual = SPEC.padrao;
/* mensal = deck inteiro; trimestral = sem os slides que recortam o MÊS.
   O comitê trimestral olha o acumulado do ANO, não o do trimestre — então os
   slides de YTD, que já existem, são justamente os que ficam. */
let modo = 'mensal';
const doModo = d => modo === 'trimestral' ? d.filter(x => !x.so_mensal) : d;
let slides = doModo(SPEC.decks[mesAtual]);
let idx = 0;
const stage = document.getElementById('stage');

function render(){
  const s = slides[idx], n = slides.length;
  stage.innerHTML = `<div class="slide" id="slide" style="background:#${LAYOUT.fundo(s)}">${slideHTML(s)}</div>`;
  fit();
  document.getElementById('pos').textContent = `${idx + 1} / ${n}`;
  document.getElementById('prev').disabled = idx === 0;
  document.getElementById('next').disabled = idx === n - 1;
  const sel = document.getElementById('ir');
  if (sel.value !== String(idx)) sel.value = String(idx);
  location.hash = `#${mesAtual}/${idx + 1}`;
  const btnEditar = document.getElementById('editar');
  if (btnEditar) btnEditar.disabled = !ehEditavel(s);
}
/* escala o slide de 1280×720 pra caber na área disponível, mantendo a proporção.
   translate(-50%,-50%) faz parte do transform (não só do CSS base) porque
   setar style.transform aqui SUBSTITUI o valor inteiro — só "scale(...)"
   perdia a centralização do CSS. Achado em 31/08/2026: com place-items:center
   (grid), o slide de 1280px dentro de um stage de celular (bem menor) saía
   fora da tela — grid/flex têm min-width:auto implícito no filho, que recusa
   encolher o item pra centralizar certo quando ele é muito maior que o
   container. position:absolute + translate(-50%,-50%) centraliza certo
   sempre, não depende do tamanho relativo dos dois. */
function fit(){
  const el = document.getElementById('slide');
  if (!el) return;
  const box = stage.getBoundingClientRect();
  const pad = document.body.classList.contains('play') ? 0 : 32;
  const s = Math.min((box.width - pad) / 1280, (box.height - pad) / 720);
  el.style.transform = `translate(-50%,-50%) scale(${s})`;
}
const go = i => { idx = Math.max(0, Math.min(slides.length - 1, i)); render(); };
/* anda um slide; na apresentação pula o oculto, fora dela mostra tudo */
function anda(d){
  let i = idx + d;
  if (document.body.classList.contains('play'))
    while (i > 0 && i < slides.length - 1 && slides[i].oculto) i += d;
  go(i);
}

function trocaMes(mes){
  if (!SPEC.decks[mes]) return;
  mesAtual = mes;
  slides = doModo(SPEC.decks[mes]);
  document.getElementById('mes').value = mes;
  listaSlides();
  atualizaAviso();
  go(Math.min(idx, slides.length - 1));
  aplicaConteudoAoVivo(mes);   // async, de propósito — reflete quando chegar, sem travar a troca
}
function listaSlides(){
  document.getElementById('ir').innerHTML = slides
    .map((s, i) => `<option value="${i}">${String(i + 1).padStart(2, '0')} · ${esc((s.titulo || s.t).slice(0, 52))}${s.oculto ? ' (oculto)' : ''}</option>`).join('');
}
function atualizaAviso(){
  const p = slides.filter(s => s.t === 'pendente').length;
  document.getElementById('aviso').textContent =
    (p ? `${p} de ${slides.length} slides sem fonte` : 'todos os slides com fonte') +
    (SPEC.avisos && SPEC.avisos.length ? ` · ${SPEC.avisos[0]}` : '');
}

window.addEventListener('resize', fit);
/* Atalho de slide só vale FORA de campo de texto e com o editor fechado.
   Antes o guard era só SELECT: digitando no editor, o espaço virava "próximo
   slide" com preventDefault — não dava pra escrever "VOLUMOSO E CONCENTRADO" —
   e as setas, Home/End e 'p' também mexiam no deck por baixo do modal. */
const _digitando = el => !!el && (el.isContentEditable
  || el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT');
const _editorAberto = () => {
  const ov = document.getElementById('editorOverlay');
  return !!ov && ov.style.display !== 'none';
};
document.addEventListener('keydown', e => {
  if (_digitando(e.target) || _editorAberto()) return;
  if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') { anda(1); e.preventDefault(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { anda(-1); e.preventDefault(); }
  else if (e.key === 'Home') go(0);
  else if (e.key === 'End') go(slides.length - 1);
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

/* ---- export PDF (impressão do navegador) ----
   Terceira saída do mesmo spec. Não rasteriza nada e não precisa de
   biblioteca: monta TODOS os slides do mês num container de impressão e
   deixa o navegador paginar — texto continua texto (selecionável, nítido em
   qualquer zoom) e as fotos vão embutidas, que é o que já vem no spec.
   1280×720px a 96dpi = 13,333×7,5in, exatamente o @page — uma página por
   slide, sem escala e sem margem, então o PDF sai igual ao deck na tela.
   Vale pro que está sendo visto: mês e modo (mensal/trimestral) atuais.
   No hub o deck roda dentro de um iframe, e print() chamado DE DENTRO do
   frame imprime só o frame (Chrome/Edge/Firefox) — é por isso que a chamada
   fica aqui e não no window.parent, que imprimiria a página do hub. */
function exportarPdf(){
  let caixa = document.getElementById('impressao');
  if (!caixa) {
    caixa = document.createElement('div');
    caixa.id = 'impressao';
    document.body.appendChild(caixa);
  }
  caixa.innerHTML = slides.filter(s => !s.oculto)
    .map(s => `<div class="slide" style="background:#${LAYOUT.fundo(s)}">${slideHTML(s)}</div>`).join('');
  document.body.classList.add('imprimindo');
  /* afterprint dispara tanto no salvar quanto no cancelar; sem a limpeza o
     DOM ficaria com 70 slides pendurados atrás da tela. */
  const limpa = () => {
    document.body.classList.remove('imprimindo');
    caixa.innerHTML = '';
    window.removeEventListener('afterprint', limpa);
  };
  window.addEventListener('afterprint', limpa);
  window.print();
}

/* ---- export PPTX (mesmo spec, outra saída) ---- */
async function exportarPptx(btn){
  btn.disabled = true; btn.textContent = 'Gerando…';
  try {
    const p = new PptxGenJS();
    p.defineLayout({name:'HPG', width:10, height:5.625});      // 16:9, igual ao deck original
    p.layout = 'HPG';
    p.title = `Relatório de Desempenho Estratégico — ${SPEC.labels[mesAtual]}`;
    // logos viram base64 antes: o pptxgen não busca arquivo sozinho. As fotos
    // já vêm como data URI no spec; vídeo entra pelo poster + hyperlink.
    const imgs = {};
    for (const src of [LAYOUT.LOGO_OURO, LAYOUT.LOGO_NAVY]) imgs[src] = await dataURI(src);
    slides.forEach(s => pptSlide(p, s, imgs));
    const rotulo = modo === 'trimestral' ? 'TRIMESTRAL' : 'MENSAL';
    await p.writeFile({fileName: `RELATORIO ${rotulo}_PG_${mesAtual}.pptx`});
    btn.textContent = 'Exportar PPTX';
  } catch (e) {
    btn.textContent = 'Falhou — ver console';
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}
/* o pptxgen precisa da imagem em base64; o logo é do mesmo diretório, fetch resolve */
async function dataURI(url){
  try {
    const b = await (await fetch(url)).blob();
    return await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); });
  } catch { return null; }
}

/* ---- slide do PPTX ----
   Fundo, primitivas do layout e, no slide oculto, a marca de oculto do
   PowerPoint. Slide sem conteúdo leva o texto da pendência no lugar do cartão
   da tela (o botão de escrever não tem o que fazer num arquivo). */
function pptSlide(p, s, imgs){
  const sl = p.addSlide();
  sl.background = {color: LAYOUT.fundo(s)};
  if (s.oculto) sl.hidden = true;
  LAYOUT.pptx(p, sl, LAYOUT.prims(s, {linkVideo}), imgs);
  if (s.t === 'pendente') {
    const T = (t, o) => sl.addText(String(t || ''), Object.assign({fontFace:'Calibri', color:'666666', margin:[0,0,0,0]}, o));
    T('SEM CONTEÚDO AINDA', {x:0.5, y:1.6, w:9, h:0.3, fontSize:9, bold:true, color:'C09200', charSpacing:2});
    if (s.fonte) {
      T('BASE QUE VAI ALIMENTAR', {x:0.5, y:2.1, w:9, h:0.22, fontSize:8, charSpacing:1});
      T(s.fonte, {x:0.5, y:2.34, w:9, h:0.4, fontSize:10, color:'1A1A1A'});
      T('POR QUE AINDA NÃO TEM', {x:0.5, y:2.9, w:9, h:0.22, fontSize:8, charSpacing:1});
      T(s.motivo, {x:0.5, y:3.14, w:9, h:0.5, fontSize:10, color:'1A1A1A'});
    }
  }
}

/* ---- boot ---- */
document.getElementById('mes').innerHTML = SPEC.meses
  .map(m => `<option value="${m}">${esc(SPEC.labels[m])}</option>`).join('');
document.getElementById('mes').onchange = e => trocaMes(e.target.value);
document.getElementById('modo').onchange = e => {
  modo = e.target.value;
  idx = 0;                       // a numeração muda; voltar ao início evita cair fora
  trocaMes(mesAtual);
};
document.getElementById('prev').onclick = () => anda(-1);
document.getElementById('next').onclick = () => anda(1);
document.getElementById('ir').onchange = e => go(+e.target.value);
document.getElementById('play').onclick = play;
/* ---- versões do conteúdo ----
   O histórico é gravado por trigger no banco a cada salvar (ver a migration
   20260923120000): aqui é só a janela pra ver, cravar e restaurar. "Cravar"
   congela o estado atual com um rótulo — é o que se faz antes de levar o deck
   ao comitê; o rastro automático guarda o de antes de cada edição. */
async function listaVersoes(mes){
  const sb = hubSb();
  if (!sb) return [];
  const { data } = await sb.from('comite_conteudo_versao')
    .select('versao,cravada,rotulo,criado_por,criado_em')
    .eq('mes', mes).order('versao', {ascending: false});
  return data || [];
}

async function abreVersoes(){
  const sb = hubSb();
  if (!sb) return;
  let ov = document.getElementById('versoesOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'versoesOverlay';
    ov.innerHTML = `<div id="versoesPanel">
      <div class="ed-head"><h3 id="vsTitulo"></h3>
        <button type="button" id="vsFechar" aria-label="Fechar">✕</button></div>
      <div class="vs-cravar">
        <input id="vsRotulo" placeholder="Rótulo (ex.: levado ao comitê de setembro)">
        <button type="button" id="vsCravar" class="primary">Cravar versão de agora</button>
      </div>
      <div id="vsLista"></div>
    </div>`;
    document.body.appendChild(ov);
    ov.addEventListener('click', e => { if (e.target === ov) ov.style.display = 'none'; });
    document.getElementById('vsFechar').onclick = () => { ov.style.display = 'none'; };
    document.getElementById('vsCravar').onclick = async () => {
      const rot = document.getElementById('vsRotulo').value;
      const { error } = await sb.rpc('comite_cravar_versao', {p_mes: mesAtual, p_rotulo: rot});
      if (error) return alert('Não deu pra cravar: ' + error.message);
      document.getElementById('vsRotulo').value = '';
      pintaVersoes();
    };
  }
  ov.style.display = 'flex';
  document.getElementById('vsTitulo').textContent = `Versões — ${SPEC.labels[mesAtual]}`;
  pintaVersoes();
}

async function pintaVersoes(){
  const lista = document.getElementById('vsLista');
  lista.innerHTML = '<p class="vs-vazio">carregando…</p>';
  const vs = await listaVersoes(mesAtual);
  if (!vs.length) {
    lista.innerHTML = '<p class="vs-vazio">Nenhuma versão ainda. A primeira aparece no próximo salvar, '
      + 'ou agora, se você cravar uma.</p>';
    return;
  }
  const dt = t => { try { return new Date(t).toLocaleString('pt-BR', {dateStyle:'short', timeStyle:'short'}); }
                    catch(e){ return t; } };
  lista.innerHTML = vs.map(v => `<div class="vs-item${v.cravada ? ' vs-cravada' : ''}">
    <div class="vs-meta">
      <b>v${v.versao}</b>${v.cravada ? ' <span class="tag ok">cravada</span>' : ''}
      <span class="vs-quando">${dt(v.criado_em)}</span>
      <span class="vs-quem">${esc(v.criado_por || '')}</span>
    </div>
    ${v.rotulo ? `<div class="vs-rotulo">${esc(v.rotulo)}</div>` : ''}
    <button type="button" class="vs-voltar" data-restaura="${v.versao}">restaurar</button>
  </div>`).join('');
  lista.querySelectorAll('[data-restaura]').forEach(b => b.onclick = async () => {
    const v = +b.dataset.restaura;
    if (!confirm(`Restaurar a v${v} de ${SPEC.labels[mesAtual]}?

`
      + 'O conteúdo de agora não se perde: ele vira uma versão nova antes de ser substituído.')) return;
    const { error } = await hubSb().rpc('comite_restaurar_versao', {p_mes: mesAtual, p_versao: v});
    if (error) return alert('Não deu pra restaurar: ' + error.message);
    delete conteudoCache[mesAtual];
    await aplicaConteudoAoVivo(mesAtual);
    pintaVersoes();
  });
}

document.getElementById('versoes').onclick = abreVersoes;

/* Barra de atualizar: só pra quem edita o comitê — quem só assiste não dispara
   pipeline. Ao terminar, recarrega o spec do bucket em vez de mandar o usuário
   apertar F5: o deck inteiro é montado a partir dele. */
function montaBarraJob(){
  if (!window.HubJob || !souEditor) return;
  const slot = document.getElementById('barraJob');
  if (!slot || slot.dataset.pronto) return;
  slot.dataset.pronto = '1';
  window.HubJob.barra(slot, 'comite', {aoTerminar: () => location.reload()});
}

document.getElementById('pdf').onclick = exportarPdf;
document.getElementById('pptx').onclick = e => exportarPptx(e.currentTarget);

document.getElementById('editar').onclick = () => abreEditor(slides[idx]);
/* mesmo editor, chamado do cartao do slide vazio */
document.body.addEventListener('click', e => {
  if (e.target.closest('[data-abrir-editor]')) abreEditor(slides[idx]);
  const v = e.target.closest('[data-video]');
  if (v) { e.preventDefault(); abreVideo(v.dataset.video); }
});

/* ---- player do vídeo ----
   Dentro do deck o clique no frame toca aqui mesmo, sem sair do slide. Vindo
   do PPTX ou do PDF chega pelo `?video=`, com a página abrindo direto no
   player. Nos dois casos a URL assinada é pedida agora: a do arquivo exportado
   já teria vencido. */
async function abreVideo(path){
  if (!path) return;
  let ov = document.getElementById('videoOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'videoOverlay';
    ov.innerHTML = `<div id="videoBox">
      <button type="button" id="videoFechar" aria-label="Fechar">✕</button>
      <video id="videoPlayer" controls playsinline></video>
      <p id="videoErro" hidden></p>
    </div>`;
    document.body.appendChild(ov);
    ov.addEventListener('click', e => { if (e.target === ov) fechaVideo(); });
    document.getElementById('videoFechar').onclick = fechaVideo;
  }
  const player = document.getElementById('videoPlayer');
  const erro = document.getElementById('videoErro');
  ov.style.display = 'flex';
  player.hidden = true; erro.hidden = true;
  const url = await urlAssinada(path, 7200);
  if (!url) {
    erro.textContent = 'Não foi possível abrir o vídeo. Entre no hub e tente de novo.';
    erro.hidden = false;
    return;
  }
  player.src = url; player.hidden = false;
  player.play().catch(() => {});   // autoplay bloqueado: fica nos controles
}
function fechaVideo(){
  const ov = document.getElementById('videoOverlay');
  if (!ov) return;
  const player = document.getElementById('videoPlayer');
  if (player) { player.pause(); player.removeAttribute('src'); player.load(); }
  ov.style.display = 'none';
}
document.addEventListener('keydown', e => {
  const ov = document.getElementById('videoOverlay');
  if (e.key === 'Escape' && ov && ov.style.display !== 'none') { fechaVideo(); e.stopPropagation(); }
}, true);

const alvo = /^#([\d-]+)\/(\d+)$/.exec(location.hash);
if (alvo && SPEC.decks[alvo[1]]) { idx = +alvo[2] - 1; mesAtual = alvo[1]; }
trocaMes(mesAtual);
checaEditor();
/* `?video=<path>` — é para onde o frame do PPTX e do PDF aponta. O mês vem no
   próprio path (`2026-08/...`), então o deck abre no mês certo antes de tocar. */
const pedido = new URLSearchParams(location.search).get('video');
if (pedido) {
  const mesDoVideo = pedido.split('/')[0];
  if (SPEC.decks[mesDoVideo] && mesDoVideo !== mesAtual) trocaMes(mesDoVideo);
  abreVideo(pedido);
}
