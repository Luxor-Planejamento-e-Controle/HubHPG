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
/* Paleta do PPTX = a do relatório que a Ana levava ao comitê: fundo branco,
   faixa dourada no topo, cabeçalho de tabela navy com texto branco. O deck na
   tela usa as mesmas cores (ver .slide em deck.css), então a prévia é fiel ao
   arquivo que sai. */
const C = {bg:'FFFFFF', bg2:'FAFAFA', card:'F4F6F8', line:'E0E0E0', ink:'1A1A1A',
           ink2:'444444', ink3:'666666', amber:'C09200', pos:'1E7A46', neg:'C0392B',
           head:'0D2035', headInk:'FFFFFF'};
const LOGO = 'assets/pg-logo.png';
/* área útil do slide: 1280×720 menos cabeçalho (96) e rodapé (44) */
const BODY_H = 720 - 96 - 44;

/* ---- formatação ---- */
const nf = (v, d) => v.toLocaleString('pt-BR', {minimumFractionDigits:d, maximumFractionDigits:d});
const rs = v => v == null ? '—' : (v < 0 ? '-' : '') + 'R$ ' + nf(Math.abs(v), 0);
const dk = v => v == null ? '—' : (v >= 0 ? '+' : '−') + nf(Math.abs(v), 0) + 'k';
const dpct = v => v == null ? 'N/A' : (v >= 0 ? '+' : '−') + nf(Math.abs(v) * 100, 0) + '%';
const cls = v => v == null ? '' : v >= 0 ? 'pos' : 'neg';
const esc = s => String(s == null ? '' : s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const brdata = s => {
  if (!s) return '—';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s));
  return m ? `${m[3]}/${m[2]}/${m[1].slice(2)}` : String(s);
};
/* R$ k, como no deck original: R$18.703k */
function rsk(v){
  if (v == null || v === 0) return '—';
  const s = v < 0 ? '-' : '', a = Math.abs(v) / 1000;
  return s + 'R$' + nf(a, a < 100 ? 1 : 0) + 'k';
}
/* Altura de linha e corpo de fonte que fazem N linhas caberem em `alt`.
   Sem isso a tabela do DRE (40+ linhas) vazava por cima do rodapé — foi o que
   obrigou a diminuir o zoom do navegador pra conseguir ler. */
function ajusta(n, alt = BODY_H, maxH = 34){
  // Piso de 12px/9pt cortava o fim do resumo financeiro: são 48 linhas mais o
  // cabeçalho, e o relatório da Ana põe tudo num slide só. Com 9px/7,5pt cabe
  // inteiro e continua legível na projeção.
  const h = Math.max(9, Math.min(maxH, alt / Math.max(n, 1)));
  return {h, fs: Math.max(7.5, Math.min(16, h * 0.58))};
}

/* ---- render HTML ---- */
const head = s => `<div class="s-head"><h1>${esc(s.titulo)}</h1>${s.sub ? `<div class="sub">${esc(s.sub)}</div>` : ''}</div>
  <img class="s-mark" src="${LOGO}" alt="">`;
const foot = (s, i, n) => `<div class="s-foot"><span>${esc(SPEC.labels[mesAtual])}</span>
  ${s.obs ? `<span class="obs">${esc(s.obs)}</span>` : ''}<span class="spacer"></span><span>${i + 1} / ${n}</span></div>`;
const kpiRow = ks => `<div class="kpis" style="grid-template-columns:repeat(${ks.length},1fr)">` +
  ks.map(k => `<div class="kpi"><div class="v">${esc(k.v)}</div><div class="l">${esc(k.l)}</div><div class="s">${esc(k.s)}</div></div>`).join('') + `</div>`;

function tabelaHTML(cols, rows, fmtCel, alt, larguras){
  const {h, fs} = ajusta(rows.length + 1, alt);
  return `<table class="t" style="font-size:${fs.toFixed(1)}px">
    ${larguras ? `<colgroup>${larguras.map(w => `<col style="width:${w}">`).join('')}</colgroup>` : ''}
    <thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(r => {
      const cel = r.cells || r;
      const nv = r.nivel != null ? r.nivel : (r.total ? 0 : null);
      return `<tr class="${nv === 0 ? 'tot' : nv === 1 ? 'sub' : ''}" style="height:${h.toFixed(1)}px">` +
        cel.map((c, j) => fmtCel(c, j, r)).join('') + `</tr>`;
    }).join('')}</tbody></table>`;
}

const R = {
  capa: s => `<div class="capa"><img src="${LOGO}" alt="">
    <h1>${esc(s.titulo)}</h1><div class="mes">${esc(s.mes)}</div><div class="org">${esc(s.org)}</div></div>`,

  encerramento: s => `<div class="fim"><img src="${LOGO}" alt=""><h1>${esc(s.titulo)}</h1></div>`,

  divisor: s => `<div class="divisor"><img class="dv-logo" src="${LOGO}" alt="">
    <div class="n">0${s.n}</div>
    <h1>${esc(s.titulo)}</h1><div class="sub">${esc(s.sub)}</div></div>`,

  agenda: s => head(s) + `<div class="s-body"><div class="agenda">` +
    s.itens.map(it => `<div class="it"><div class="n">${esc(it.n)}</div>
      <h3>${esc(it.titulo)}</h3><p>${esc(it.sub)}</p></div>`).join('') + `</div></div>`,

  /* S04–S07, S10, S13, S14 — Orçado | Realizado | ∆ R$ k | ∆ %
     Três níveis: 0 grupo (dourado), 1 subgrupo (branco, recuado), 2 folha. */
  dre: s => head(s) + `<div class="s-body">` + tabelaHTML(
    ['NATUREZA', 'ORÇADO', 'REALIZADO', '∆ R$ k', '∆ %'],
    s.linhas.map(l => ({nivel: l.nivel == null ? (l.total ? 0 : 2) : l.nivel, cells: [l.nome, ...l.v]})),
    (c, j, r) => j === 0
      ? `<td class="nome n${r.nivel}">${esc(c)}</td>`
      : j === 3 ? `<td class="${cls(c)}">${dk(c)}</td>`
      : j === 4 ? `<td class="${cls(c)}">${dpct(c)}</td>`
      : `<td>${rs(c)}</td>`,
    // -26: `ajusta` reparte a altura igualmente entre as linhas, mas o cabeçalho
    // da tabela é mais alto que uma linha (padding + caixa alta). Sem descontar
    // isso, o resumo de 48 linhas estourava a moldura e a última —
    // "Resultado após Investimentos" — ficava fora do slide.
    BODY_H - 26, ['40%', '', '', '12%', '10%']) + `</div>`,

  /* KPIs + tabela (S11, S16–S20, S29, S37) */
  kpis_tabela: s => head(s) + `<div class="s-body">${kpiRow(s.kpis)}` + tabelaHTML(
    s.tabela.cols, s.tabela.rows,
    (c, j) => `<td${j === 0 ? ' class="nome"' : ''}>${esc(c)}</td>`,
    BODY_H - 108) + `</div>`,

  /* tabela pura (S30, S32–S35) */
  tabela: s => head(s) + `<div class="s-body">` + (s.rows.length
    ? tabelaHTML(s.cols, s.rows, (c, j) => {
        if ((s.moeda || []).includes(j)) return `<td>${rs(c)}</td>`;
        if ((s.data || []).includes(j)) return `<td>${brdata(c)}</td>`;
        return `<td${j === 0 ? ' class="nome"' : ''}>${esc(c)}</td>`;
      }, BODY_H)
    : vazio('○', 'nada neste recorte', 'Sem registro para o filtro deste slide',
            'A base foi lida e respondeu vazio — não é falta de fonte.')) + `</div>`,

  /* S12 — KPIs + matriz título × meses */
  matriz: s => head(s) + `<div class="s-body">${kpiRow(s.kpis)}` + tabelaHTML(
    s.cols, s.rows.map(r => ({total: /Saldo/.test(r[0]), cells: r})),
    (c, j) => j === 0 ? `<td class="nome">${esc(c)}</td>`
                      : `<td class="${typeof c === 'number' && c ? cls(c) : ''}">${c ? rsk(c) : '—'}</td>`,
    BODY_H - 108) + `</div>`,

  /* S09 — investimentos mês a mês */
  lista_mes: s => {
    const n = s.meses.reduce((a, m) => a + 1 + m.itens.length, 0);
    const {fs} = ajusta(n, BODY_H, 28);
    return head(s) + `<div class="s-body"><div class="lista" style="font-size:${fs.toFixed(1)}px">` +
      s.meses.map(m => `<div class="m"><span class="mes">${esc(m.mes)}</span>
          <span class="tag">Animais e produtos</span><span class="tot">${rs(m.total)}</span></div>` +
        m.itens.map(it => `<div class="it"><span></span><span class="d">${esc(it.desc)}</span><span>${rs(it.valor)}</span></div>`).join('')
      ).join('') + `</div></div>`;
  },

  /* S08 — comentários do DRE: categoria · texto · delta */
  comentarios: s => {
    const {h, fs} = ajusta(s.itens.length, BODY_H, 64);
    return head(s) + `<div class="s-body"><div class="coment">` + s.itens.map(i =>
      `<div class="li" style="min-height:${h.toFixed(0)}px;font-size:${Math.min(15, fs + 1.5).toFixed(1)}px">
        <div class="cat">${esc(i.cat)}</div>
        <div class="txt">${esc(i.txt)}</div>
        <div class="d ${/^[-−]/.test(i.delta) ? 'neg' : 'pos'}">${esc(i.delta)}</div>
      </div>`).join('') + `</div></div>`;
  },

  /* S24+ — resultados de exposição: animal e seus prêmios, em duas colunas.
     Poucos animais deixavam metade do slide vazia (texto sempre no mesmo
     tamanho fixo) — escala fonte/espaçamento pra preencher a altura do
     corpo, igual ao ajusta() das tabelas, só que por linha de texto em vez
     de linha de tabela. */
  resultados: s => {
    const linhas = s.animais.reduce((a, an) => a + 1 + an.premios.length, 0);
    const porColuna = Math.max(1, Math.ceil(linhas / 2));
    const h = Math.max(24, Math.min(54, BODY_H / porColuna));
    const fsNome = Math.max(13, Math.min(21, h * 0.42));
    const fsP = Math.max(12.5, Math.min(18, h * 0.36));
    const gap = Math.max(14, Math.min(32, h * 0.55));
    return head(s) + `<div class="s-body"><div class="premios" style="--rf-nome:${fsNome.toFixed(1)}px;--rf-p:${fsP.toFixed(1)}px;--rf-gap:${gap.toFixed(1)}px">` +
      s.animais.map(a => `<div class="an"><div class="nome">${esc(a.nome)}</div>` +
        a.premios.map(p => `<div class="p">${esc(p)}</div>`).join('') + `</div>`).join('') +
      `</div></div>`;
  },

  /* S38 — histórico de manejo, mês a mês */
  manejo: s => {
    const {h, fs} = ajusta(s.itens.length, BODY_H, 86);
    return head(s) + `<div class="s-body"><div class="manejo" style="font-size:${Math.min(15, fs + 1.5).toFixed(1)}px">` + s.itens.map(([m, t]) =>
      `<div class="li" style="min-height:${h.toFixed(0)}px"><div class="m">${esc(m)}</div>
        <div class="t">${esc(t)}</div></div>`).join('') + `</div></div>`;
  },

  /* S39+ — fotos do mês */
  // a grade vem do spec (s.grade = [colunas, linhas]) para bater com a do PPTX; o
  // último slide do mês raramente fecha com 6 fotos, e julho/26 tem uma só
  fotos: s => head(s) + `<div class="s-body"><div class="fotos" style="grid-template-columns:repeat(${(s.grade||[3])[0]},1fr)">` +
    // img ja vem como data URI: as fotos nao existem como arquivo no site (repo
    // e site sao publicos), vem embutidas no spec, que sai do bucket privado.
    // Video e um <a> de verdade, nao um <button>: no "Salvar como PDF" o Chrome
    // preserva href e o frame no PDF fica clicavel, que e o combinado.
    s.fotos.map(f => {
      const o = itemFoto(f);
      const bg = o.img ? ` style="background-image:url('${o.img}')"` : '';
      return o.video
        ? `<a class="f f-video${o.img ? '' : ' f-sem-capa'}" href="${escAttr(linkVideo(o.video))}"
             data-video="${escAttr(o.video)}"${bg}><span class="play" aria-hidden="true"></span></a>`
        : `<div class="f"${bg}></div>`;
    }).join('') +
    `</div></div>`,

  pendente: s => head(s) + `<div class="s-body">` + (editorDe(s)
    ? vazioEditavel(s)
    : vazio('◇', 'slide em aberto', s.titulo, s.sub || '', s.fonte, s.motivo)) + `</div>`,
};

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
const GRADE_FOTOS = {1:[1,1], 2:[2,1], 3:[3,1], 4:[2,2], 5:[3,2], 6:[3,2]};
const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const ABR_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const TIPOS_EDITAVEIS = new Set(['comentarios', 'manejo', 'fotos', 'resultados']);
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
async function montaSlidesAoVivo(mes){
  const c = await buscaConteudoAoVivo(mes);
  if (!c) return null;
  const [ano, mNum] = mes.split('-').map(Number);
  const out = {};

  if (c.comentarios && c.comentarios.length) {
    out.comentarios = [{t:'comentarios', n:8,
      titulo:`COMENTÁRIOS — VARIAÇÕES YTD JAN–${ABR_PT[mNum-1].toUpperCase()} ${ano}`,
      sub:'DRE 2026 | HPG · principais destaques acumulados por categoria',
      itens: c.comentarios}];
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
    out.manejo = [{t:'manejo', n:38, titulo:'MANEJO — PONTOS DE MELHORIA E DECISÕES', itens: c.manejo}];
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
        if (!ehVideo(a)) { const img = await fotoDataUri(a); return img ? {img, video: null} : null; }
        return {img: await fotoDataUri(posterDe(a)), video: a};
      }))).filter(Boolean);
      if (!itens.length) continue;
      const n = Math.ceil(itens.length / FOTOS_POR_SLIDE);
      for (let k = 0; k < n; k++) {
        const bloco = itens.slice(k * FOTOS_POR_SLIDE, (k + 1) * FOTOS_POR_SLIDE);
        const [cols, rows] = GRADE_FOTOS[bloco.length];
        let sub = g.tema ? `Obras e melhorias realizadas · ${g.tema}` : `Registros de ${MESES_PT[mNum-1]} ${ano}`;
        if (n > 1) sub += ` (${k+1}/${n})`;
        s.push({t:'fotos', n:39, titulo:'MANEJO — FOTOS E REGISTROS', sub, grade:[cols, rows], fotos:bloco});
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
  const body = (R[s.t] || R.pendente)(s);
  stage.innerHTML = `<div class="slide" id="slide">${body}${s.t === 'capa' || s.t === 'encerramento' ? '' : foot(s, idx, n)}</div>`;
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
    .map((s, i) => `<option value="${i}">${String(i + 1).padStart(2, '0')} · ${esc((s.titulo || s.t).slice(0, 52))}</option>`).join('');
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
  if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') { go(idx + 1); e.preventDefault(); }
  else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { go(idx - 1); e.preventDefault(); }
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
  caixa.innerHTML = slides.map((s, i) => {
    const corpo = (R[s.t] || R.pendente)(s);
    const rodape = s.t === 'capa' || s.t === 'encerramento' ? '' : foot(s, i, slides.length);
    return `<div class="slide">${corpo}${rodape}</div>`;
  }).join('');
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
    const logo = await dataURI(LOGO);
    // fotos precisam virar base64 antes: o pptxgen não busca arquivo sozinho
    // item JA carrega o data URI (fotos vem embutidas no spec, nao como
    // arquivo). Video entra pelo poster + hyperlink, montados no pptSlide.
    const imgs = {};
    for (const s of slides) for (const f of (s.fotos || [])) {
      const img = itemFoto(f).img;
      if (img && !(img in imgs)) imgs[img] = img;
    }
    slides.forEach((s, i) => pptSlide(p, s, i, logo, imgs));
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

/* ---- métricas de texto do PPTX ----
   O pptxgen não faz layout: ele só declara geometria e o PowerPoint
   recalcula na hora de mostrar. Duas consequências que estouravam o slide:
   linha de tabela CRESCE quando a célula quebra em várias linhas (rowH é
   mínimo, não altura), e caixa de texto NÃO cresce — o texto vaza por cima
   do que vier embaixo. Nos dois casos a altura precisa ser estimada aqui,
   antes de escrever, e a fonte tem que caber no que sobra até o rodapé.
   Segoe UI tem ~0,5em de largura média por caractere; conferido contra o
   render do PowerPoint, a conta erra pouco e pra cima — o lado seguro. */
const CHAR_EM = 0.5;
const LINHA_EM = 1.25;          // entrelinha que o PowerPoint aplica
const MARGEM_CEL = 0.08;        // polegadas comidas pelas laterais da célula de tabela
const MARGEM_TXT = 0.22;        // idem numa caixa de texto (inset padrão ~0,1in de cada lado)
const alturaLinha = fs => fs * LINHA_EM / 72;
/* quantas linhas o texto ocupa numa caixa de `larguraIn` polegadas */
function linhasTexto(txt, larguraIn, fs){
  const s = String(txt == null ? '' : txt).trim();
  if (!s) return 1;
  const cap = Math.max(1, Math.floor(larguraIn / (fs * CHAR_EM / 72)));
  let linhas = 1, usado = 0;
  for (const palavra of s.split(/\s+/)) {
    const espaco = usado ? 1 : 0;
    if (usado + espaco + palavra.length <= cap) { usado += espaco + palavra.length; continue; }
    if (usado) { linhas++; usado = 0; }                 // fecha a linha atual
    usado = palavra.length;
    while (usado > cap) { linhas++; usado -= cap; }     // palavra maior que a linha
  }
  return linhas;
}
/* altura de cada linha da tabela, respeitando fonte própria de célula (o
   cabeçalho tem a sua) */
function alturasTabela(rows, colW, fs){
  return rows.map(r => Math.max(...r.map((c, j) => {
    const f = (c && c.options && c.options.fontSize) || fs;
    const txt = c && c.text != null ? c.text : c;
    return linhasTexto(txt, (colW[j] || 9.2 / r.length) - MARGEM_CEL, f) * alturaLinha(f);
  })));
}
const somaAlturas = a => a.reduce((x, v) => x + v, 0);

function pptSlide(p, s, i, logo, imgs){
  const sl = p.addSlide();
  sl.background = {color: C.bg};
  // faixa dourada do topo — a marca do relatório da Ana, em todo slide
  sl.addShape(p.ShapeType.rect, {x:0, y:0, w:10, h:0.055, fill:{color:C.amber}, line:{width:0}});
  const T = (t, o) => sl.addText(t, Object.assign({fontFace:'Segoe UI', color:C.ink}, o));

  if (s.t === 'capa' || s.t === 'encerramento'){
    if (logo) sl.addImage({data:logo, x:6.1, y:1.1, w:2.6, h:2.6});
    T(s.titulo, {x:0.6, y:s.t === 'capa' ? 1.9 : 2.4, w:5.4, h:1.1, fontSize:s.t === 'capa' ? 28 : 26,
                 bold:true, color:s.t === 'capa' ? C.ink : C.amber});
    if (s.mes) T(s.mes, {x:0.6, y:3.0, w:5.4, h:0.4, fontSize:15, color:C.amber});
    if (s.org) T(s.org, {x:0.6, y:3.8, w:5.4, h:0.4, fontSize:10, color:C.ink3, charSpacing:3});
    return;
  }
  if (s.t === 'divisor'){
    T('0' + s.n, {x:6.6, y:0.4, w:2.9, h:1.4, fontSize:72, bold:true, color:C.line, align:'right'});
    T(s.titulo, {x:0.6, y:2.1, w:6, h:0.8, fontSize:26, bold:true});
    T(s.sub, {x:0.6, y:2.95, w:8.8, h:0.4, fontSize:12, color:C.ink3});
    return;
  }

  /* título comprido quebrava em duas linhas, cobria o subtítulo e a tabela — que
     começa em y fixo — não recuava. Reduz a fonte quando é longo, em vez de deixar
     quebrar. */
  const tituloLongo = (s.titulo || '').length > 58;
  T(s.titulo, {x:0.41, y:0.13, w:8.6, h:0.4, fontSize:tituloLongo ? 12 : 15,
               bold:true, valign:'top'});
  if (s.sub) T(s.sub, {x:0.41, y:0.52, w:8.6, h:0.3, fontSize:9, color:C.ink3});
  if (logo) sl.addImage({data:logo, x:9.05, y:0.14, w:0.55, h:0.55});
  T(`${SPEC.labels[mesAtual]}   ·   ${i + 1}/${slides.length}${s.obs ? '   ·   ' + s.obs : ''}`,
    {x:0.41, y:5.25, w:9.2, h:0.25, fontSize:8, color:C.ink3});

  /* Altura de tabela: rowH no pptxgen é MÍNIMO, o PowerPoint estica a linha pra
     caber o texto quebrado mais a margem interna da célula. Dividir a altura útil
     pelo nº de linhas (o que se fazia aqui) só acerta quando toda célula é de uma
     linha — com "25% FERNANDO SANTOS SILVEIRA / 25% YURI SEMANSKY ENGLER" numa
     coluna de 1,3in a célula vira 5 linhas e a tabela passa por cima do rodapé;
     era o estouro dos slides de vendas/embriões/DRE.
     Agora a altura de cada linha é ESTIMADA (linhasTexto por célula) e a fonte
     desce até o total caber em `alt` — o declarado bate com o renderizado.
     `margin` é por CÉLULA e a UNIDADE depende do primeiro valor: o pptxgen lê
     como PONTOS quando margin[0] >= 1 e como POLEGADAS quando é menor. Com
     margem vertical 0 cai-se sempre no ramo de polegadas — daí [0,.03,0,.03]:
     zero em cima e embaixo (era a margem vertical que fazia cada linha crescer
     ~3pt além do rowH) e 0,03in nas laterais, pouco o bastante pra não
     estrangular coluna estreita. Passar [0,2,0,2] achando que era ponto vira
     2 POLEGADAS de margem: a coluna fica menor que um caractere, o PowerPoint
     quebra a cada letra e a tabela explode pra 75in de altura. */
  const tbl = (rows, opts) => {
    const o = opts || {};
    const y = o.y || 0.95;
    const alt = 5.15 - y;
    const colW = o.colW || new Array(rows[0].length).fill(9.2 / rows[0].length);
    let fs = Math.min(9, o.fontSize || 9), alturas = alturasTabela(rows, colW, fs);
    while (fs > 5 && somaAlturas(alturas) > alt) {
      fs -= 0.25;
      alturas = alturasTabela(rows, colW, fs);
    }
    return sl.addTable(rows, Object.assign({
      x:0.41, y, w:9.2, border:{type:'solid', pt:0.4, color:C.line},
      fontFace:'Segoe UI', color:C.ink, valign:'middle',
    }, o, {fontSize:fs, rowH:alturas, margin:[0, 0.03, 0, 0.03]}));
  };
  const th = t => ({text:String(t), options:{bold:true, fontSize:8, color:C.headInk,
                                             fill:{color:C.head}, align:'right'}});
  const kpis = () => {
    const w = 9.2 / s.kpis.length;
    s.kpis.forEach((k, j) => {
      const x = 0.41 + j * w;
      sl.addShape(p.ShapeType.roundRect, {x, y:0.95, w:w - 0.14, h:0.82,
        fill:{color:C.card}, line:{color:C.line, width:0.5}, rectRadius:0.05});
      T(k.v, {x:x + 0.13, y:1.0, w:w - 0.4, h:0.34, fontSize:17, bold:true, color:C.amber});
      T(k.l, {x:x + 0.13, y:1.34, w:w - 0.4, h:0.22, fontSize:8, color:C.ink2});
      T(k.s, {x:x + 0.13, y:1.53, w:w - 0.4, h:0.2, fontSize:7, color:C.ink3});
    });
  };

  if (s.t === 'agenda'){
    s.itens.forEach((it, k) => {
      const x = 0.41 + k * 1.83;
      sl.addShape(p.ShapeType.rect, {x, y:1.15, w:1.7, h:0.04, fill:{color:C.amber}});
      T(it.n, {x, y:1.3, w:1.7, h:0.4, fontSize:18, bold:true, color:C.amber});
      T(it.titulo, {x, y:1.8, w:1.7, h:0.35, fontSize:10, bold:true});
      T(it.sub, {x, y:2.2, w:1.7, h:0.7, fontSize:8, color:C.ink3});
    });
    return;
  }
  if (s.t === 'dre'){
    const rows = [[th('NATUREZA'), th('ORÇADO'), th('REALIZADO'), th('∆ R$ k'), th('∆ %')]];
    s.linhas.forEach(l => rows.push([
      {text:(l.total ? '' : '   ') + l.nome, options:{align:'left', bold:l.total, color:l.total ? C.amber : C.ink2}},
      {text:rs(l.v[0]), options:{align:'right', bold:l.total}},
      {text:rs(l.v[1]), options:{align:'right', bold:l.total}},
      {text:dk(l.v[2]), options:{align:'right', color:l.v[2] == null ? C.ink : l.v[2] >= 0 ? C.pos : C.neg}},
      {text:dpct(l.v[3]), options:{align:'right', color:l.v[3] == null ? C.ink3 : l.v[3] >= 0 ? C.pos : C.neg}},
    ]));
    tbl(rows, {colW:[3.7, 1.7, 1.7, 1.1, 1.0]});
    return;
  }
  if (s.t === 'tabela'){
    if (!s.rows.length){ T('Nenhum registro nesse recorte.', {x:0.41, y:2.4, w:8, h:0.4, fontSize:12, color:C.ink3}); return; }
    const rows = [s.cols.map(th)];
    s.rows.forEach(r => rows.push(r.map((c, j) => ({
      text:(s.moeda || []).includes(j) ? rs(c) : (s.data || []).includes(j) ? brdata(c) : String(c == null ? '—' : c),
      options:{align:j === 0 ? 'left' : 'right'}}))));
    const w0 = 9.2 / s.cols.length;
    tbl(rows, {colW:[w0 * 1.6].concat(new Array(s.cols.length - 1).fill((9.2 - w0 * 1.6) / (s.cols.length - 1)))});
    return;
  }
  if (s.t === 'kpis_tabela' || s.t === 'matriz'){
    kpis();
    if (s.t === 'kpis_tabela'){
      const rows = [s.tabela.cols.map(th)];
      s.tabela.rows.forEach(r => rows.push(r.map((c, j) =>
        ({text:String(c), options:{align:j === 0 ? 'left' : 'right'}}))));
      const w0 = 9.2 / s.tabela.cols.length;
      tbl(rows, {y:1.95, colW:[w0 * 1.7].concat(new Array(s.tabela.cols.length - 1)
        .fill((9.2 - w0 * 1.7) / (s.tabela.cols.length - 1)))});
    } else {
      const rows = [s.cols.map(th)];
      s.rows.forEach(r => rows.push(r.map((c, j) => j === 0
        ? {text:String(c), options:{align:'left', bold:/Saldo/.test(c), color:/Saldo/.test(c) ? C.amber : C.ink2}}
        : {text:c ? rsk(c) : '—', options:{align:'right', color:typeof c === 'number' && c < 0 ? C.neg : C.ink}})));
      tbl(rows, {y:1.95, colW:[1.9].concat(new Array(s.cols.length - 1).fill((9.2 - 1.9) / (s.cols.length - 1)))});
    }
    return;
  }
  if (s.t === 'lista_mes'){
    /* passo fixo por item não serve: descrição comprida quebra em 2+ linhas,
       a caixa de texto não cresce e o texto invade o item seguinte — no fim
       da lista o acumulado ia por cima do rodapé. Cada item anda a altura
       que ele realmente ocupa, e a fonte desce se o total não couber. */
    const alt = 5.15 - 0.95, largDesc = 6.1;
    const planeja = fs => {
      const lh = alturaLinha(fs), itens = [];
      s.meses.forEach((m, k) => {
        itens.push({m, gap: k ? lh * 0.55 : 0, h: lh * 1.5});   // respiro entre grupos de mês
        m.itens.forEach(it => itens.push({it, gap: 0, h: linhasTexto(it.desc, largDesc - MARGEM_TXT, fs) * lh}));
      });
      return {itens, total: somaAlturas(itens.map(x => x.h + x.gap))};
    };
    let fs = 9, plano = planeja(fs);
    while (fs > 5 && plano.total > alt) { fs -= 0.25; plano = planeja(fs); }
    let y = 0.95;
    plano.itens.forEach(l => {
      y += l.gap;
      if (l.m) {
        T(l.m.mes, {x:0.41, y, w:0.9, h:l.h, fontSize:fs, bold:true, color:C.amber, valign:'top'});
        T('ANIMAIS E PRODUTOS', {x:1.35, y, w:3, h:l.h, fontSize:fs - 1.5, color:C.ink3, valign:'top'});
        T(rs(l.m.total), {x:7.6, y, w:2.0, h:l.h, fontSize:fs + 1, bold:true, align:'right', valign:'top'});
      } else {
        T(l.it.desc, {x:1.35, y, w:largDesc, h:l.h, fontSize:fs - 1, color:C.ink2, valign:'top'});
        T(rs(l.it.valor), {x:7.6, y, w:2.0, h:l.h, fontSize:fs - 1, align:'right', color:C.ink2, valign:'top'});
      }
      y += l.h;
    });
    return;
  }
  if (s.t === 'comentarios'){
    const rows = [[th('CATEGORIA'), th('DESTAQUE DO ACUMULADO'), th('∆')]];
    s.itens.forEach(it => rows.push([
      {text:it.cat, options:{align:'left', bold:true, color:C.amber}},
      {text:it.txt, options:{align:'left', color:C.ink2}},
      {text:it.delta, options:{align:'right', bold:true, color:/^[-−]/.test(it.delta) ? C.neg : C.pos}},
    ]));
    tbl(rows, {colW:[1.9, 6.3, 1.0], fontSize:8, valign:'top'});
    return;
  }
  if (s.t === 'resultados'){
    /* duas colunas. Passo fixo de 0,24in por prêmio ignorava que prêmio
       comprido quebra em 2 linhas numa coluna de 4,3in, e o corte entre as
       colunas era por altura estimada errada (0,24 fixo) — a segunda metade
       descia por cima do rodapé. Altura real por bloco, corte pelo meio
       dessa altura, e fonte que desce se a coluna mais alta não couber. */
    const alt = 5.15 - 0.95, largCol = 4.4;
    const planeja = fs => {
      const lh = alturaLinha(fs);
      const blocos = s.animais.map(a => {
        const premios = a.premios.map(pr => ({txt:pr, h: linhasTexto('🏆  ' + pr, largCol - 0.14 - MARGEM_TXT, fs) * lh}));
        const nomeH = lh * 1.3;
        return {nome:a.nome, nomeH, premios, h: nomeH + somaAlturas(premios.map(x => x.h)) + lh * 0.35};
      });
      return {blocos, total: somaAlturas(blocos.map(b => b.h))};
    };
    const reparte = plano => {
      const meta = plano.total / 2, cols = [[], []];
      let acc = 0;
      plano.blocos.forEach(b => { if (acc < meta) { cols[0].push(b); acc += b.h; } else cols[1].push(b); });
      return cols;
    };
    const maiorCol = cols => Math.max(...cols.map(c => somaAlturas(c.map(b => b.h))));
    let fs = 10, plano = planeja(fs), cols = reparte(plano);
    while (fs > 6 && maiorCol(cols) > alt) { fs -= 0.5; plano = planeja(fs); cols = reparte(plano); }
    cols.forEach((lista, k) => {
      let y = 0.95;
      lista.forEach(b => {
        T(b.nome, {x:0.41 + k * 4.7, y, w:largCol, h:b.nomeH, fontSize:fs + 1, bold:true, color:C.amber, valign:'top'});
        y += b.nomeH;
        b.premios.forEach(pr => {
          T('🏆  ' + pr.txt, {x:0.55 + k * 4.7, y, w:largCol - 0.14, h:pr.h, fontSize:fs, color:C.ink, valign:'top'});
          y += pr.h;
        });
        y += alturaLinha(fs) * 0.35;
      });
    });
    return;
  }
  if (s.t === 'manejo'){
    const rows = [[th('MÊS'), th('INTERVENÇÕES E DECISÕES')]];
    s.itens.forEach(([m, t]) => rows.push([
      {text:m, options:{align:'left', bold:true, color:C.amber}},
      {text:t, options:{align:'left', color:C.ink}},
    ]));
    tbl(rows, {colW:[0.9, 8.3], fontSize:8.5, valign:'top'});
    return;
  }
  if (s.t === 'fotos'){
    // mesma grade do HTML, vinda do spec — não recalcular aqui, senão as duas saídas
    // divergem quando o mês fecha com menos de 6 fotos
    const [cols, linhas] = s.grade || [3, 2];
    const gap = 0.12;
    const w = (9.2 - gap * (cols - 1)) / cols;
    const alt = (4.15 - gap * (linhas - 1)) / linhas;
    s.fotos.forEach((f, k) => {
      const o = itemFoto(f);
      const x = 0.41 + (k % cols) * (w + gap), y = 0.95 + Math.floor(k / cols) * (alt + gap);
      /* Video vira o frame com hyperlink pro hub — o arquivo NAO e embutido:
         um .pptx com os videos dentro passa de 100 MB e nao sai por e-mail.
         Sem poster (formato que o navegador de quem subiu nao decodificou)
         entra um retangulo com o rotulo, pra ninguem achar que sumiu. */
      const link = o.video ? {hyperlink:{url: linkVideo(o.video), tooltip:'Abrir o vídeo no hub'}} : {};
      const d = o.img && imgs ? imgs[o.img] : null;
      if (d) {
        sl.addImage({data:d, x, y, w, h:alt, sizing:{type:'cover', w, h:alt}, ...link});
      } else if (o.video) {
        sl.addShape(p.ShapeType.roundRect, {x, y, w, h:alt, fill:{color:C.card}, line:{color:C.line, width:1}, rectRadius:0.06});
        T('▶ VÍDEO', {x, y: y + alt / 2 - 0.15, w, h:0.3, fontSize:11, bold:true, color:C.amber, align:'center', ...link});
      }
      if (o.video && d) {
        T('▶', {x, y, w:0.34, h:0.28, fontSize:12, bold:true, color:'FFFFFF', align:'center', ...link});
      }
    });
    return;
  }
  // pendente
  T('SEM DADO AINDA', {x:0.41, y:1.6, w:3, h:0.3, fontSize:9, bold:true, color:C.amber, charSpacing:2});
  T('BASE QUE VAI ALIMENTAR', {x:0.41, y:2.2, w:8.6, h:0.22, fontSize:8, color:C.ink3, charSpacing:1});
  T(s.fonte, {x:0.41, y:2.44, w:8.6, h:0.5, fontSize:10});
  T('POR QUE AINDA NÃO TEM', {x:0.41, y:3.1, w:8.6, h:0.22, fontSize:8, color:C.ink3, charSpacing:1});
  T(s.motivo, {x:0.41, y:3.34, w:8.6, h:0.5, fontSize:10});
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
document.getElementById('prev').onclick = () => go(idx - 1);
document.getElementById('next').onclick = () => go(idx + 1);
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
