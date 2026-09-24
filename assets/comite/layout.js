/* Desenho dos slides do comitê — a geometria do relatório da Ana.

   O relatório que o haras conhece (RELATORIO_MENSAL_PG_JULHO26) não usa tabela
   de PowerPoint: cada slide é um punhado de retângulos e caixas de texto em
   posições fixas. Este arquivo reproduz ISSO — cada tipo de slide vira uma lista
   de primitivas (retângulo, texto, imagem, arco) em coordenadas de 1280×720,
   tiradas do próprio arquivo dela (x, y, largura, altura, corpo de fonte, cor) —
   e a MESMA lista é desenhada no HTML e no PPTX. Assim a tela, o PDF e o
   PowerPoint não têm como divergir entre si nem do modelo.

   Unidades: px num slide de 1280×720 (10×5,625 in → 128 px/in); fonte em pt,
   como no PowerPoint (1 pt = 128/72 px). */
'use strict';

(function () {
  const PX_IN = 128;
  const PT = PX_IN / 72;
  const FONTE = 'Calibri';
  const COR = {
    navy: '0D2035', azul: '1A3A5C', azul2: '2E5F9A', ouro: 'C09200', ouro2: 'D4A820',
    ardosia: '445566', vinho: '7B1313', vermelho: '9B2020', verde: '27AE60', neg: 'C0392B',
    tinta: '1A1A1A', cinza: '666666', cinzaClaro: 'BBBBBB', linha: 'E0E0E0',
    banda: 'E5EAF2', zebra: 'F5F7FA', trilho: 'E8EDF2', metaBar: 'C8D4E8', destaque: 'FFF8E1',
  };
  const LOGO_OURO = 'assets/comite/logo-ouro.png';
  const LOGO_NAVY = 'assets/comite/logo-navy.png';

  /* ---------------- medida de texto (Calibri) ----------------
     Largura de avanço de cada caractere, em unidades de 2048 por em. O
     PowerPoint não encolhe texto sozinho — caixa que não cabe vaza por cima do
     vizinho —, então o layout mede antes de escrever e reduz o corpo ou abrevia.
     A tabela é da Calibri Regular; negrito é ~3,5% mais largo. */
  const LARG = {' ':463,'!':544,'"':821,'#':1038,'$':1038,'%':1464,'&':1397,"'":452,'(':621,')':621,
    '*':1038,'+':1038,',':511,'-':627,'.':517,'/':791,':':548,';':548,'<':1038,'=':1038,'>':1038,
    '?':950,'@':1823,'A':1185,'B':1114,'C':1092,'D':1260,'E':1000,'F':941,'G':1292,'H':1276,'I':516,
    'J':653,'K':1064,'L':861,'M':1751,'N':1322,'O':1356,'P':1058,'Q':1378,'R':1112,'S':941,'T':998,
    'U':1314,'V':1162,'W':1822,'X':1063,'Y':998,'Z':959,'a':981,'b':1076,'c':866,'d':1076,'e':1019,
    'f':625,'g':964,'h':1076,'i':470,'j':490,'k':931,'l':470,'m':1636,'n':1076,'o':1080,'p':1076,
    'q':1076,'r':714,'s':801,'t':686,'u':1076,'v':925,'w':1464,'x':887,'y':927,'z':809,
    '—':1843,'–':1024,'·':517,'×':1038,'∆':1170,'▲':1300,'▼':1300,'✓':1400,'…':1400,'º':700,'ª':700,
    '|':940,'_':1024,'[':621,']':621};
  const semAcento = c => c.normalize('NFD').replace(/[̀-ͯ]/g, '');
  function largura(txt, pt, negrito) {
    let u = 0;
    for (const ch of String(txt == null ? '' : txt)) {
      let w = LARG[ch];
      if (w == null) w = LARG[semAcento(ch)];
      if (w == null) w = /\d/.test(ch) ? 1038 : (ch.codePointAt(0) > 0x2000 ? 2048 : 1000);
      u += w;
    }
    return u / 2048 * pt * PT * (negrito ? 1.035 : 1);
  }
  /* o maior corpo (até `pt`, sem passar de `minimo`) que faz o texto caber numa
     linha de `w` px; se nem no mínimo cabe, corta com reticência */
  function encaixa(txt, pt, w, negrito, minimo) {
    const t = String(txt == null ? '' : txt), alvo = w * 0.97;
    let p = pt;
    const piso = minimo || pt;
    while (p > piso && largura(t, p, negrito) > alvo) p = Math.max(piso, p - 0.25);
    if (largura(t, p, negrito) <= alvo) return {t, pt: p};
    let s = t;
    while (s.length > 1 && largura(s + '…', p, negrito) > alvo) s = s.slice(0, -1);
    return {t: s.trimEnd() + '…', pt: p};
  }
  /* quantas linhas o texto ocupa em `w` px, quebrando por palavra */
  function nLinhas(txt, pt, w, negrito) {
    const palavras = String(txt || '').split(/\s+/).filter(Boolean);
    if (!palavras.length) return 1;
    const esp = largura(' ', pt, negrito);
    let n = 1, usado = 0;
    for (const p of palavras) {
      const lp = largura(p, pt, negrito);
      if (!usado) { usado = lp; continue; }
      if (usado + esp + lp <= w * 0.97) usado += esp + lp;
      else { n++; usado = lp; }
    }
    return n;
  }
  const altLinha = pt => pt * PT * 1.2;

  /* ---------------- primitivas ---------------- */
  const R = (x, y, w, h, o) => Object.assign({k: 'r', x, y, w, h}, o || {});
  const T = (x, y, w, h, t, o) => Object.assign({k: 't', x, y, w, h, t: String(t == null ? '' : t)}, o || {});
  const I = (x, y, w, h, src, o) => Object.assign({k: 'i', x, y, w, h, src}, o || {});
  const A = (cx, cy, r, esp, a0, a1, fill) => ({k: 'a', cx, cy, r, esp, a0, a1, fill});
  const H = html => ({k: 'h', html});
  /* célula de tabela: texto de uma linha que já sai medido */
  function C(x, y, w, h, t, o) {
    const q = Object.assign({pt: 8, al: 'l'}, o || {});
    const f = encaixa(t, q.pt, w, q.b, q.min || Math.max(6, q.pt - 1.5));
    return T(x, y, w, h, f.t, Object.assign({}, q, {pt: f.pt}));
  }

  /* ---------------- formatação (a do relatório) ---------------- */
  const nf = (v, d) => Number(v).toLocaleString('pt-BR', {minimumFractionDigits: d, maximumFractionDigits: d});
  /* arredonda como a planilha: meio pra longe do zero. Math.round leva o meio
     pra cima, e aí -43.937,5 saía -43.937 contra os -43.938 da face */
  const arred = v => Math.sign(v) * Math.round(Math.abs(v));
  const reais = v => v == null ? '—' : (v < 0 ? '-' : '') + 'R$ ' + nf(Math.abs(arred(v)), 0);
  const deltaK = v => {
    if (v == null) return '—';
    const r = arred(v);
    return r === 0 ? '0k' : (r > 0 ? '+' : '-') + nf(Math.abs(r), 0) + 'k';
  };
  const deltaP = v => {
    if (v == null) return 'N/A';
    // o relatório escreve -2409%, sem separador de milhar no percentual
    const r = arred(v * 100);
    return r === 0 ? '0%' : (r > 0 ? '+' : '-') + Math.abs(r) + '%';
  };
  /* R$18.703k / R$7,5k / -R$1.091k — a matriz da movimentação */
  function reaisK(v) {
    if (!v) return '—';
    const a = Math.abs(v) / 1000;
    const r1 = Math.round(a * 10) / 10;
    const s = a >= 1000 ? nf(Math.round(a), 0) : (r1 % 1 === 0 ? nf(r1, 0) : nf(r1, 1));
    return (v < 0 ? '-' : '') + 'R$' + s + 'k';
  }
  const milhoes = v => 'R$ ' + nf(v / 1e6, 1) + 'M';
  const corDelta = (v, neutra) => v == null ? neutra : (arred(v) === 0 ? COR.cinza : (v > 0 ? COR.verde : COR.neg));

  /* ---------------- moldura ---------------- */
  function claro(s, P) {
    P.push(R(0, 0, 1280, 9, {fill: COR.ouro}));
    // o logo navy: 1180,29 → 58×57 (a imagem dela tem folga; esta é recortada)
    P.push(I(1180.4, 29.1, 58.6, 57.1, LOGO_NAVY, {logo: 1}));
    const t = encaixa(s.titulo || '', 15, 1069, true, 11);
    P.push(T(64, 19.2, 1069, 61.4, t.t, {pt: t.pt, b: 1, c: COR.navy, va: 'm'}));
    if (s.sub) {
      const u = encaixa(s.sub, 8.5, 1069, false, 7);
      P.push(T(64, 80.6, 1069, 30.7, u.t, {pt: u.pt, c: COR.cinza, va: 'm'}));
    }
    return P;
  }
  const CORES_CARD = {
    navy: [COR.navy, COR.ouro2, COR.cinzaClaro], ouro: [COR.ouro, COR.navy, COR.azul],
    azul: [COR.azul, COR.ouro2, COR.cinzaClaro], ardosia: [COR.ardosia, COR.ouro2, COR.cinzaClaro],
    vinho: [COR.vinho, 'FFAAAA', 'FFAAAA'], vermelho: [COR.vermelho, 'FFAAAA', 'FFAAAA'],
  };
  const ORDEM_CARD = ['navy', 'azul', 'ouro', 'vinho'];
  /* cartões de indicador: blocos sólidos, valor branco centrado, rótulo em
     negrito e a linha de apoio embaixo */
  function cartoes(P, ks, h, ptValor) {
    const n = ks.length, gap = 7.7, w = (1190.4 - gap * (n - 1)) / n;
    ks.forEach((k, j) => {
      const x = 44.8 + j * (w + gap), y = 134.4;
      const [fundo, cl, cs] = CORES_CARD[k.cor || ORDEM_CARD[j % 4]] || CORES_CARD.navy;
      P.push(R(x, y, w, h, {fill: fundo}));
      const pv = k.pt || ptValor;
      const v = encaixa(k.v, pv, w - 24, true, 12);
      P.push(T(x, y + 7.7, w, h * 0.5, v.t, {pt: v.pt, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
      P.push(C(x + 8, y + h * 0.522, w - 16, h * 0.28, k.l, {pt: 8.5, b: 1, c: cl, al: 'c', va: 'm', min: 7}));
      if (k.s) P.push(C(x + 8, y + h * 0.782, w - 16, h * 0.2, k.s, {pt: 7.5, c: cs, al: 'c', va: 'm', min: 6.5}));
    });
  }
  function cabecalho(P, y, h, cols, pt) {
    P.push(R(44.8, y, 1190.4, h, {fill: COR.navy}));
    cols.forEach(([x, w, al, t]) => { if (t) P.push(C(x, y + 3.8, w, h - 7.6, t, {pt: pt || 8, b: 1, c: 'FFFFFF', al, va: 'm'})); });
  }
  /* passo de linha que faz `n` linhas caberem entre `y0` e `yMax`, sem passar do
     passo do relatório — slide com poucas linhas não estica */
  const passo = (n, y0, yMax, max) => Math.min(max, (yMax - y0) / Math.max(n, 1));

  const L = {};

  /* ---------------- capa, agenda, divisor, encerramento ---------------- */
  L.capa = s => [
    R(0, 0, 1280, 57.6, {fill: COR.ouro}),
    I(514.7, 140, 251, 244.7, LOGO_OURO, {logo: 1}),
    T(64, 476.2, 1152, 48.6, s.titulo, {pt: 12, c: COR.ouro, al: 'c', va: 'm', sp: 2}),
    T(64, 535, 1152, 87, s.mes, {pt: 34, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}),
    T(64, 633.6, 1152, 35.8, s.org, {pt: 9, c: COR.ouro2, al: 'c', va: 'm'}),
  ];
  L.agenda = s => {
    const P = [R(0, 0, 1280, 57.6, {fill: COR.ouro}),
      I(62.6, 89, 66.9, 65.3, LOGO_OURO, {logo: 1}),
      T(179.2, 70.4, 1011, 70.4, s.titulo, {pt: 24, b: 1, c: 'FFFFFF', va: 'm'}),
      T(179.2, 138.2, 1011, 35.8, s.sub, {pt: 9.5, c: COR.ouro2, va: 'm'})];
    s.itens.forEach((it, k) => {
      const x = 44.8 + k * 243.2;
      P.push(R(117.8 + k * 243.2, 198.4, 97.3, 97.3, {fill: COR.ouro, forma: 'elipse'}));
      P.push(T(117.8 + k * 243.2, 198.4, 97.3, 97.3, it.n, {pt: 15, b: 1, c: COR.navy, al: 'c', va: 'm'}));
      P.push(T(x, 314.9, 240.6, 53.8, it.titulo, {pt: 9, b: 1, c: COR.ouro, al: 'c', va: 'm'}));
      P.push(T(x, 368.6, 240.6, 53.8, it.sub, {pt: 8, c: COR.cinzaClaro, al: 'c', va: 'm'}));
      if (k < s.itens.length - 1) P.push(R(285.4 + k * 243.2, 204.8, 2.6, 243.2, {fill: '334466'}));
    });
    return P;
  };
  L.divisor = s => [
    R(0, 0, 57.6, 720, {fill: COR.ouro}),
    I(108.5, 58.1, 71.1, 69.3, LOGO_OURO, {logo: 1}),
    T(960, 38.4, 281.6, 281.6, '0' + s.n, {pt: 110, b: 1, c: COR.ouro, alfa: 0.35, al: 'r', va: 'm'}),
    T(102.4, 211.2, 1139, 105, s.titulo, {pt: 36, b: 1, c: 'FFFFFF', va: 'm'}),
    T(102.4, 326.4, 1139, 48.6, s.sub, {pt: 12, c: COR.ouro2, va: 'm'}),
  ];
  L.encerramento = s => [
    R(0, 0, 1280, 57.6, {fill: COR.ouro}),
    I(523, 154.6, 234.3, 228.4, LOGO_OURO, {logo: 1}),
    T(64, 480, 1152, 66.6, s.titulo, {pt: 24, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}),
    T(64, 556.8, 1152, 44.8, s.sub || '', {pt: 12, c: COR.ouro2, al: 'c', va: 'm'}),
  ];

  /* ---------------- pendências da apresentação anterior ---------------- */
  L.pendencias = s => {
    const P = claro(Object.assign({}, s, {sub: ''}), []);
    let pt = 18;
    const wTxt = 939;
    const altura = p => s.itens.reduce((a, t) => a + nLinhas(t, p, wTxt) * altLinha(p), 0);
    while (pt > 11 && 154 + altura(pt) > 690) pt -= 0.5;
    let y = 154;
    s.itens.forEach(t => {
      const h = nLinhas(t, pt, wTxt) * altLinha(pt);
      P.push(T(163, y, 34, altLinha(pt), '✓', {pt: pt * 0.8, c: COR.navy, va: 'm'}));
      P.push(T(203, y, wTxt, h, t, {pt, c: COR.navy, va: 't', wrap: 1}));
      y += h;
    });
    return P;
  };

  /* ---------------- DRE: resumo, caixa, casa e análises ---------------- */
  const COLS_DRE = {
    resumo:  [[48.6, 460.8, 'l'], [512, 194.6, 'r'], [709.1, 194.6, 'r'], [906.2, 128, 'r'], [1036.8, 105, 'r']],
    caixa:   [[48.6, 512, 'l'], [563.2, 198.4, 'r'], [764.2, 198.4, 'r'], [965.1, 128, 'r'], [1095.7, 92.2, 'r']],
    casa:    [[48.6, 512, 'l'], [563.2, 194.6, 'r'], [760.3, 194.6, 'r'], [957.4, 128, 'r'], [1088, 99.8, 'r']],
    analise: [[48.6, 576, 'l'], [627.2, 172.8, 'r'], [802.6, 172.8, 'r'], [977.9, 128, 'r'], [1108.5, 92.2, 'r']],
  };
  /* pesos de linha do relatório: grupo em preto, detalhe cinza recuado,
     resultado em faixa azul-clara e negrito, a última invertida em navy; na
     análise, subgrupo em azul negrito e folha com recuo maior */
  const ESTILO = {
    grupo:  {pt: 8, c: COR.tinta},
    det:    {pt: 7.5, c: COR.cinza, rec: '  '},
    folha:  {pt: 7.5, c: COR.cinza, rec: '    '},
    sub:    {pt: 7.5, b: 1, c: '002060', rec: '  '},
    banda:  {pt: 8, b: 1, c: COR.navy, fill: COR.banda},
    total:  {pt: 8, b: 1, c: COR.navy, fill: COR.banda},
    rotulo: {pt: 8, b: 1, c: COR.navy, fill: COR.banda},
    fim:    {pt: 8, b: 1, c: 'FFFFFF', fill: COR.navy},
  };
  L.dre = s => {
    const P = claro(s, []);
    const C5 = COLS_DRE[s.layout] || COLS_DRE.resumo;
    const nomes = ['NATUREZA', 'ORÇADO', 'REALIZADO', '∆ R$ k', '∆ %'];
    P.push(R(44.8, 124.2, 1190.4, 35.8, {fill: COR.navy}));
    C5.forEach(([x, w, al], j) => P.push(T(x, 128, w, 28.2, nomes[j], {pt: 8, b: 1, c: 'FFFFFF', al, va: 'm'})));
    const n = s.linhas.length;
    const p = passo(n, 160, 712, s.layout === 'caixa' ? 32.6 : 31.8);
    const esc = p < 26 ? p / 26 : 1;           // linha espremida: fonte acompanha
    s.linhas.forEach((l, i) => {
      const e = ESTILO[l.estilo] || ESTILO.grupo;
      const y = 160 + i * p;
      P.push(R(44.8, y, 1190.4, p, {fill: e.fill || (i % 2 ? COR.zebra : 'FFFFFF'), line: COR.linha}));
      const ty = y + p * 0.13, th = p * 0.74, pt = e.pt * Math.max(0.85, esc);
      const o = {pt, b: e.b, c: e.c, va: 'm'};
      P.push(C(C5[0][0], ty, C5[0][1], th, (e.rec || '') + l.nome, Object.assign({}, o, {al: 'l'})));
      if (l.estilo === 'rotulo') return;
      const [orc, real, dk, pc] = l.v;
      P.push(C(C5[1][0], ty, C5[1][1], th, reais(orc), Object.assign({}, o, {al: 'r'})));
      P.push(C(C5[2][0], ty, C5[2][1], th, reais(real), Object.assign({}, o, {al: 'r'})));
      const cd = corDelta(dk, e.c);
      P.push(C(C5[3][0], ty, C5[3][1], th, deltaK(dk), Object.assign({}, o, {al: 'r', c: cd})));
      const cp = pc == null ? cd : corDelta(pc * 1000, e.c);
      P.push(C(C5[4][0], ty, C5[4][1], th, deltaP(pc), Object.assign({}, o, {al: 'r', c: cp})));
    });
    return P;
  };

  /* ---------------- investimentos, mês a mês ---------------- */
  L.lista_mes = s => {
    const P = claro(s, []);
    const nb = s.meses.length, ni = s.meses.reduce((a, m) => a + m.itens.length, 0);
    const natural = nb * 31.2 + ni * 28.9 + (nb - 1) * 3.3;
    const f = Math.min(1, (712 - 124.2) / natural);
    const hb = 31.2 * f, hi = 28.9 * f, gap = 3.3 * f, esc = Math.max(0.82, f);
    let y = 124.2;
    s.meses.forEach((m, k) => {
      if (k) y += gap;
      const atual = m.atual;
      P.push(R(44.8, y, 1190.4, hb, {fill: atual ? COR.ouro : COR.navy}));
      P.push(C(53.8, y, 153.6, hb, m.rotulo || m.mes, {pt: 8.5 * esc, b: 1, c: atual ? COR.navy : 'FFFFFF', va: 'm'}));
      P.push(C(211.2, y, 704, hb, 'ANIMAIS E PRODUTOS', {pt: 8.5 * esc, c: atual ? COR.navy : COR.ouro2, va: 'm'}));
      P.push(C(998.4, y, 230.4, hb, reais(m.total), {pt: 8.5 * esc, b: 1, c: atual ? COR.navy : 'FFFFFF', al: 'r', va: 'm'}));
      y += hb;
      m.itens.forEach((it, j) => {
        P.push(R(44.8, y, 1190.4, hi, {fill: j % 2 ? 'FFFFFF' : COR.zebra, line: COR.linha}));
        P.push(C(70.4, y, 900, hi, it.desc, {pt: 7.5 * esc, c: COR.tinta, va: 'm', min: 6}));
        P.push(C(998.4, y, 230.4, hi, reais(it.valor), {pt: 7.5 * esc, c: COR.cinza, al: 'r', va: 'm'}));
        y += hi;
      });
    });
    return P;
  };

  /* ---------------- investimentos do mês, por bloco ----------------
     A geometria do slide de agosto/2026 do relatório: faixa azul por bloco
     (nome e total em dourado), linhas favorecido | descrição | valor e a faixa
     navy do total. Mês com muito lançamento: o passo encolhe junto. */
  L.investimentos = s => {
    const P = claro(s, []);
    const Y0 = 115.2, HB = 34.6, HI = 35.8, PI = 38.4, GB = 6.4, GT = 9, HT = 43.5;
    const ni = s.blocos.reduce((a, b) => a + b.itens.length, 0), nb = s.blocos.length;
    const natural = nb * HB + ni * PI + (nb - 1) * GB - (PI - HI) + GT + HT;
    const f = Math.min(1, (712 - Y0) / natural), esc = Math.max(0.8, f);
    let y = Y0;
    s.blocos.forEach((b, k) => {
      if (k) y += GB * f;
      P.push(R(44.8, y, 1187.8, HB * f, {fill: COR.azul}));
      P.push(C(53.8, y + 3.8 * f, 768, 26.9 * f, b.nome, {pt: 8.5 * esc, b: 1, c: 'FFFFFF', va: 'm'}));
      P.push(C(896, y + 3.8 * f, 332.8, 26.9 * f, reais(b.total), {pt: 8.5 * esc, b: 1, c: COR.ouro, al: 'r', va: 'm'}));
      y += HB * f;
      b.itens.forEach((it, j) => {
        const h = HI * f;
        P.push(R(44.8, y, 1187.8, h, {fill: j % 2 ? COR.zebra : 'FFFFFF'}));
        P.push(C(53.8, y + 3.8 * f, 358.4, h - 7.6 * f, it.quem, {pt: 8 * esc, c: COR.tinta, va: 'm', min: 6.5}));
        // descrição comprida quebra em duas linhas antes de ser abreviada
        const d = it.desc || '';
        let pd = 8 * esc;
        if (nLinhas(d, pd, 537.6) > 1) pd = 7.25 * esc;
        if (nLinhas(d, pd, 537.6) > 1 && nLinhas(d, pd, 537.6) <= 2 && h >= 2 * altLinha(pd) + 2)
          P.push(T(416, y + 1, 537.6, h - 2, d, {pt: pd, c: COR.cinza, va: 'm', wrap: 1}));
        else
          P.push(C(416, y + 3.8 * f, 537.6, h - 7.6 * f, d, {pt: pd, c: COR.cinza, va: 'm', min: 6}));
        P.push(C(957.4, y + 3.8 * f, 268.8, h - 7.6 * f, reais(it.valor), {pt: 8 * esc, b: 1, c: COR.navy, al: 'r', va: 'm'}));
        y += PI * f;
      });
    });
    y += (GT - (PI - HI)) * f;
    P.push(R(44.8, y, 1187.8, HT * f, {fill: COR.navy}));
    P.push(C(53.8, y + 7.7 * f, 768, 28.2 * f, s.rotulo_total, {pt: 9 * esc, b: 1, c: 'FFFFFF', va: 'm'}));
    P.push(C(896, y + 7.7 * f, 332.8, 28.2 * f, reais(s.total), {pt: 11 * esc, b: 1, c: COR.ouro, al: 'r', va: 'm'}));
    return P;
  };

  /* ---------------- estoque em equinos ---------------- */
  const CORES_BARRA = [COR.navy, COR.azul, COR.azul2, COR.ouro, COR.ouro2, '4A7FC0', '6A9FD0', 'AABBCC'];
  L.estoque = s => {
    const P = claro(s, []);
    cartoes(P, s.kpis, 117.8, 28);
    P.push(R(96.1, 284.5, 1087.7, 34.6, {fill: COR.navy}));
    P.push(T(105.1, 288.3, 384, 26.9, 'CATEGORIA', {pt: 8.5, b: 1, c: 'FFFFFF', va: 'm'}));
    P.push(T(492.9, 288.3, 83.2, 26.9, 'Nº', {pt: 8.5, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
    P.push(T(578.7, 288.3, 70.4, 26.9, '%', {pt: 8.5, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
    const tot = s.rows.reduce((a, r) => a + r[1], 0) || 1;
    const p = passo(s.rows.length, 319.1, 712, 39);
    s.rows.forEach((r, i) => {
      const y = 319.1 + i * p, hh = p - 1.2;
      P.push(R(96.1, y, 1087.7, hh, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: 'DEDEDE'}));
      P.push(C(105.1, y, 384, hh, r[0], {pt: 8.5, c: COR.tinta, va: 'm'}));
      P.push(T(492.9, y, 83.2, hh, r[1], {pt: 9, b: 1, c: COR.navy, al: 'c', va: 'm'}));
      P.push(T(578.7, y, 70.4, hh, r[2], {pt: 8.5, c: COR.cinza, al: 'c', va: 'm'}));
      const bh = Math.min(16.6, hh * 0.44), by = y + (hh - bh) / 2;
      P.push(R(652.9, by, 512, bh, {fill: COR.trilho}));
      P.push(R(652.9, by, Math.max(2, 512 * r[1] / tot), bh, {fill: CORES_BARRA[i % CORES_BARRA.length]}));
    });
    return P;
  };

  /* ---------------- resumo da movimentação ---------------- */
  L.movimentacao = s => {
    const P = claro(s, []);
    cartoes(P, s.kpis, 108.8, 20);
    const nm = s.cols.length - 1;
    const cw = Math.min(129.5, (1230 - 291.8) / Math.max(nm, 1));
    P.push(R(44.8, 263.7, 1190.4, 34.6, {fill: COR.navy}));
    P.push(T(53.8, 267.5, 243.2, 26.9, s.cols[0], {pt: 8, b: 1, c: 'FFFFFF', va: 'm'}));
    s.cols.slice(1).forEach((c, j) => P.push(T(291.8 + j * cw, 267.5, cw, 26.9, c, {pt: 8, b: 1, c: 'FFFFFF', al: 'c', va: 'm'})));
    const p = passo(s.rows.length, 298.2, 712, 46.1);
    s.rows.forEach((r, i) => {
      const y = 298.2 + i * p, hh = p - 1.3;
      const ini = i === 0, fim = i === s.rows.length - 1;
      P.push(R(44.8, y, 1190.4, hh, {fill: fim ? COR.navy : ini ? COR.banda : (i % 2 ? COR.zebra : 'FFFFFF'), line: COR.linha}));
      const o = {pt: 8, b: ini || fim ? 1 : 0, c: fim ? 'FFFFFF' : COR.tinta, va: 'm'};
      P.push(C(53.8, y, 236.8, hh, r[0], o));
      r.slice(1).forEach((v, j) => P.push(C(291.8 + j * cw, y, cw, hh, reaisK(v), Object.assign({}, o, {al: 'c'}))));
    });
    return P;
  };
  L.matriz = L.movimentacao;

  /* ---------------- funil de prenhez ---------------- */
  L.funil = s => {
    const P = claro(s, []);
    cartoes(P, s.kpis, 153.6, 36);
    P.push(R(44.8, 313.6, 1187.8, 34.6, {fill: COR.navy}));
    P.push(T(53.8, 317.4, 691.2, 26.9, s.cab, {pt: 9, b: 1, c: 'FFFFFF', va: 'm'}));
    P.push(T(748.8, 317.4, 153.6, 26.9, 'Nº', {pt: 9, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
    P.push(T(906.2, 317.4, 361, 26.9, 'REFERÊNCIA', {pt: 9, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
    const p = passo(s.rows.length, 348.2, 712, 45.24);
    s.rows.forEach((r, i) => {
      const y = 348.2 + i * p, hh = p - 2.5;
      const aborto = /Aborto/.test(r[0]), conf = i === s.rows.length - 1;
      P.push(R(44.8, y, 1187.8, hh, {fill: conf ? COR.navy : aborto ? 'FFF0F0' : (i % 2 ? COR.zebra : 'FFFFFF'), line: 'DEDEDE'}));
      P.push(T(53.8, y, 691.2, hh, r[0], {pt: 9.5, b: conf ? 1 : 0, c: conf ? 'FFFFFF' : aborto ? COR.neg : COR.tinta, va: 'm'}));
      P.push(T(748.8, y, 153.6, hh, r[1], {pt: 12, b: 1, c: conf ? COR.ouro2 : aborto ? COR.neg : COR.navy, al: 'c', va: 'm'}));
      P.push(T(906.2, y, 361, hh, r[2], {pt: 9, c: conf ? COR.cinzaClaro : COR.cinza, al: 'c', va: 'm'}));
    });
    return P;
  };

  /* ---------------- garanhões ---------------- */
  L.garanhoes = s => {
    const P = claro(s, []);
    if (s.kpis.length) cartoes(P, s.kpis, 108.8, 24);
    const y0 = s.kpis.length ? 263.7 : 124.2;
    cabecalho(P, y0, 34.6, [[53.8, 345.6, 'l', 'GARANHÃO'], [403.2, 35.8, 'c', 'T'], [442.9, 70.4, 'c', 'LAV'],
      [517.1, 70.4, 'c', 'CONF'], [591.4, 79.4, 'c', 'ÍND%'], [674.6, 448, 'l', 'LAVADOS vs CONFIRMADOS']], 7.5);
    const maxLav = Math.max(1, ...s.rows.map(r => r.lav));
    const r0 = y0 + 34.5;
    const p = passo(s.rows.length, r0, 682, 34.84);
    s.rows.forEach((r, i) => {
      const y = r0 + i * p, hh = p - 2.5;
      P.push(R(44.8, y, 1187.8, hh, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: COR.linha}));
      P.push(C(53.8, y, 345.6, hh, r.nome, {pt: 7.5, c: COR.tinta, va: 'm'}));
      P.push(T(403.2, y, 35.8, hh, r.t || '—', {pt: 7.5, b: 1, c: COR.azul, al: 'c', va: 'm'}));
      P.push(T(442.9, y, 70.4, hh, r.lav, {pt: 8, c: COR.cinza, al: 'c', va: 'm'}));
      P.push(T(517.1, y, 70.4, hh, r.conf, {pt: 8, b: 1, c: COR.navy, al: 'c', va: 'm'}));
      const ind = r.lav ? r.conf / r.lav : 0;
      const ci = ind >= 0.6 ? COR.verde : ind >= 0.4 ? COR.ouro : COR.neg;
      P.push(T(591.4, y, 79.4, hh, Math.round(ind * 100) + '%', {pt: 8, b: 1, c: ci, al: 'c', va: 'm'}));
      const bh = Math.min(24.4, hh * 0.76), by = y + (hh - bh) / 2;
      P.push(R(674.6, by, 448, bh, {fill: COR.trilho}));
      P.push(R(674.6, by, Math.max(1, 448 * r.lav / maxLav), bh / 2, {fill: COR.metaBar}));
      if (r.conf) P.push(R(674.6, by + bh / 2, Math.max(1, 448 * r.conf / maxLav), bh / 2, {fill: COR.azul}));
    });
    const yl = Math.min(687.6, r0 + s.rows.length * p + 6);
    P.push(R(674.6, yl, 35.8, 17.9, {fill: COR.metaBar}));
    P.push(T(715.5, yl, 192, 17.9, 'Lavados (+)', {pt: 7.5, c: COR.cinza, va: 'm'}));
    P.push(R(905, yl, 35.8, 17.9, {fill: COR.azul}));
    P.push(T(945.9, yl, 192, 17.9, 'Confirmados', {pt: 7.5, c: COR.cinza, va: 'm'}));
    return P;
  };

  /* ---------------- comparativo entre safras ---------------- */
  const CORES_SAFRA = ['557799', COR.azul, COR.azul2];
  L.comparativo = s => {
    const P = claro(s, []);
    P.push(R(44.8, 124.2, 1187.8, 34.6, {fill: COR.navy}));
    P.push(T(53.8, 128, 249.6, 26.9, 'ESTAÇÃO', {pt: 8, b: 1, c: 'FFFFFF', va: 'm'}));
    const cx = j => 304.6 + j * 75.52;
    s.meses.forEach((m, j) => P.push(T(cx(j), 128, 75.5, 26.9, m, {pt: 7.5, b: 1, c: 'FFFFFF', al: 'c', va: 'm'})));
    const max = Math.max(1, ...s.safras.flatMap(b => b.meses.map(v => v || 0)));
    const escala = 33 / max;
    const velhas = s.safras.filter(b => !b.atual).length;
    s.safras.forEach((b, k) => {
      const y = 158.7 + k * 98.6, atual = b.atual;
      const cor = atual ? COR.ouro : CORES_SAFRA[Math.max(0, CORES_SAFRA.length - velhas + k)] || COR.azul;
      P.push(R(44.8, y, 1187.8, 48.6, {fill: atual ? COR.destaque : (k % 2 ? COR.zebra : 'FFFFFF'), line: COR.linha}));
      P.push(C(53.8, y + 11.5, 249.6, 25.6, b.rotulo, {pt: 7.5, b: atual ? 1 : 0, c: atual ? COR.navy : COR.tinta, va: 'm'}));
      b.meses.forEach((v, j) => {
        if (!v) {
          P.push(T(cx(j), y + 11.5, 75.5, 25.6, '-', {pt: 8, b: atual ? 1 : 0, c: COR.cinzaClaro, al: 'c', va: 'm'}));
          return;
        }
        const h = Math.max(2.6, v * escala), topo = y + 48.6 - h;
        P.push(R(cx(j) + 15.1, topo, 45.3, h, {fill: cor}));
        P.push(T(cx(j), topo - 21.3, 75.5, 21.3, v, {pt: 8, b: atual ? 1 : 0, c: COR.tinta, al: 'c', va: 'b'}));
      });
      P.push(T(1210.9, y + 11.5, 69, 25.6, b.total, {pt: 8, b: 1, c: atual ? COR.ouro : COR.navy, al: 'c', va: 'm'}));
      P.push(R(44.8, y + 48.7, 1187.8, 35.8, {fill: 'F0F0F0', line: 'EEEEEE'}));
      P.push(T(53.8, y + 53.8, 1164.8, 25.6, b.meta, {pt: 7.5, i: 1, c: COR.cinza, va: 'm'}));
    });
    return P;
  };

  /* ---------------- doadoras: meta × realizado ---------------- */
  L.doadoras = s => {
    const P = claro(s, []);
    cabecalho(P, 124.2, 34.6, [[53.8, 371.2, 'l', 'DOADORA'], [428.8, 53.8, 'c', 'META'], [485.1, 53.8, 'c', 'REAL'],
      [542.7, 537.6, 'l', 'META vs REALIZADO'], [1136.6, 87, 'c', '%']]);
    const max = Math.max(1, ...s.rows.map(r => Math.max(r.meta, r.real)));
    const p = passo(s.rows.length, 158.7, 686, 33);
    s.rows.forEach((r, i) => {
      const y = 158.7 + i * p, hh = p - 2.6;
      P.push(R(44.8, y, 1187.8, hh, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: COR.linha}));
      P.push(C(53.8, y, 371.2, hh, r.nome, {pt: 7.5, c: COR.tinta, va: 'm'}));
      P.push(T(428.8, y, 53.8, hh, r.meta, {pt: 7.5, c: COR.cinza, al: 'c', va: 'm'}));
      const atingiu = r.real >= r.meta && r.real > 0;
      const cr = !r.real ? COR.neg : atingiu ? COR.verde : COR.ouro;
      P.push(T(485.1, y, 53.8, hh, r.real, {pt: 8, b: 1, c: cr, al: 'c', va: 'm'}));
      const bh = Math.min(13.2, hh * 0.46), by = y + (hh - bh) / 2;
      P.push(R(542.7, by, 537.6, bh, {fill: COR.trilho}));
      if (r.meta) P.push(R(542.7, by, 537.6 * r.meta / max, bh, {fill: COR.metaBar}));
      if (r.real) P.push(R(542.7, by, 537.6 * r.real / max, bh * 0.65, {fill: COR.ouro}));
      const pct = !r.real ? '—' : r.meta ? Math.round(r.real / r.meta * 100) + '%' : '999%';
      P.push(T(1136.6, y, 87, hh, pct, {pt: 8, b: 1, c: !r.real ? COR.cinzaClaro : cr, al: 'c', va: 'm'}));
    });
    const yl = Math.min(692.5, 158.7 + s.rows.length * p + 6);
    [[542.7, COR.metaBar, 'Meta'], [762.9, COR.ouro, 'Realizado (abaixo)'], [983, COR.verde, 'Meta atingida']].forEach(([x, c, t]) => {
      P.push(R(x, yl, 35.8, 19.2, {fill: c}));
      P.push(T(x + 41, yl, 179.2, 19.2, t, {pt: 7, c: COR.cinza, va: 'm'}));
    });
    return P;
  };

  /* ---------------- coberturas disponíveis ---------------- */
  L.coberturas = s => {
    const P = claro(s, []);
    cabecalho(P, 124.2, 34.6, [[53.8, 742.4, 'l', 'GARANHÃO'], [800, 309.8, 'l', 'COBERTURAS DISPONÍVEIS'], [1152, 79.4, 'c', 'SAL.']]);
    const max = Math.max(10, ...s.rows.map(r => r.saldo));
    const p = passo(s.rows.length, 158.7, 712, 36.5);
    s.rows.forEach((r, i) => {
      const y = 158.7 + i * p, hh = p - 2.6;
      P.push(R(44.8, y, 1187.8, hh, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: COR.linha}));
      P.push(C(53.8, y, 742.4, hh, r.nome, {pt: 8.5, c: COR.tinta, va: 'm'}));
      const bh = Math.min(20.1, hh * 0.6), by = y + (hh - bh) / 2;
      P.push(R(800, by, 309.8, bh, {fill: COR.trilho}));
      P.push(R(800, by, 309.8 * r.saldo / max, bh, {fill: COR.azul}));
      P.push(T(1152, y, 79.4, hh, r.saldo, {pt: 9.5, b: 1, c: COR.navy, al: 'c', va: 'm'}));
    });
    return P;
  };

  /* ---------------- exposições: programação ---------------- */
  function programacao(s) {
    const P = claro(s, []);
    P.push(R(44.8, 128, 1190.4, 34.6, {fill: COR.navy}));
    [[53.8, 614.4, 'l', 'EVENTO'], [672, 230.4, 'c', 'DATA'], [905, 172.8, 'c', 'LOCAL'], [1081.6, 147.2, 'c', 'STATUS']]
      .forEach(([x, w, al, t]) => P.push(T(x, 131.8, w, 26.9, t, {pt: 8.5, b: 1, c: 'FFFFFF', al, va: 'm'})));
    const n = s.rows.length;
    const p = passo(n, 163.8, 712, 53.8);
    s.rows.forEach((r, i) => {
      const y = 163.8 + i * p, hh = p - 1.3, ult = i === n - 1;
      P.push(R(44.8, y, 1190.4, hh, {fill: ult ? COR.destaque : (i % 2 ? 'FFFFFF' : COR.zebra), line: COR.linha}));
      const [ev, dt, lo, st] = r.map(x => (x == null || x === '') ? '—' : x);
      P.push(C(53.8, y, 614.4, hh, ev, {pt: 9, b: ult ? 1 : 0, c: COR.tinta, va: 'm'}));
      P.push(C(672, y, 230.4, hh, dt, {pt: 9, c: COR.cinza, al: 'c', va: 'm'}));
      P.push(C(905, y, 172.8, hh, lo, {pt: 9, c: COR.cinza, al: 'c', va: 'm'}));
      const realizada = /realiz/i.test(st);
      P.push(C(1081.6, y, 147.2, hh, st, {pt: 9, b: ult ? 1 : 0, c: ult ? COR.ouro : realizada ? COR.verde : COR.cinza, al: 'c', va: 'm'}));
    });
    return P;
  }

  /* ---------------- resultados de exposição ---------------- */
  L.resultados = s => {
    const P = claro(s, []);
    /* o mesmo animal pode vir em várias entradas (um prêmio por linha, como o
       haras escreve no editor); o relatório junta tudo num cartão só */
    const junta = [];
    (s.animais || []).forEach(a => {
      const nome = String(a.nome || '').trim(), ja = junta.find(x => x.nome.toUpperCase() === nome.toUpperCase());
      if (ja) ja.premios = ja.premios.concat(a.premios || []);
      else junta.push({nome, premios: (a.premios || []).slice()});
    });
    s = Object.assign({}, s, {animais: junta});
    const cartao = a => 31.7 + a.premios.length * 39.7 + 5.9;
    const total = s.animais.reduce((x, a) => x + cartao(a), 0);
    const cols = [[], []];
    let acc = 0;
    s.animais.forEach(a => { (acc < total / 2 ? cols[0] : cols[1]).push(a); acc += cartao(a); });
    const alto = Math.max(...cols.map(c => c.reduce((x, a) => x + cartao(a), 0)));
    const f = Math.min(1, (712 - 128) / Math.max(alto, 1));
    cols.forEach((lista, k) => {
      const x = k ? 640 : 44.8;
      let y = 128;
      lista.forEach((a, j) => {
        const hc = 31.7 * f;
        P.push(R(x, y, 582.4, hc, {fill: j % 2 ? COR.azul : COR.navy}));
        P.push(C(x + 9, y, 569.6, hc, String(a.nome || '').toUpperCase(), {pt: 10 * Math.max(0.85, f), b: 1, c: 'FFFFFF', va: 'm'}));
        y += hc;
        a.premios.forEach((pr, i) => {
          const hp = 39.7 * f;
          P.push(R(x, y, 582.4, hp, {fill: i % 2 ? 'FFFFFF' : COR.zebra, line: COR.linha}));
          P.push(C(x + 12.8, y, 563.2, hp, '🏆  ' + pr, {pt: 9 * Math.max(0.85, f), c: COR.navy, va: 'm', min: 6.5}));
          y += hp;
        });
        y += 5.9 * f;
      });
    });
    return P;
  };

  /* ---------------- vendas: resultado acumulado ---------------- */
  L.vendas_acum = s => {
    const P = claro(s, []);
    cartoes(P, s.kpis, 105, 16);
    P.push(R(44.8, 256, 1187.8, 41, {fill: 'E0E0E0', line: 'CCCCCC'}));
    if (s.pct > 0) P.push(R(44.8, 256, 1187.8 * Math.min(1, s.pct), 41, {fill: COR.ouro}));
    P.push(T(51.2, 257.3, 1177.6, 38.4, s.barra, {pt: 10, b: 1, c: COR.tinta, al: 'c', va: 'm'}));
    const n = s.colunas.length;
    if (!n) return P;
    const slot = 1187.8 / Math.max(n, 7), bw = slot * 0.646;
    const x0 = 44.8 + (1187.8 - slot * n) / 2;
    const max = Math.max(1, ...s.colunas.map(c => c.v));
    P.push(R(41, 631.6, 1187.8, 2.6, {fill: COR.cinzaClaro}));
    s.colunas.forEach((c, j) => {
      const x = x0 + j * slot + (slot - bw) / 2, h = Math.max(2, 280 * c.v / max);
      P.push(R(x, 631.6 - h, bw, h, {fill: COR.azul}));
      P.push(T(x - 28.8, 631.6 - h - 32, bw + 57.6, 28.2, milhoes(c.v), {pt: 8, b: 1, c: COR.azul, al: 'c', va: 'm'}));
      P.push(T(x, 638, bw, 28.2, c.rot, {pt: 9, b: 1, c: COR.tinta, al: 'c', va: 'm'}));
    });
    return P;
  };

  /* ---------------- vendas: detalhamento por mês e evento ---------------- */
  L.vendas_mes = s => {
    const P = claro(s, []);
    cabecalho(P, 124.2, 34.6, [[53.8, 172.8, 'l', 'MÊS'], [230.4, 569.6, 'l', 'EVENTO / ORIGEM'],
      [803.8, 204.8, 'r', 'VALOR'], [1011.2, 217.6, 'l', 'VISUAL']]);
    const nb = s.meses.length, ne = s.meses.reduce((a, m) => a + m.eventos.length, 0);
    const f = Math.min(1, (714 - 158.7) / (nb * 27.8 + ne * 32.5));
    const hb = 27.8 * f, he = 32.5 * f, esc = Math.max(0.85, f);
    const max = Math.max(1, ...s.meses.map(m => m.total));
    let y = 158.7;
    s.meses.forEach(m => {
      P.push(R(44.8, y, 1187.8, hb, {fill: COR.banda, line: 'D0D8E8'}));
      P.push(C(53.8, y, 172.8, hb, m.mes, {pt: 9 * esc, b: 1, c: COR.navy, va: 'm'}));
      P.push(C(803.8, y, 204.8, hb, 'Total: ' + reais(m.total), {pt: 9 * esc, b: 1, c: COR.navy, al: 'r', va: 'm'}));
      const bh = Math.min(17.9, hb * 0.64), by = y + (hb - bh) / 2;
      P.push(R(1011.2, by, 211.2, bh, {fill: COR.trilho}));
      P.push(R(1011.2, by, Math.max(1, 211.2 * m.total / max), bh, {fill: COR.azul}));
      y += hb;
      m.eventos.forEach(([ev, v]) => {
        P.push(R(44.8, y, 1187.8, he, {fill: 'FFFFFF', line: 'EEEEEE'}));
        P.push(C(230.4, y, 569.6, he, ev, {pt: 8.5 * esc, c: COR.cinza, va: 'm'}));
        P.push(C(803.8, y, 204.8, he, reais(v), {pt: 8.5 * esc, c: COR.tinta, al: 'r', va: 'm'}));
        y += he;
      });
    });
    return P;
  };

  /* ---------------- inadimplência (o painel de cobrança) ---------------- */
  const PAINEL = {fundo: '10303A', card: '133845', borda: '1F4A58', txt: 'FFFFFF', apoio: '9FB6BE',
                  teal: '346C78', vermelho: 'F80000', laranja: 'F86404', amarelo: 'F4BC00', cinza: 'D8D8DC',
                  sobeRuim: 'FF6B6B', desceBom: '5CD6A0'};
  const brl2 = v => 'R$ ' + nf(v, 2);
  const pct1 = v => nf(v, 1).replace(/,0$/, '') + '%';
  L.inadimplencia = s => {
    const P = claro(s, []);
    const k = s.k, a = s.ant;
    P.push(R(0, 141, 1280, 535, {fill: PAINEL.fundo}));
    const delta = (v, va, ruimSubir) => {
      if (!a || va == null || !va) return null;
      const d = (v - va) / Math.abs(va) * 100, sobe = d >= 0;
      return {t: (sobe ? '▲ ' : '▼ ') + pct1(Math.abs(d)) + ' vs ' + s.ref_ant,
              c: ruimSubir == null ? PAINEL.apoio : ((sobe === ruimSubir) ? PAINEL.sobeRuim : PAINEL.desceBom)};
    };
    const cards = [
      ['Total em Aberto', k.total, PAINEL.txt, `${k.clientes} cliente(s) · ticket médio ${brl2(k.clientes ? k.total / k.clientes : 0)}`, delta(k.total, a && a.total, null)],
      ['A Vencer', k.avencer, PAINEL.txt, `${pct1(k.total ? k.avencer / k.total * 100 : 0)} do total`, delta(k.avencer, a && a.avencer, false)],
      ['Vencidos', k.venc, PAINEL.txt, `${pct1(k.total ? k.venc / k.total * 100 : 0)} do total · ${k.clientes_venc} cliente(s)`, delta(k.venc, a && a.venc, true)],
      ['Inadimplentes', k.inad, 'FF3B3B', `${pct1(k.total ? k.inad / k.total * 100 : 0)} do total`, delta(k.inad, a && a.inad, true)],
      ['Ação Judicial', k.aj, 'FF8A3D', `${pct1(k.total ? k.aj / k.total * 100 : 0)} do total`, delta(k.aj, a && a.aj, true)],
      ['Não Entregues', k.ne, 'FFC928', `${pct1(k.total ? k.ne / k.total * 100 : 0)} do total`, delta(k.ne, a && a.ne, true)],
    ];
    cards.forEach(([rot, v, cv, apoio, d], j) => {
      const x = 16 + (j % 3) * 419, y = 153 + Math.floor(j / 3) * 86, w = 409, h = 78;
      P.push(R(x, y, w, h, {fill: PAINEL.card, line: PAINEL.borda}));
      P.push(T(x + 14, y + 6, w - 28, 18, rot, {pt: 7.5, c: PAINEL.apoio, va: 'm'}));
      P.push(T(x + 14, y + 22, w - 28, 30, brl2(v), {pt: 14, b: 1, c: cv, va: 'm'}));
      P.push(C(x + 14, y + 52, (w - 28) * 0.62, 18, apoio, {pt: 6.5, c: PAINEL.apoio, va: 'm', min: 5.5}));
      if (d) P.push(C(x + 14 + (w - 28) * 0.62, y + 52, (w - 28) * 0.38, 18, d.t, {pt: 6.5, c: d.c, va: 'm', al: 'r', min: 5.5}));
    });
    // rosca: composição do total em aberto
    P.push(R(16, 330, 460, 334, {fill: PAINEL.card, line: PAINEL.borda}));
    P.push(T(30, 338, 430, 24, 'Composição do Total em Aberto', {pt: 8, b: 1, c: PAINEL.txt, va: 'm'}));
    const fatias = [[k.avencer, PAINEL.teal, 'A Vencer'], [k.inad, PAINEL.vermelho, 'Inadimplentes'],
                    [k.aj, PAINEL.laranja, 'Ação Judicial'], [k.ne, PAINEL.amarelo, 'Não Entregues']];
    const soma = fatias.reduce((x, f) => x + Math.max(0, f[0]), 0) || 1;
    let ang = 0;
    fatias.forEach(([v, c]) => {
      const d = Math.max(0, v) / soma * 360;
      if (d > 0.3) P.push(A(160, 505, 118, 44, ang, ang + d, c));
      ang += d;
    });
    fatias.forEach(([, c, t], j) => {
      P.push(R(318, 470 + j * 22, 11, 11, {fill: c, forma: 'elipse'}));
      P.push(T(335, 465 + j * 22, 130, 21, t, {pt: 7, b: 1, c: PAINEL.txt, va: 'm'}));
    });
    // barras: índice por ano de emissão
    P.push(R(486, 330, 778, 334, {fill: PAINEL.card, line: PAINEL.borda}));
    P.push(T(500, 338, 520, 24, 'Índice de Inadimplência por Ano de Emissão', {pt: 8, b: 1, c: PAINEL.txt, va: 'm'}));
    P.push(T(1040, 338, 210, 24, '% do total emitido', {pt: 6.5, c: PAINEL.apoio, al: 'r', va: 'm'}));
    const series = [['venc', PAINEL.cinza, 'Vencido/Total'], ['inad', PAINEL.vermelho, 'Inadimplente/Total'],
                    ['aj', PAINEL.laranja, 'Ação Judicial/Total'], ['ne', PAINEL.amarelo, 'Não Entregue/Total']];
    series.forEach(([, c, t], j) => {
      const x = 700 + j * 138;
      P.push(R(x, 372, 10, 10, {fill: c, forma: 'elipse'}));
      P.push(T(x + 14, 366, 124, 22, t, {pt: 6, c: PAINEL.txt, va: 'm'}));
    });
    const gx = 548, gy = 400, gw = 700, gh = 220;
    for (let t = 0; t <= 100; t += 20) {
      const y = gy + gh - gh * t / 100;
      P.push(R(gx, y, gw, 0.8, {fill: '28505C'}));
      P.push(T(gx - 42, y - 8, 36, 16, t + '%', {pt: 6, c: PAINEL.apoio, al: 'r', va: 'm'}));
    }
    const anos = s.anos || [], slot = gw / Math.max(anos.length, 1), bw = Math.min(14, slot / 5.2);
    anos.forEach((an, j) => {
      const x0 = gx + j * slot + (slot - bw * 4 - 6) / 2;
      series.forEach(([ch, c], q) => {
        const h = gh * Math.max(0, Math.min(1, an[ch] || 0));
        if (h > 0.5) P.push(R(x0 + q * (bw + 2), gy + gh - h, bw, h, {fill: c}));
      });
      P.push(T(gx + j * slot, gy + gh + 4, slot, 18, an.ano, {pt: 6.5, c: PAINEL.apoio, al: 'c', va: 'm'}));
    });
    return P;
  };

  /* ---------------- contratos de embrião ---------------- */
  L.contratos = s => {
    const P = claro(s, []);
    const cols = [[53.8, 256, 'l'], [313.6, 211.2, 'l'], [527.4, 83.2, 'c'], [614.4, 230.4, 'l'],
                  [848.6, 44.8, 'c'], [897.3, 128, 'r'], [1029.1, 198.4, 'c']];
    cabecalho(P, 124.2, 34.6, cols.map((c, j) => [c[0], c[1], c[2], s.cols[j]]), 7.5);
    if (!s.rows.length) {
      P.push(T(64, 200, 1100, 40, 'Nenhum contrato neste recorte.', {pt: 11, c: COR.cinza, va: 'm'}));
      return P;
    }
    const p = passo(s.rows.length, 158.7, 712, 34.45);
    s.rows.forEach((r, i) => {
      const y = 158.7 + i * p, hh = p - 1.3;
      const cores = [COR.tinta, COR.tinta, COR.cinza, COR.cinza, COR.cinza, COR.tinta, COR.verde];
      r.forEach((v, j) => {
        const [x, w, al] = cols[j];
        const ins = al === 'l' ? 12.8 : 0, insR = al === 'r' ? 12.8 : 0;
        P.push(C(x + ins, y, w - ins - insR, hh, v == null || v === '' ? '—' : v,
          {pt: j === 6 ? 7 : 7.5, b: j === 6 ? 1 : 0, c: cores[j], al, va: 'm', min: 6}));
      });
    });
    return P;
  };

  /* ---------------- manejo, mês a mês ---------------- */
  L.manejo = s => {
    const P = claro(s, []);
    const n = s.itens.length;
    const p = passo(n, 124.2, 712, 78.8), h = p - 6.5;
    const atual = String(s.atual || '').toUpperCase();
    s.itens.forEach(([m, t], i) => {
      const y = 124.2 + i * p, eAtual = String(m || '').toUpperCase().slice(0, 3) === atual.slice(0, 3) && i === n - 1;
      P.push(R(44.8, y, 92.2, h, {fill: eAtual ? COR.ouro : COR.azul}));
      P.push(T(44.8, y, 92.2, h, String(m || '').slice(0, 1).toUpperCase() + String(m || '').slice(1, 3).toLowerCase(),
        {pt: 12, b: 1, c: eAtual ? COR.navy : 'FFFFFF', al: 'c', va: 'm'}));
      P.push(R(143.4, y, 1091.8, h, {fill: eAtual ? COR.destaque : (i % 2 ? COR.zebra : 'FFFFFF'), line: 'DEDEDE'}));
      let pt = 8.5;
      while (pt > 6.5 && nLinhas(t, pt, 1066.2) * altLinha(pt) > h - 8) pt -= 0.25;
      P.push(T(156.2, y + 4, 1066.2, h - 8, t, {pt, c: COR.tinta, va: 'm', wrap: 1}));
    });
    return P;
  };

  /* ---------------- comentários do mês ---------------- */
  /* comentários: a faixa escura do relatório — categoria e ∆ num cartão à
     esquerda, o texto à direita (uma natureza por linha quando vem do Trello).
     A faixa cresce com o texto; se o slide não comporta, o corpo desce. */
  const linhasP = (txt, pt, w) => String(txt || '').split('\n').reduce((a, par) => a + nLinhas(par, pt, w), 0);
  L.comentarios = s => {
    const P = claro(s, []);
    const Y0 = 124.2, GAP = 1.2, W_TXT = 966.4;
    const alt = p => s.itens.map(i => Math.max(63.9, linhasP(i.txt, p, W_TXT) * altLinha(p) + 10.4));
    let pt = 7.63;
    while (pt > 6.25 && alt(pt).reduce((a, b) => a + b + GAP, 0) > 712 - Y0) pt -= 0.125;
    let y = Y0;
    alt(pt).forEach((h, i) => {
      const it = s.itens[i];
      P.push(R(44.8, y, 1187.8, h, {fill: COR.navy}));
      P.push(R(243.2, y, 986.9, h, {fill: i % 2 ? '0F2238' : '0A1828'}));
      P.push(C(53.8, y + 3.9, 179.2, 29.3, it.cat, {pt: 8.14, b: 1, c: 'FFFFFF', va: 'm', min: 6.5}));
      const d = String(it.delta || '').trim();
      if (d) {
        const neg = /^[-−]/.test(d), zero = !/[1-9]/.test(d);
        P.push(C(53.8, y + 33.8, 179.2, 27.4, d, {pt: 8.64, b: 1, c: zero ? 'BBBBBB' : (neg ? 'FFAAAA' : '7BE0A0'), va: 'm'}));
      }
      P.push(T(262.4, y + 5.2, W_TXT, h - 10.4, it.txt, {pt, c: 'FFFFFF', va: 'm', wrap: 1}));
      y += h + GAP;
    });
    return P;
  };

  /* ---------------- fotos: inteiras, lado a lado ---------------- */
  /* O relatório põe a foto inteira, na proporção dela, ocupando a altura útil —
     não recorta em quadradinhos. Aqui: tenta uma fileira; se ela ficar baixa
     demais, divide em duas; fica a arrumação que dá mais área de foto. */
  function arrumaFotos(itens) {
    const X0 = 14, X1 = 1266, Y0 = 114, Y1 = 708, GAP = 10;
    const ar = itens.map(f => (f.w && f.h) ? f.w / f.h : 4 / 3);
    const fileira = (lista, y, hMax) => {
      const soma = lista.reduce((a, i) => a + ar[i], 0);
      const h = Math.min(hMax, (X1 - X0 - GAP * (lista.length - 1)) / soma);
      const larg = soma * h + GAP * (lista.length - 1);
      let x = X0 + (X1 - X0 - larg) / 2;
      return lista.map(i => { const w = ar[i] * h; const r = {i, x, y, w, h}; x += w + GAP; return r; });
    };
    const area = rs => rs.reduce((a, r) => a + r.w * r.h, 0);
    const todos = itens.map((_, i) => i);
    let melhor = fileira(todos, Y0, Y1 - Y0);
    melhor = melhor.map(r => Object.assign(r, {y: Y0 + (Y1 - Y0 - r.h) / 2}));
    /* três paisagens numa fileira ficam baixinhas no meio do slide; o relatório
       as põe grandes, ocupando a altura. Mosaico: a primeira à esquerda, as
       outras duas empilhadas, com recorte leve para fechar as caixas */
    if (itens.length === 3 && melhor[0].h < 400) {
      const wa = (X1 - X0 - GAP) * 0.6, wb = X1 - X0 - GAP - wa, hb = (Y1 - Y0 - GAP) / 2;
      return [{i: 0, x: X0, y: Y0, w: wa, h: Y1 - Y0, corta: 1},
              {i: 1, x: X0 + wa + GAP, y: Y0, w: wb, h: hb, corta: 1},
              {i: 2, x: X0 + wa + GAP, y: Y0 + hb + GAP, w: wb, h: hb, corta: 1}];
    }
    if (itens.length >= 3) {
      for (let corte = 1; corte < itens.length; corte++) {
        const hRow = (Y1 - Y0 - GAP) / 2;
        const a = fileira(todos.slice(0, corte), Y0, hRow), b = fileira(todos.slice(corte), Y0 + hRow + GAP, hRow);
        const ha = a[0].h, hb = b[0].h, sobra = (Y1 - Y0 - GAP - ha - hb) / 2;
        a.forEach(r => { r.y = Y0 + sobra; });
        b.forEach(r => { r.y = Y0 + sobra + ha + GAP; });
        if (area(a.concat(b)) > area(melhor) * 1.08) melhor = a.concat(b);
      }
    }
    return melhor;
  }
  L.fotos = (s, ctx) => {
    const P = claro(s, []);
    const itens = (s.fotos || []).map(f => typeof f === 'string' ? {img: f} : f);
    arrumaFotos(itens).forEach(r => {
      const f = itens[r.i];
      if (f.img) P.push(I(r.x, r.y, r.w, r.h, f.img, {link: f.video && ctx && ctx.linkVideo ? ctx.linkVideo(f.video) : null,
        video: f.video || null, corta: r.corta ? 1 : 0, ar: (f.w && f.h) ? f.w / f.h : null}));
      else P.push(R(r.x, r.y, r.w, r.h, {fill: 'F4F6F8', line: COR.linha}));
      if (f.video) {
        const cx = r.x + r.w / 2, cy = r.y + r.h / 2;
        P.push(R(cx - 27, cy - 27, 54, 54, {fill: '04223B', alfa: 0.62, line: 'FFFFFF', forma: 'elipse', link: ctx && ctx.linkVideo ? ctx.linkVideo(f.video) : null}));
        P.push(T(cx - 27, cy - 27, 54, 54, '▶', {pt: 14, b: 1, c: 'FFFFFF', al: 'c', va: 'm'}));
      }
    });
    return P;
  };

  /* ---------------- genéricos (compatibilidade) ---------------- */
  L.tabela = s => {
    if (s.n === 23) return programacao(s);
    const P = claro(s, []);
    const n = s.cols.length, w = 1190.4 / n;
    cabecalho(P, 124.2, 34.6, s.cols.map((c, j) => [44.8 + j * w + 9, w - 18, j ? 'r' : 'l', c]));
    const p = passo(s.rows.length, 158.7, 712, 32);
    s.rows.forEach((r, i) => {
      const y = 158.7 + i * p;
      P.push(R(44.8, y, 1190.4, p - 1.3, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: COR.linha}));
      r.forEach((v, j) => P.push(C(44.8 + j * w + 9, y, w - 18, p - 1.3, v == null ? '—' : v, {pt: 8, c: COR.tinta, al: j ? 'r' : 'l', va: 'm'})));
    });
    return P;
  };
  L.kpis_tabela = s => {
    const P = claro(s, []);
    cartoes(P, s.kpis || [], 108.8, 20);
    const t = s.tabela || {cols: [], rows: []};
    const n = t.cols.length || 1, w = 1190.4 / n;
    cabecalho(P, 263.7, 34.6, t.cols.map((c, j) => [44.8 + j * w + 9, w - 18, j ? 'c' : 'l', c]));
    const p = passo(t.rows.length, 298.2, 712, 36);
    t.rows.forEach((r, i) => {
      const y = 298.2 + i * p;
      P.push(R(44.8, y, 1190.4, p - 1.3, {fill: i % 2 ? COR.zebra : 'FFFFFF', line: COR.linha}));
      r.forEach((v, j) => P.push(C(44.8 + j * w + 9, y, w - 18, p - 1.3, v, {pt: 8.5, c: COR.tinta, al: j ? 'c' : 'l', va: 'm'})));
    });
    return P;
  };
  L.pendente = s => claro(s, []);

  const ESCUROS = new Set(['capa', 'agenda', 'divisor', 'encerramento']);
  const fundo = s => ESCUROS.has(s.t) ? COR.navy : 'FFFFFF';
  const prims = (s, ctx) => (L[s.t] || L.pendente)(s, ctx || {});

  /* ---------------- saída HTML ---------------- */
  const escH = x => String(x == null ? '' : x).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
  const rgba = (hex, a) => {
    const n = parseInt(hex, 16);
    return a == null ? '#' + hex : `rgba(${n >> 16 & 255},${n >> 8 & 255},${n & 255},${a})`;
  };
  const JUST = {l: 'flex-start', c: 'center', r: 'flex-end'};
  const ALINH = {t: 'flex-start', m: 'center', b: 'flex-end'};
  function arcoSVG(p) {
    const pt = (a, r) => [p.cx + r * Math.sin(a * Math.PI / 180), p.cy - r * Math.cos(a * Math.PI / 180)];
    const r1 = p.r, r0 = p.r - p.esp, grande = (p.a1 - p.a0) > 180 ? 1 : 0;
    const [x1, y1] = pt(p.a0, r1), [x2, y2] = pt(p.a1, r1), [x3, y3] = pt(p.a1, r0), [x4, y4] = pt(p.a0, r0);
    return `<path d="M${x1} ${y1} A${r1} ${r1} 0 ${grande} 1 ${x2} ${y2} L${x3} ${y3} A${r0} ${r0} 0 ${grande} 0 ${x4} ${y4} Z" fill="#${p.fill}"/>`;
  }
  function html(ps) {
    const arcos = ps.filter(p => p.k === 'a');
    return ps.map(p => {
      if (p.k === 'h') return p.html;
      if (p.k === 'a') return '';
      const box = `left:${p.x.toFixed(1)}px;top:${p.y.toFixed(1)}px;width:${p.w.toFixed(1)}px;height:${p.h.toFixed(1)}px`;
      if (p.k === 'r') {
        const st = box + (p.fill ? `;background:${rgba(p.fill, p.alfa)}` : '') +
          (p.line ? `;border:1px solid #${p.line}` : '') + (p.forma === 'elipse' ? ';border-radius:50%' : '');
        return p.link ? `<a class="pr" href="${escH(p.link)}" data-video="${escH(p.video || '')}" style="${st}"></a>`
                      : `<div class="pr" style="${st}"></div>`;
      }
      if (p.k === 'i') {
        const img = `<img class="pi${p.logo ? ' logo' : ''}${p.corta ? ' corta' : ''}" src="${escH(p.src)}" alt="" style="${box}">`;
        return p.link ? `<a href="${escH(p.link)}" data-video="${escH(p.video || '')}">${img}</a>` : img;
      }
      const st = `${box};font-size:${(p.pt * PT).toFixed(2)}px;color:${rgba(p.c || COR.tinta, p.alfa)};` +
        `justify-content:${JUST[p.al || 'l']};align-items:${ALINH[p.va || 'm']};text-align:${{l: 'left', c: 'center', r: 'right'}[p.al || 'l']}` +
        (p.b ? ';font-weight:700' : '') + (p.i ? ';font-style:italic' : '') +
        (p.sp ? `;letter-spacing:${(p.sp * PT).toFixed(2)}px` : '');
      return `<div class="pt${p.wrap ? ' quebra' : ''}" style="${st}"><span>${escH(p.t)}</span></div>`;
    }).join('') + (arcos.length ? `<svg class="parcos" viewBox="0 0 1280 720" width="1280" height="720">${arcos.map(arcoSVG).join('')}</svg>` : '');
  }

  /* ---------------- saída PPTX ---------------- */
  const AL = {l: 'left', c: 'center', r: 'right'};
  const VA = {t: 'top', m: 'middle', b: 'bottom'};
  const pol = v => +(v / PX_IN).toFixed(4);
  function pptx(pres, sl, ps, imagens) {
    for (const p of ps) {
      if (p.k === 'h') continue;
      if (p.k === 'a') {
        const d = p.r * 2;
        const a0 = ((p.a0 - 90) % 360 + 360) % 360, a1 = ((p.a1 - 90) % 360 + 360) % 360;
        sl.addShape(pres.ShapeType.blockArc, {x: pol(p.cx - p.r), y: pol(p.cy - p.r), w: pol(d), h: pol(d),
          fill: {color: p.fill}, angleRange: [a0, a1], arcThicknessRatio: p.esp / p.r});
        continue;
      }
      const g = {x: pol(p.x), y: pol(p.y), w: pol(Math.max(p.w, 0.5)), h: pol(Math.max(p.h, 0.5))};
      if (p.k === 'r') {
        const o = Object.assign({}, g);
        if (p.fill) o.fill = {color: p.fill, transparency: p.alfa != null ? Math.round((1 - p.alfa) * 100) : 0};
        if (p.line) o.line = {color: p.line, width: 0.75};
        if (p.link) o.hyperlink = {url: p.link, tooltip: 'Abrir o vídeo no hub'};
        sl.addShape(p.forma === 'elipse' ? pres.ShapeType.ellipse : pres.ShapeType.rect, o);
        continue;
      }
      if (p.k === 'i') {
        const data = imagens && imagens[p.src] ? imagens[p.src] : (String(p.src).startsWith('data:') ? p.src : null);
        if (!data) continue;
        const o = Object.assign({data}, g);
        if (p.corta && p.ar) {
          /* recorte: o pptxgen quer o tamanho da IMAGEM inteira em w/h e a
             janela visível em sizing — a imagem cobre a caixa e sobra dos lados */
          let iw = p.w, ih = p.w / p.ar;
          if (ih < p.h) { ih = p.h; iw = p.h * p.ar; }
          Object.assign(o, {w: pol(iw), h: pol(ih), sizing: {type: 'cover', w: g.w, h: g.h}});
        }
        if (p.link) o.hyperlink = {url: p.link, tooltip: 'Abrir o vídeo no hub'};
        sl.addImage(o);
        continue;
      }
      const o = Object.assign({}, g, {fontFace: FONTE, fontSize: +p.pt.toFixed(2), bold: !!p.b, italic: !!p.i,
        color: p.c || COR.tinta, align: AL[p.al || 'l'], valign: VA[p.va || 'm'], margin: [0, 0, 0, 0],
        wrap: !!p.wrap, fit: 'none'});
      if (p.sp) o.charSpacing = p.sp;
      if (p.alfa != null) o.transparency = Math.round((1 - p.alfa) * 100);
      sl.addText(p.t, o);
    }
  }

  window.LAYOUT = {prims, html, pptx, fundo, ESCUROS, COR, LOGO_OURO, LOGO_NAVY, largura, encaixa};
})();
