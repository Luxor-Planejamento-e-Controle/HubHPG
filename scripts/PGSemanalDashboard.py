"""
PGSemanalDashboard — gera dashboard_semanal.html (self-contained) do semanal_data.json.

Cada semana = um snapshot completo (docx-semente OU capturado pelo script). O dashboard
é 100% orientado ao snapshot da semana selecionada. Seletor lista só semanas com snapshot,
mais recente no topo. Read-only por padrão; botão Editar habilita override (localStorage).
Camada loadState/saveState isolada p/ Supabase depois.

Uso:
    python PGSemanalDocx.py            # 1x: parseia relatórios oficiais -> semanal_docx.json
    python PGSemanalReport.py --seed   # 1x: semeia histórico (baseline) dos docx
    python PGSemanalReport.py DD/MM/AAAA DD/MM/AAAA   # semana nova (automático)
    python PGSemanalDashboard.py       # gera o HTML
"""

from __future__ import annotations

import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent  # raiz do projeto (scripts/ fica 1 nível abaixo)
JSON_IN = BASE_DIR / "bases" / "semanal_data.json"
HTML_OUT = BASE_DIR / "dashboards" / "dashboard_semanal.html"

TEMPLATE = r"""<!doctype html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Atualização Semanal — Haras Pao Grande</title>
<style>
  :root{/* Marca Haras Pao Grande — navy #04223B (289C) + dourado #CA9703 (117C) */
    --amber:#CA9703;--teal:#7FA8C4;--bg:#04223B;--card:#0A3050;--line:#1B486B;
    --txt:#EAF0F4;--mut:#93AABC;--pos:#4CC38A;--neg:#F07A7A}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--txt);
    font-family:"Segoe UI",system-ui,-apple-system,Arial,sans-serif;font-size:15px}
  header{padding:20px 28px 16px;border-bottom:1px solid var(--line);
    display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
  .brand{display:flex;align-items:center;gap:15px}
  .brand .logo{height:58px;width:auto;flex:none}
  h1{margin:0;font-size:21px;font-weight:600;letter-spacing:.3px}
  h1 .mark{color:var(--amber)} h1 .teal{color:var(--teal)}
  .wrap{max-width:1600px;margin:0 auto;padding:20px 28px 60px}
  .toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
  select,button{background:#072B49;color:var(--txt);border:1px solid var(--line);
    border-radius:8px;padding:8px 12px;font-size:14px;font-family:inherit;cursor:pointer}
  select{min-width:230px}
  button:hover{border-color:var(--amber)}
  button.on{background:var(--amber);color:#04223B;border-color:var(--amber);font-weight:600}
  .src{font-size:11px;color:var(--mut);border:1px solid var(--line);border-radius:20px;padding:3px 10px}
  .src.docx{color:var(--teal);border-color:var(--teal)}
  /* grid 2 colunas, ORDEM 1→5 (row-major). Seções largas ocupam a linha toda;
     receptoras+headcount ficam lado a lado. */
  .sections{display:grid;grid-template-columns:1fr 1fr;gap:18px;align-items:stretch}
  .panel{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 22px;
    display:flex;flex-direction:column}
  .panel.full{grid-column:1 / -1}
  /* mobile: o nº de colunas dos KPIs vem inline via JS (repeat(N,...), N até 7
     nas seções largas — ver renderPanel no script) porque cada seção tem uma
     contagem diferente de cards. !important pisa nesse inline pra achatar em
     telas estreitas, senão o painel de PRODUÇÃO tenta encaixar 7 cards numa
     coluna só e os números ficam espremidos/sobrepostos. */
  @media(max-width:900px){
    .sections{grid-template-columns:1fr}
    .kpis{grid-template-columns:repeat(3,minmax(0,1fr)) !important}
    .wrap{padding:16px 16px 40px}
  }
  @media(max-width:600px){
    .kpis{grid-template-columns:repeat(2,minmax(0,1fr)) !important}
    header{padding:14px 16px 12px;gap:10px}
    .brand .logo{height:38px}
    h1{font-size:16px}
    .wrap{padding:14px 12px 32px}
    .panel{padding:14px 16px}
    select{min-width:0;flex:1 1 auto}
    .kpi{min-height:88px;padding:11px 12px 10px}
    .kpi .val{font-size:24px}
    table{font-size:11px}
    th,td{padding:5px 7px}
  }
  @media(max-width:380px){.kpis{grid-template-columns:1fr !important}}
  .panel h2{margin:0;font-size:16px;color:var(--amber);font-weight:600;letter-spacing:.3px}
  .panel h2 .n{color:var(--teal);font-weight:700;margin-right:4px}
  .panel .sub{color:var(--mut);font-size:12.5px;margin:4px 0 0}
  .kpis{display:grid;gap:12px;margin-top:14px}
  .kpi{background:#072B49;border:1px solid var(--line);border-radius:10px;
    padding:13px 15px 12px;position:relative;display:flex;flex-direction:column;
    justify-content:flex-start;min-height:104px}
  .kpi .lab{color:var(--mut);font-size:10.5px;text-transform:uppercase;letter-spacing:.45px;
    line-height:1.3;min-height:28px}
  .kpi .val{font-size:30px;font-weight:700;margin-top:7px;outline:none;line-height:1}
  .kpi.editing .val{border-bottom:2px dashed var(--amber);cursor:text}
  .kpi.edited{border-color:var(--amber)}
  .kpi.manual{border-style:dashed;border-color:var(--teal)}
  .kpi .tag{position:absolute;top:8px;right:10px;font-size:8px;color:var(--teal);
    text-transform:uppercase;letter-spacing:.5px;font-weight:700}
  .kpi .rst{position:absolute;bottom:5px;right:8px;font-size:10px;color:var(--mut);cursor:pointer;display:none}
  .kpi.edited.editing .rst{display:block}
  .kpi .rst:hover{color:var(--neg)}
  .kpi .val .pos{color:var(--pos)} .kpi .val .neg{color:var(--neg)}
  /* comentário do KPI: a composição do número, não o número */
  /* comentário curto sob o número — só onde o próprio relatório escreve a
     abertura. Lista continua sendo tabela, nunca texto dentro do card. */
  .kpi .val .nota{display:block;font-size:11.5px;font-weight:400;color:var(--mut);
    line-height:1.35;margin-top:7px}
  /* detalhe de uma coluna só: lista, não tabela */
  .det-lista{color:var(--txt);font-size:12.5px;line-height:1.6;
    border:1px solid var(--line);border-radius:8px;padding:8px 11px}
  .kpi .chk{position:absolute;top:6px;right:8px;font-size:9px;font-weight:700;letter-spacing:.3px}
  /* link de download com cara de botão: é ação, tem de parecer ação */
  .btn-pdf{display:inline-block;margin-left:8px;padding:6px 12px;border-radius:8px;
    border:1px solid var(--line-strong,var(--line));background:transparent;
    color:var(--ink,#EAF0F4);font-size:12.5px;text-decoration:none;cursor:pointer;
    transition:.13s}
  .btn-pdf:hover{background:var(--surface-2,#0F3E63);border-color:var(--teal)}
  .kpi .chk.ok{color:var(--pos)} .kpi .chk.no{color:var(--amber)}
  /* Impressão = o PDF do dashboard. O navegador é o motor; aqui só se garante que
     nada fica cortado: sem scroll interno, sem barra de ferramentas, e as cores
     do tema preservadas (senão sai um documento branco sem hierarquia). */
  /* espelha as regras de impressão que mudam a ALTURA, para a medição do zoom
     bater com o que vai para o papel */
  .medindo-print .toolbar,.medindo-print .tag{display:none}
  .medindo-print .panel{padding:8px 10px;margin:0 0 6px}
  .medindo-print .kpi{min-height:0;padding:6px 8px}
  .medindo-print .kpi .lab{font-size:9.5px;min-height:0}
  .medindo-print .kpi .val{font-size:21px;margin-top:1px}
  .medindo-print .kpi .val .nota{font-size:9.5px;margin-top:2px}
  .medindo-print .kpis{gap:8px;margin-top:8px}
  .medindo-print table{font-size:10px}
  .medindo-print thead th,.medindo-print tbody td{padding:2px 5px}
  .medindo-print .sections{display:block;gap:0}
  .medindo-print .det{margin-top:6px;padding-top:4px}
  .medindo-print .det-h{font-size:10.5px;margin-bottom:3px}
  .medindo-print .panel h2{font-size:14px;margin-bottom:1px}
  .medindo-print header{padding:4px 0 6px;min-height:0}
  .medindo-print header .logo{height:24px}
  .medindo-print header h1{font-size:16px}
  @media print{
    /* margem ZERO na folha: com margem, o papel branco aparece em volta do painel
       escuro. O respiro vira padding do body, que já é da cor do tema. */
    @page{size:A4 landscape;margin:0}
    html,body{background:var(--bg) !important;-webkit-print-color-adjust:exact;
      print-color-adjust:exact;font-size:12.5px}
    body{padding:6mm !important}
    .wrap{max-width:none !important}
    .toolbar,#btnEdit,#btnImg,#btnPdf,#pdfEmb,.rst,.tag{display:none !important}
    .det-b,.scroll{overflow:visible !important;max-height:none !important}
    /* O painel INTEIRO com break-inside:avoid empurrava cada seção para uma folha
       nova e deixava meia página em branco — foi o que saiu na primeira tentativa.
       Quem não pode partir é a faixa de KPIs e cada tabela; o painel pode. */
    .panel{break-inside:auto;page-break-inside:auto;padding:10px 12px;margin:0 0 8px}
    .panel:last-child{margin-bottom:0}
    /* título de seção não fica órfão no pé da folha, separado dos seus KPIs */
    .panel h2{break-after:avoid;page-break-after:avoid}
    /* sem break-avoid: com o zoom o conteúdo cabe inteiro, e o avoid era justamente
       o que empurrava a última tabela para a segunda folha */
    .kpis,.det{break-inside:auto;page-break-inside:auto}
    .sections{display:block !important;gap:0 !important}
    /* folha em branco no fim: qualquer altura sobrando depois do conteúdo vira
       uma página inteira pintada de fundo escuro */
    html,body{height:auto !important;min-height:0 !important;margin:0 !important}
    /* cabe numa folha: o zoom é calculado no clique, medindo a altura real */
    body{zoom:var(--print-zoom,1)}
    .wrap{margin:0 !important;padding:0 !important}
    header{padding:4px 0 6px !important;margin:0 !important;min-height:0 !important}
    header .logo{height:24px}
    header h1{font-size:16px}
    .panel{padding:8px 10px;margin:0 0 6px}
    .panel h2{font-size:14px;margin-bottom:1px}
    .kpi{min-height:0 !important;padding:6px 8px}
    .kpi .lab{font-size:9.5px;min-height:0}
    .kpi .val{font-size:21px;margin-top:1px}
    .kpi .val .nota{font-size:9.5px;margin-top:2px}
    .kpis{gap:8px;margin-top:8px}
    .det{margin-top:6px;padding-top:4px}
    .det-h{font-size:10.5px;margin-bottom:3px}
    table{font-size:10px}
    thead th,tbody td{padding:2px 5px}
  }
  /* detalhe integrado no card, sempre visível, sem scroll lateral */
  .det{margin-top:16px;border-top:1px solid var(--line);padding-top:12px}
  .det-h{color:var(--teal);font-size:11.5px;font-weight:700;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:8px}
  .det-h .c{color:var(--amber);margin-left:4px}
  /* tabela inteira à vista: nada de scroll interno vertical. Só o horizontal fica,
     pra tabela larga não empurrar a página. */
  .det-b{overflow-x:auto;border:1px solid var(--line);border-radius:8px}
  table{width:100%;border-collapse:collapse;font-size:12px}
  th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;line-height:1.35}
  /* sem scroll interno o sticky grudaria na viewport da página e vários cabeçalhos
     de tabela ficariam flutuando um sobre o outro — melhor fixo na tabela */
  th{color:var(--mut);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.3px;
    background:#0c2740;white-space:nowrap}
  td{color:var(--txt);white-space:normal;overflow-wrap:break-word}
  tr:last-child td{border-bottom:none}
</style></head>
<body>
<header>
  <div class="brand">
    <img class="logo" src="data:image/png;base64,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" alt="Haras Pao Grande">
    <h1><span class="mark">Atualização Semanal</span> <span class="teal">•</span> Haras Pao Grande</h1>
  </div>
  <div class="toolbar">
    <select id="semana"></select>
    <button id="btnEdit">Editar</button>
    <button id="btnImg">Exportar imagem</button>
    <button id="btnPdf">Exportar PDF</button>
    <span id="pdfEmb"></span>
  </div>
</header>
<div class="wrap"><div class="sections" id="sections"></div></div>

<script>
const DATA = __DATA__;
const CAL = DATA.calendario||[];
const SNAPS = DATA.snapshots||{};       // dados CALCULADOS (o que vai pro dash)
const KEY_OV="hpg_semanal_overrides_v1", KEY_WK="hpg_semanal_semana_v1"; // NUNCA bump
function loadState(k){ try{return JSON.parse(localStorage.getItem(k));}catch(e){return null;} }
function saveState(k,v){ localStorage.setItem(k,JSON.stringify(v)); }
let overrides = loadState(KEY_OV) || {};
const IDS = CAL.map(w=>w.id);
const LATEST = CAL.length? CAL[CAL.length-1].id : null;
let semana = loadState(KEY_WK) || DATA.semana_atual || LATEST;
if(!IDS.includes(semana)) semana = LATEST;
let editMode=false;

// overrides POR SEMANA (dado manual varia toda semana): chave = "semana|path"
function ovKey(p){ return semana+"|"+p; }
function hasOv(p){ return ovKey(p) in overrides; }
function getOv(p){ return overrides[ovKey(p)]; }
function setOv(p,v){ overrides[ovKey(p)]=v; saveState(KEY_OV,overrides); }
function delOv(p){ delete overrides[ovKey(p)]; saveState(KEY_OV,overrides); }
function kByP(p){ for(const s of SECTIONS)for(const k of s.kpis)if(k.p===p)return k; return null; }
function effVal(p){ if(hasOv(p))return getOv(p); const k=kByP(p); return k?rawVal(k):null; }

function snap(){ return SNAPS[semana]||{}; }
function curWeek(){ return CAL.find(w=>w.id===semana)||null; }
function br(iso){ if(!iso)return""; const p=iso.split("-"); return p[2]+"/"+p[1]+"/"+p[0].slice(2); }
function g(getter){ try{ return getter(snap()); }catch(e){ return null; } }
function deltaTxt(s){
  // formato +entradas / −saídas. Vem do diff da população (automático); override
  // manual em sai.en/sai.sa continua valendo se alguém corrigir na mão.
  // Lê do snapshot, não de rawVal: os KPIs de saída/entrada devolvem HTML.
  const hc=s.headcount||{};
  const e=Number(hasOv("sai.en")?getOv("sai.en"):hc.delta_entradas)||0;
  const x=Number(hasOv("sai.sa")?getOv("sai.sa"):hc.delta_saidas)||0;
  return `<span class="pos">+${e}</span> / <span class="neg">−${x}</span>`;
}

// nome de local em texto corrido
const LOCNOME={"PAO GRANDE":"Pao Grande","ARRENDAMENTO CESAR FURTADO":"arrendamento",
  "SOCIO":"sócio","SOCIO - VENDIDA":"sócio (vendida)","VENDIDA":"vendida",
  "MATO GROSSO":"Mato Grosso","VOLTOU PARA A CENTRAL":"central"};
function loc(v){ return v? (LOCNOME[v]||String(v).toLowerCase()) : v; }

/* "309" sozinho não diz nada. Descreve quem é e o trajeto, no próprio card:
   receptora 309 · Pao Grande → sócio. Até 3; acima disso a tabela de detalhe
   entra (ver minRows no det) e aqui fica só a contagem. */
// Card e quadrado: rotulo + numero. Quem lista e a tabela da secao, no mesmo
// formato das outras — despejar as movimentacoes dentro do card misturava os dois
// padroes e ainda repetia o que a tabela 'Entradas na semana' ja mostra.
function movVal(s,n,rows){ return n==null? null : String(n); }

// estrutura espelha o docx. cada KPI: {p:chave override, l:label, get:fn(snap)->valor}
const SECTIONS = [
 {n:"1", t:"PRODUÇÃO", wide:true, kpis:[
    {p:"prod.conf", l:"Embriões confirmados na semana", get:s=>s.confirmados_semana},
    // transição de estação: o relatório publica as duas safras enquanto a nova não anda
    {p:"prod.acum", l:s=>`Acumulado na estação ${s.safra_atual_rotulo||""}`.trim(),
     get:s=>s.acumulado_estacao},
    {p:"prod.acum2", l:s=>`Acumulado na estação ${s.safra_proxima_rotulo||""}`.trim(),
     skip:s=>!s.safra_proxima_rotulo, get:s=>s.acumulado_estacao_proxima},
    {p:"prod.mes", l:"Acumulado no mês", get:s=>s.acumulado_mes},
    {p:"prod.nasc", l:"Nascimentos na semana", get:s=>s.nascimentos},
    {p:"prod.abor", l:"Abortos / óbitos na semana", get:s=>s.abortos_obitos},
  ], det:[["Embriões confirmados na semana","confirmados"],["Nascimentos na semana","nascimentos"],["Abortos / óbitos na semana","abortos_obitos"]]},
 {n:"2", t:"RECEPTORAS", kpis:[
    {p:"rec.tot", l:"Total receptoras", get:s=>s.receptoras?.total},
    {p:"rec.pre", l:"Prenhas", get:s=>s.receptoras?.prenhas},
    {p:"rec.vaz", l:"Vazias", get:s=>s.receptoras?.vazias},
    {p:"rec.doa", l:"Doadoras (estação)", get:s=>s.receptoras?.doadoras},
    // ciclando fica ao lado de doadoras: as duas contam a mesma égua, uma no
    // cadastro e outra no estado reprodutivo. Sem fonte em planilha — avaliação do
    // veterinário, em _cache/semanal_manual.json (override do navegador não entra
    // no snapshot).
    {p:"rec.cic", l:"Doadoras ciclando", manual:true,
     get:s=>s.receptoras?.doadoras_ciclando},
    {p:"rec.idx", l:"Índice eficiência (vazias/doadoras)"},  // derivado: vazias ÷ doadoras
 ]},
 {n:"3", t:"HEADCOUNT", kpis:[
    {p:"hc.tot", l:"Total geral de animais", get:s=>s.headcount?.total},
    {p:"hc.del", l:"Δ vs semana passada", delta:true},
    {p:"hc.fpg", l:"Fazenda Pao Grande", get:s=>s.headcount?.fazenda_pg},
    {p:"hc.arr", l:"Arrendamento", get:s=>s.headcount?.arrendamento},
    {p:"hc.cte", l:"Centro de Treinamento", get:s=>s.headcount?.cte},
    {p:"hc.soc", l:"Sócios", get:s=>s.headcount?.socio},
 ]},
 {n:"4", t:"TERCEIROS NA PROPRIEDADE", wide:true, kpis:[
    // "Total terceiros" É "vendidos pendentes" (o próprio relatório oficial escreve
    // "05 (vendidos pendentes)" nesta linha) — só animal, embrião não ocupa espaço
    // físico e tem lista própria na seção 5. A lista de detalhe fica logo abaixo,
    // nesta seção, não na 5 — pedido de 28/08/2026.
    {p:"ter.tot", l:"Total terceiros", get:s=>s.terceiros?.terceiros_propriedade},
    {p:"ter.doa", l:"Doadoras terceiros", get:s=>s.terceiros?.doadoras_terceiros},
    {p:"ter.out", l:"Outros terceiros (cavalgada / treino)", get:s=>s.terceiros?.outros_terceiros},
 ], det:[["Vendidos pendentes","terceiros_propriedade"]]},
 {n:"5", t:"SAÍDAS", wide:true, kpis:[
    {p:"sai.sa", l:"Saídas na semana",
     get:s=>movVal(s,s.movimento?.saidas,(s.detalhe||{}).saidas)},
    // o relatório escreve a abertura nesta linha: "07 (05 animais e 02 embriões)"
    {p:"sai.vp", l:"Vendidos pendentes de saída", html:true, get:s=>{
      const t=s.terceiros||{}, n=t.vendidos_pendentes;
      if(n==null) return null;
      const a=t.vendidos_pendentes_animais, e=t.vendidos_pendentes_embrioes;
      const partes=[];
      if(a) partes.push(`${a} ${a===1?"animal":"animais"}`);
      if(e) partes.push(`${e} ${e===1?"embrião":"embriões"}`);
      return partes.length? `${n}<span class="nota">${partes.join(" e ")}</span>` : String(n);
    }},
    // mesma descrição do relatório oficial: "07 (2 animais e 5 embriões)". A quebra
    // vai em fonte menor, como comentário — o número é que é o KPI.
    // o relatório publica ANIMAIS nesta linha; embrião tem KPI próprio ao lado.
    // Semana antiga (antes da abertura existir) cai no total, que era só animal.
    {p:"sai.sp", l:"Animais em sociedade pendentes de saída",
     get:s=>s.terceiros?.sociedade_pendentes_animais ?? s.terceiros?.sociedade_pendentes},
    {p:"sai.se", l:"Embriões em sociedade aguardando entrega",
     get:s=>s.terceiros?.sociedade_pendentes_embrioes},
    {p:"sai.tr", l:"Transferências internas", get:s=>s.movimento?.transferencias},
    {p:"sai.en", l:"Entradas na semana",
     get:s=>movVal(s,s.movimento?.entradas,(s.detalhe||{}).entradas)},
 ], det:[["Saídas na semana","saidas"],["Entradas na semana","entradas"],
         ["Embriões vendidos pendentes de saída","terceiros_vendidos_embrioes"],
         ["Em sociedade pendentes de saída","terceiros_sociedade"]]},
];

function rawVal(k){
  if(k.delta) return deltaTxt(snap());
  if(k.p==="rec.idx"){                                   // derivado: vazias ÷ doadoras (usa manuais)
    const vaz=Number(effVal("rec.vaz")), doa=Number(effVal("rec.doa"));
    return (doa && !isNaN(vaz))? Math.round(vaz/doa*10)/10 : null;
  }
  return g(k.get);
}
function fmtVal(k){
  if(hasOv(k.p)) return getOv(k.p);
  const v=rawVal(k);
  if(k.delta) return v==null?"--":v;   // já vem com HTML
  return (v===null||v===undefined||v==="")?"--":v;
}

// rótulos amigáveis; mostro TODAS as colunas disponíveis (só escondo vazias + internas)
const LABELS={doadora:"Doadora",garanhao:"Garanhão",local:"Local",cotas_pg:"Cota PG",socio:"Sócio",
  fatia:"Fatia",data_ia:"IA / Cobrição",receptora:"Receptora",data_confirmacao:"Confirmação",
  sexo_potro:"Sexo",nome_potro:"Potro",data_paricao:"Parição",data_aborto:"Aborto",data_obito:"Óbito",
  status:"Status",categoria:"Categoria",comprador:"Comprador / Sócio",nome:"Nome",cota:"Cota",
  tipo:"Tipo",obs:"Obs",animal:"Animal",data:"Data",ocorrencia:"Ocorrência",produto:"Produto (animal)",
  mae:"Mãe",pai:"Pai"};
const HIDE=new Set(["key","confirmado","reposicao","origem","especie","chave",
  "afeta_headcount","era_receptora_contada"]);   // internos, não exibir
const DATECOLS=new Set(["data","data_ia","data_confirmacao","data_paricao","data_aborto","data_obito"]);
const PCTCOLS=new Set(["cota","cotas_pg"]);
function lab(c){ return LABELS[c] || c.replace(/_/g," "); }

function detTable(title, key, opts){
  let rows = (snap().detalhe||{})[key];
  if(!rows || !rows.length) return "";
  // minRows: detalhe que já cabe descrito no card só vira tabela quando é muito
  if(opts && opts.minRows && rows.length < opts.minRows) return "";
  let head, body, nCols=1, unica=null;
  if(typeof rows[0]==="string"){
    head="<th>Detalhe</th>";
    body=rows.map(r=>`<tr><td>${r}</td></tr>`).join("");
  }else{
    // todas as colunas, menos internas e menos as 100% vazias nesta semana
    let cols=Object.keys(rows[0]).filter(c=>!HIDE.has(c));
    cols=cols.filter(c=>rows.some(r=>r[c]!=null && String(r[c]).trim()!==""));
    nCols=cols.length; unica=cols[0];
    head=cols.map(c=>`<th>${lab(c)}</th>`).join("");
    body=rows.map(r=>`<tr>${cols.map(c=>{let v=r[c];
      if(DATECOLS.has(c))v=br(v);
      else if(PCTCOLS.has(c)&&v!=null&&v!=="")v=Math.round(Number(v)*100)+"%";
      return `<td>${v==null?"":v}</td>`;}).join("")}</tr>`).join("");
  }
  // Uma coluna só não é tabela — é lista. Tabela de largura inteira com um
  // cabeçalho "ANIMAL" e uma célula "309" é pior que não mostrar nada.
  if(nCols===1){
    const itens=rows.map(r=>typeof r==="string"?r:(r[unica]??"")).filter(v=>String(v).trim()!=="");
    return `<div class="det"><div class="det-h">${title}<span class="c">${rows.length}</span></div>
      <div class="det-lista">${itens.join(" · ")}</div></div>`;
  }
  return `<div class="det"><div class="det-h">${title}<span class="c">${rows.length}</span></div>
    <div class="det-b"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div></div>`;
}

function findK(p){ for(const s of SECTIONS)for(const k of s.kpis)if(k.p===p)return k; return null; }

function render(){
  const sel=document.getElementById("semana");
  sel.innerHTML=CAL.slice().reverse().map(w=>`<option value="${w.id}" ${w.id===semana?"selected":""}>${br(w.ini)}–${br(w.fim)}</option>`).join("");
  document.getElementById("sections").innerHTML = SECTIONS.map(sec=>{
    // KPI com skip() só aparece nas semanas em que faz sentido (ex.: a safra que
    // começa, que não existia nas semanas congeladas antes da transição)
    const kpis=sec.kpis.filter(k=>!(k.skip && k.skip(snap()))).map(k=>{
      const edited=hasOv(k.p);
      const cls="kpi"+(edited?" edited":"")+(editMode?" editing":"")+(k.manual?" manual":"");
      const tag=k.manual?`<span class="tag" title="dado manual — preencher toda semana (Alexandre / grupo)">manual</span>`:"";
      const lb = (typeof k.l==="function") ? k.l(snap()) : k.l;   // safra vem do dado
      return `<div class="${cls}"><div class="lab">${lb}</div>
        <div class="val" contenteditable="${editMode && !k.html}" data-path="${k.p}">${fmtVal(k)}</div>
        ${tag}<span class="rst" data-path="${k.p}">reset</span></div>`;
    }).join("");
    const det=(sec.det||[]).map(([t,key,opts])=>detTable(t,key,opts)).join("");
    const sub=sec.sub?`<div class="sub">${sec.sub(snap())}</div>`:"";
    const vis=sec.kpis.filter(k=>!(k.skip && k.skip(snap()))).length;
    // teto 7: com a transição de estação a produção tem 7 KPIs, e 6 jogava o
    // último para uma segunda linha, deixando meia faixa vazia
    const cols=sec.wide?Math.min(vis,7):3;   // half=3 (receptoras/headcount iguais)
    return `<div class="panel${sec.wide?' full':''}"><h2><span class="n">${sec.n})</span>${sec.t}</h2>${sub}
      <div class="kpis" style="grid-template-columns:repeat(${cols},minmax(0,1fr))">${kpis}</div>${det}</div>`;
  }).join("");

  if(editMode){
    document.querySelectorAll(".val[contenteditable=true]").forEach(v=>{
      v.addEventListener("blur",e=>{
        const p=e.target.dataset.path, txt=e.target.textContent.trim();
        const k=findK(p); let seed=k&&!k.delta?rawVal(k):null;
        let val=txt==="--"?"":txt; if(val!=="" && !isNaN(val)) val=Number(val);
        if(String(val)===String(seed==null?"":seed)) delOv(p); else setOv(p,val);
        render();
      });
    });
    document.querySelectorAll(".rst").forEach(b=>b.addEventListener("click",e=>{
      delOv(e.target.dataset.path); render();
    }));
  }
  renderPdfs();
}

document.getElementById("semana").addEventListener("change",e=>{semana=e.target.value; saveState(KEY_WK,semana); render();});
document.getElementById("btnEdit").addEventListener("click",e=>{
  editMode=!editMode; e.target.classList.toggle("on",editMode);
  e.target.textContent=editMode?"Concluir edição":"Editar"; render();
});
/* Exportar imagem: PNG do dashboard inteiro, pra mandar no grupo.
   Feito com SVG <foreignObject> + canvas — nada de biblioteca externa, o HTML
   tem que continuar self-contained e funcionando offline. Requisito que isso
   impõe: toda imagem embutida precisa ser data URI (o logo é), senão o canvas
   fica "tainted" e o toBlob falha. */
/* PDF do dashboard: uma página só.
   O motor de PDF é o do navegador (sem biblioteca — o HTML segue self-contained),
   mas o conteúdo é mais alto que uma folha. Então mede-se a altura real e aplica-se
   `zoom` na proporção que cabe em A4 paisagem, antes de chamar print().
   Medir é obrigatório: com fator fixo, semana com poucas linhas sai minúscula e
   semana cheia continua estourando. */
// A4 paisagem a 96dpi (1122x794), menos os 6mm de padding de cada lado.
// A folga de 4% cobre o arredondamento do navegador ao aplicar zoom — sem ela o
// conteúdo encosta no limite e escorrega uma linha para a folha seguinte.
const A4L={largura:1077,altura:749,folga:0.96};
document.getElementById("btnPdf").onclick=()=>{
  const alvo=document.body;
  // mede com a folha de impressão aplicada, senão a conta é da tela e erra feio
  alvo.classList.add("medindo-print");
  const h=alvo.scrollHeight, w=alvo.scrollWidth;
  alvo.classList.remove("medindo-print");
  const z=Math.min(1, (A4L.altura/h)*A4L.folga, (A4L.largura/w)*A4L.folga);
  document.documentElement.style.setProperty("--print-zoom", z.toFixed(3));
  window.print();
};

/* Os PDFs de embriões, embutidos como data URI pelo build — um par por semana.
   A pasta de origem acumula todas as semanas já geradas, então filtra pela
   selecionada e refaz a cada troca (chamado de dentro de render()); sem isso os
   botões de toda semana já processada ficavam empilhados na barra, sem sumir
   quando o filtro mudava. */
function renderPdfs(){
  const cx=document.getElementById("pdfEmb");
  cx.innerHTML="";
  for(const p of (DATA.pdfs||[])){
    if(p.semana!==semana) continue;
    const a=document.createElement("a");
    a.href=p.uri; a.download=p.nome; a.className="btn-pdf";
    // rótulo curto: o nome do arquivo tem a data e não cabe na barra
    a.textContent=/SOCIO|SÓCIO/i.test(p.nome)? "Embriões sócios/vendidos" : "Embriões PG";
    a.title=p.nome;
    cx.appendChild(a);
  }
}

document.getElementById("btnImg").onclick=async(ev)=>{
  const btn=ev.target, rotulo=btn.textContent;
  btn.textContent="Gerando..."; btn.disabled=true;
  try{
    /* Dentro do foreignObject o conteúdo é um <div>, não um <body>: as regras de
       body{} (fonte e cor do texto) não casam e o SVG cai no default do navegador
       — serif e texto preto. Por isso replicamos as declarações do body no
       wrapper da captura. As variáveis de cor não precisam disso: estão no :root,
       que no SVG é o próprio <svg>, e custom property herda pra dentro. */
    const cssCaptura=`#capa{background:var(--bg);color:var(--txt);
      font-family:"Segoe UI",system-ui,-apple-system,Arial,sans-serif;font-size:24px}
      #capa header{padding-bottom:18px}
      #capa .cap-per{font-size:26px;font-weight:700;color:var(--teal);white-space:nowrap}
      #capa .cap-per .rot{display:block;font-size:14px;font-weight:700;color:var(--mut);
        text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px}
      #capa .kpi{min-height:160px;padding:17px 19px 16px}
      #capa .kpi .lab{font-size:16px;font-weight:700;min-height:62px}
      #capa .kpi .tag{font-size:10.5px;top:10px;right:12px}
      #capa .kpi .val{font-size:46px}
      #capa .kpi .val .nota{font-size:17px}
      #capa .panel h2{font-size:24px}
      #capa .panel .sub{font-size:18px}
      #capa .det-h{font-size:17px;font-weight:800}
      #capa table{font-size:19px}
      #capa th{font-size:15px;font-weight:700}
      #capa td,#capa th{padding:11px 14px}`;
    const css=[...document.querySelectorAll("style")].map(s=>s.textContent).join("\n")+cssCaptura;
    const clone=document.createElement("div");
    clone.id="capa";
    clone.appendChild(document.querySelector("header").cloneNode(true));
    clone.appendChild(document.querySelector(".wrap").cloneNode(true));
    // a toolbar sai da imagem, mas o período tem que ficar: sem ele a imagem
    // solta no grupo não diz de que semana é.
    const w=CAL.find(x=>x.id===semana);
    const per=document.createElement("div");
    per.className="cap-per";
    per.innerHTML=`<span class="rot">Semana de referência</span>`+
      (w? `${br(w.ini)} a ${br(w.fim)}` : (semana||""));
    clone.querySelector(".toolbar")?.replaceWith(per);
    clone.querySelectorAll(".det-b").forEach(d=>{           // tabela inteira, sem scroll
      d.style.maxHeight="none"; d.style.overflow="visible";
    });
    clone.querySelectorAll(".rst").forEach(r=>r.remove());
    const largura=document.querySelector(".wrap").scrollWidth;
    // mede a altura real renderizando fora da tela
    const medidor=document.createElement("div");
    medidor.style.cssText=`position:fixed;left:-99999px;top:0;width:${largura}px`;
    medidor.appendChild(clone.cloneNode(true));
    document.body.appendChild(medidor);
    const altura=medidor.firstChild.scrollHeight+40;
    document.body.removeChild(medidor);

    // o XMLSerializer já emite xmlns="http://www.w3.org/1999/xhtml" na raiz;
    // acrescentar de novo dá "Attribute xmlns redefined" e o SVG não renderiza.
    const xhtml=new XMLSerializer().serializeToString(clone);
    const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${largura}" height="${altura}">`
      +`<rect width="100%" height="100%" fill="#04223B"/>`
      +`<foreignObject width="100%" height="100%">`
      +`<style xmlns="http://www.w3.org/1999/xhtml">${css}</style>${xhtml}</foreignObject></svg>`;

    const escala=2;                                          // 2x = legível no celular
    const img=new Image();
    await new Promise((ok,erro)=>{
      img.onload=ok; img.onerror=()=>erro(new Error("falha ao renderizar o SVG"));
      img.src="data:image/svg+xml;charset=utf-8,"+encodeURIComponent(svg);
    });
    const cv=document.createElement("canvas");
    cv.width=largura*escala; cv.height=altura*escala;
    const ctx=cv.getContext("2d");
    ctx.fillStyle="#04223B"; ctx.fillRect(0,0,cv.width,cv.height);
    ctx.scale(escala,escala); ctx.drawImage(img,0,0);
    const blob=await new Promise(r=>cv.toBlob(r,"image/png"));
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download=`atualizacao_semanal_${semana||"export"}.png`;
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),5000);
  }catch(e){
    alert("Não consegui gerar a imagem: "+e.message+
          "\nAlternativa: Ctrl+P e salvar como PDF.");
  }finally{ btn.textContent=rotulo; btn.disabled=false; }
};

render();
</script>
</body></html>
"""


_RE_PDF_DATA = re.compile(r"(\d{2})-(\d{2})-(\d{4})\.pdf$")


def _pdfs_embrioes(data: dict) -> list:
    """Os PDFs de embriões, um par por semana, como data URI.

    Gerados por tools/build_pdf_embrioes.py em _cache/pdf — a pasta ACUMULA (não
    apaga o PDF da semana passada), então achado em 28/08/2026: sem a tag de
    semana, o seletor do dash ficava com os botões de TODAS as semanas empilhados,
    sem filtrar pela que estava selecionada. Cada entrada carrega 'semana' (ISO,
    tirado do nome do arquivo) — o JS filtra por isso e refaz os botões a cada
    troca de semana, igual ao resto do dash.

    Embutidos como data URI porque o dashboard é servido por srcdoc a partir de um
    bucket privado: link para arquivo externo não chegaria a lugar nenhum."""
    import base64

    pasta = BASE_DIR / "_cache" / "pdf"
    if not pasta.exists():
        return []
    out = []
    for f in sorted(pasta.glob("*.pdf")):
        m = _RE_PDF_DATA.search(f.name)
        if not m:
            print(f"  [pdf] {f.name}: nome sem data DD-MM-AAAA, não entra no dash")
            continue
        dd, mm, aaaa = m.groups()
        semana = f"{aaaa}-{mm}-{dd}"
        b64 = base64.b64encode(f.read_bytes()).decode()
        out.append({"nome": f.name, "semana": semana,
                    "uri": "data:application/pdf;base64," + b64,
                    "kb": len(b64) // 1024})
    return out


def build():
    if not JSON_IN.exists():
        raise SystemExit(f"Rode PGSemanalReport.py primeiro — falta {JSON_IN.name}")
    data = json.loads(JSON_IN.read_text(encoding="utf-8"))
    data["pdfs"] = _pdfs_embrioes(data)
    if data["pdfs"]:
        print("  [pdf] embutidos: " + ", ".join(
            f'{p["nome"]} ({p["kb"]} KB)' for p in data["pdfs"]))
    html = TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    HTML_OUT.parent.mkdir(parents=True, exist_ok=True)
    HTML_OUT.write_text(html, encoding="utf-8")
    print(f"-> {HTML_OUT.name} gravado ({HTML_OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    build()
