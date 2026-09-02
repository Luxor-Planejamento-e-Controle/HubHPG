/* Hub HPG — casca web. Mesmo desenho do LuxorP&CHub (sidebar + rotas por hash),
   com a identidade do Haras Pao Grande.

   Uma aba real: Atualização Semanal, o dashboard que o pipeline já gera.
   Comitê Mensal e Plantel/Movimentação são PLACEHOLDER — estão na navegação pra
   marcar que também vêm pra cá, mas não desenham tela nenhuma: não existe painel
   web dos dois hoje (o comitê sai do ComiteHPG.pbix, o plantel de e-mail com
   xlsx), e sem uma referência do que mostrar a casca não inventa layout.

   Pra promover um placeholder a aba de verdade: trocar `soon:true` pela função
   de render e, se tiver gráfico, copiar o vendor/echarts.min.js do LuxorP&CHub
   e voltar a tag <script> no index.html. */
'use strict';

/* ---- rotas ---- */
const ICON = {
  home:'M3 11l9-8 9 8M5 10v10h5v-6h4v6h5V10',
  semanal:'M3 5h18v16H3zM3 9h18M8 3v4M16 3v4M8 14h3M8 17h6',
  comite:'M4 20V10M10 20V4M16 20v-7M22 20H2',
  plantel:'M4 20V8l8-5 8 5v12M9 20v-6h6v6',
};
const ROUTES = [
  {id:'', title:'Início', sub:'Hub do Haras Pao Grande', icon:'home', render:renderHome},
  {id:'semanal', title:'Atualização Semanal', sub:'Fechamento da semana — plantel, produção e receptoras', icon:'semanal', render:renderSemanal},
  {id:'comite', title:'Comitê Mensal', sub:'Relatório de desempenho estratégico — deck do mês', icon:'comite', render:renderComite},
  {id:'plantel', title:'Plantel / Movimentação', sub:'Cascata de valor e movimentos do mês', icon:'plantel', soon:true,
   fonte:'mov_cascata / mov_detalhe do LuxorMonthlyP-CRoutines/PlantelHPG (LxMovimentacao.py)',
   hoje:'Hoje sai por e-mail com o xlsx anexo (LxEmailHPGPlantel.py).'},
];
function allowed(){
  const ok=(window.HUB&&window.HUB.dashboards)||[];
  return ROUTES.filter(r=>!r.id||ok.includes(r.id));
}
const byId = id => allowed().find(r=>r.id===id) || ROUTES[0];

function buildNav(){
  const nav=document.getElementById('nav'); nav.innerHTML='';
  for(const r of allowed()){
    const a=document.createElement('a'); a.href='#/'+r.id; a.className=r.soon?'locked':'';
    a.innerHTML=`<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg>`
      +`<span>${r.title}</span>`+(r.soon?'<span class="badge">em breve</span>':'');
    nav.appendChild(a);
  }
}
function router(){
  const id=(location.hash.replace(/^#\/?/,'')||'');
  const r=byId(id);
  document.getElementById('pageTitle').textContent=r.title;
  document.getElementById('pageSub').textContent=r.sub;
  document.querySelectorAll('#nav a').forEach(a=>a.classList.toggle('active',a.getAttribute('href')==='#/'+id));
  const c=document.getElementById('content'); c.className='content'; c.innerHTML='';
  (r.render||renderPlaceholder)(c,r); window.scrollTo(0,0);
}

/* ---- Início ---- */
function renderHome(el){
  const cards=allowed().filter(r=>r.id).map(r=>`
    <a class="card hover" href="#/${r.id}">
      <div class="card-title"><svg class="ico" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#CA9703" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg><h3 style="margin:0">${r.title}</h3>
        ${r.soon?'<span class="pill soon" style="margin-left:auto">em breve</span>':''}</div>
      <div class="desc">${r.sub}</div></a>`).join('');
  el.innerHTML=`<div class="hero"><h1>Haras Pao Grande</h1>
    <p>Hub dos painéis do haras.</p></div>
    <div class="grid g-3">${cards}</div>`;
}

/* ---- placeholder ----
   Aba que ainda não existe. Mostra a fonte que vai alimentar e de onde o dado
   sai hoje — nada de gráfico de mentira nem número de exemplo, que é o tipo de
   coisa que depois alguém lê como se fosse o painel. */
function renderPlaceholder(el,r){
  el.innerHTML=`
    <div class="card" style="max-width:720px">
      <div class="card-title">
        <svg class="ico" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#CA9703" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg>
        <h2>${r.title}</h2><span class="pill soon" style="margin-left:auto">em breve</span>
      </div>
      <p style="margin:0 0 18px;color:var(--ink-2)">${r.sub}.</p>
      <div class="grid g-2" style="gap:12px">
        <div><div class="label" style="font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-3);margin-bottom:5px">Base que vai alimentar</div>
          <div style="font-size:13px">${r.fonte}</div></div>
        <div><div class="label" style="font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-3);margin-bottom:5px">Como é hoje</div>
          <div style="font-size:13px">${r.hoje}</div></div>
      </div>
      <div class="banner" style="margin:20px 0 0">O layout ainda não está definido — a tela entra quando a
        referência do que mostrar estiver fechada.</div>
    </div>`;
}

/* ---- Atualização Semanal ----
   O dashboard_semanal.html já é autocontido e já usa a paleta do haras — não há
   re-skin, só o embed. O build (hub/tools/build_semanal.py) copia a saída do
   pipeline e esconde o cabeçalho interno, que duplicaria o título da aba.
   Na fase gold o HTML vem do bucket privado e entra por `srcdoc`.

   Altura do iframe acompanha o conteúdo (classe .auto-h) — achado em 28/08/2026:
   com altura fixa (100vh), o dashboard rolava por dentro do iframe E a página do
   hub rolava por fora, duas barras pro mesmo conteúdo. srcdoc é mesma origem, dá
   pra ler contentDocument e medir. Reajusta de novo quando o conteúdo muda de
   altura (troca de semana, modo de edição), via ResizeObserver no body de dentro. */
function _ajustaAlturaEmbed(f){
  const doc=f.contentDocument;
  if(!doc||!doc.documentElement) return;
  const ajusta=()=>{
    const h=Math.max(doc.documentElement.scrollHeight, doc.body?doc.body.scrollHeight:0);
    f.style.setProperty('--embed-h', h+'px');
  };
  ajusta();
  try{ new ResizeObserver(ajusta).observe(doc.body||doc.documentElement); }
  catch(e){ /* sem ResizeObserver, fica só na medida inicial + no load */ }
}

function renderSemanal(el){
  el.classList.add('flush');
  const html=window.HUB&&window.HUB.semanalHtml;
  if(!html){
    const f=document.createElement('iframe');
    f.className='embed auto-h'; f.title='Atualização Semanal';
    f.addEventListener('load',()=>_ajustaAlturaEmbed(f));
    f.src='assets/semanal/dashboard.html';
    el.appendChild(f);
    return;
  }
  const f=document.createElement('iframe');
  f.className='embed auto-h'; f.title='Atualização Semanal';
  f.addEventListener('load',()=>_ajustaAlturaEmbed(f));
  f.srcdoc=html;
  el.appendChild(f);
}

/* ---- Comitê Mensal ----
   Deck próprio (comite.html): 16:9, navegação por seta, modo apresentação e
   export PPTX. Entra em iframe porque é uma apresentação, não uma página do
   hub — dentro dele a topbar e a sidebar só atrapalhariam. */
function renderComite(el){
  el.classList.add('flush');
  el.innerHTML=`<iframe class="embed" src="comite.html" title="Comitê Mensal"></iframe>`;
}

/* ---- sidebar: desktop colapsa pra ícone, mobile vira drawer ----
   Mesmo botão (#collapseBtn), comportamento diferente por breakpoint: em
   ≤760px a sidebar sai do grid (CSS) e vira overlay fixo, então "colapsar"
   não faz sentido — o botão abre/fecha o drawer por cima do conteúdo, com um
   scrim pra fechar tocando fora e fechamento automático ao navegar. Mesma
   receita do LuxorP&CHub — achado em 31/08/2026 que esta casca (HUB HPG) é um
   site PRÓPRIO, separado, que nunca tinha recebido o fix (só o dashboard
   embutido dentro dela tinha sido corrigido). */
(function collapse(){
  const btn=document.getElementById('collapseBtn'), app=document.querySelector('.app');
  const mq=window.matchMedia('(max-width:760px)');
  const scrim=document.createElement('div');
  scrim.className='nav-scrim';
  app.appendChild(scrim);
  const closeDrawer=()=>{app.classList.remove('nav-open');document.body.style.overflow='';};
  const openDrawer=()=>{app.classList.add('nav-open');document.body.style.overflow='hidden';};
  const collapsedSalvo=()=>localStorage.getItem('hpg-collapsed')==='1';
  if(!mq.matches && collapsedSalvo())app.classList.add('collapsed');
  btn.onclick=()=>{
    if(mq.matches){
      app.classList.contains('nav-open')?closeDrawer():openDrawer();
    }else{
      app.classList.toggle('collapsed');
      localStorage.setItem('hpg-collapsed',app.classList.contains('collapsed')?'1':'0');
    }
  };
  scrim.onclick=closeDrawer;
  document.getElementById('navCloseBtn').onclick=closeDrawer;
  // #nav é recriado a cada buildNav() — delegação no container pega os links novos também
  document.getElementById('nav').addEventListener('click',e=>{if(e.target.closest('a'))closeDrawer();});
  mq.addEventListener('change',e=>{
    closeDrawer();
    app.classList.toggle('collapsed',!e.matches && collapsedSalvo());
  });
})();

/* ---- boot ----
   Offline sobe direto. Quando entrar o porteiro (assets/auth.js, fase gold),
   é ele quem chama hubBoot() depois da sessão — igual ao LuxorP&CHub. */
let booted=false;
window.hubBoot=function(){
  document.body.classList.add('hub-ready');
  if(booted){buildNav();router();return;}
  booted=true;
  const chip=document.getElementById('userChip'), out=document.getElementById('signOut');
  if(window.HUB&&window.HUB.email){
    const nome=window.HUB.nome||window.HUB.email.split('@')[0];
    document.getElementById('userName').textContent=nome;
    document.getElementById('userAvatar').textContent=
      nome.split(/[\s.]+/).slice(0,2).map(s=>s[0]||'').join('').toUpperCase();
    chip.hidden=false; out.hidden=false; out.onclick=window.hubSignOut;
  }
  buildNav();
  window.addEventListener('hashchange',router);
  router();
};
if(window.HUB&&window.HUB.offline) window.hubBoot();
