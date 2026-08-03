/* Hub HPG — casca web. Mesmo desenho do LuxorP&CHub (sidebar + rotas por hash),
   com a identidade do Haras Pao Grande.

   Hoje tem UMA aba: Atualização Semanal, o dashboard que o pipeline já gera.
   Comitê Mensal e Plantel/Movimentação entram quando a referência deles estiver
   definida — hoje não existe painel web nenhum pros dois, e a casca não inventa
   tela. Pra ligar uma aba nova: entrada em ROUTES + ICON + render, e o
   vendor/echarts.min.js de volta no index.html (é o mesmo arquivo do P&C Hub). */
'use strict';

/* ---- rotas ---- */
const ICON = {
  home:'M3 11l9-8 9 8M5 10v10h5v-6h4v6h5V10',
  semanal:'M3 5h18v16H3zM3 9h18M8 3v4M16 3v4M8 14h3M8 17h6',
};
const ROUTES = [
  {id:'', title:'Início', sub:'Hub do Haras Pao Grande', icon:'home', render:renderHome},
  {id:'semanal', title:'Atualização Semanal', sub:'Fechamento da semana — plantel, produção e receptoras', icon:'semanal', render:renderSemanal},
];
function allowed(){
  const ok=(window.HUB&&window.HUB.dashboards)||[];
  return ROUTES.filter(r=>!r.id||ok.includes(r.id));
}
const byId = id => allowed().find(r=>r.id===id) || ROUTES[0];

function buildNav(){
  const nav=document.getElementById('nav'); nav.innerHTML='';
  for(const r of allowed()){
    const a=document.createElement('a'); a.href='#/'+r.id;
    a.innerHTML=`<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg><span>${r.title}</span>`;
    nav.appendChild(a);
  }
}
function router(){
  const id=(location.hash.replace(/^#\/?/,'')||'');
  const r=byId(id);
  document.getElementById('pageTitle').textContent=r.title;
  document.getElementById('pageSub').textContent=r.sub;
  document.querySelectorAll('#nav a').forEach(a=>a.classList.toggle('active',a.getAttribute('href')==='#/'+id));
  const c=document.getElementById('content'); c.className='content'; c.innerHTML=''; r.render(c); window.scrollTo(0,0);
}

/* ---- Início ---- */
function renderHome(el){
  const cards=allowed().filter(r=>r.id).map(r=>`
    <a class="card hover" href="#/${r.id}">
      <div class="card-title"><svg class="ico" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#CA9703" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg><h3 style="margin:0">${r.title}</h3></div>
      <div class="desc">${r.sub}</div></a>`).join('');
  el.innerHTML=`<div class="hero"><h1>Haras Pao Grande</h1>
    <p>Hub dos painéis do haras.</p></div>
    <div class="grid g-3">${cards}</div>`;
}

/* ---- Atualização Semanal ----
   O dashboard_semanal.html já é autocontido e já usa a paleta do haras — não há
   re-skin, só o embed. O build (hub/tools/build_semanal.py) copia a saída do
   pipeline e esconde o cabeçalho interno, que duplicaria o título da aba.
   Na fase gold o HTML vem do bucket privado e entra por `srcdoc`. */
function renderSemanal(el){
  el.classList.add('flush');
  const html=window.HUB&&window.HUB.semanalHtml;
  if(!html){
    el.innerHTML=`<iframe class="embed" src="assets/semanal/dashboard.html" title="Atualização Semanal"></iframe>`;
    return;
  }
  const f=document.createElement('iframe');
  f.className='embed'; f.title='Atualização Semanal'; f.srcdoc=html;
  el.appendChild(f);
}

/* ---- sidebar recolhível ---- */
(function collapse(){
  const btn=document.getElementById('collapseBtn'), app=document.querySelector('.app');
  if(localStorage.getItem('hpg-collapsed')==='1')app.classList.add('collapsed');
  btn.onclick=()=>{app.classList.toggle('collapsed');
    localStorage.setItem('hpg-collapsed',app.classList.contains('collapsed')?'1':'0');};
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
