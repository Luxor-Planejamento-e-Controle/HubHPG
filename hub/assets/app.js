/* Hub HPG — casca web. Mesmo desenho do LuxorP&CHub (sidebar + rotas por hash +
   ECharts), com a identidade do Haras Pao Grande e as abas do haras.

   Abas:
     semanal  — dashboard_semanal.html do pipeline, embutido em iframe
     comite   — base_bi.parquet (comitê mensal)
     plantel  — mov_cascata / mov_detalhe / base_plantel (PlantelHPG) */
'use strict';

const C = {
  amber:'#CA9703', amberDeep:'#A87D02', teal:'#7FA8C4', tealDeep:'#4E7FA0',
  ink:'#EAF0F4', ink2:'#B6C8D6', ink3:'#93AABC', pos:'#4CC38A', neg:'#F07A7A',
  warn:'#E0B341', line:'rgba(255,255,255,.12)'
};
/* série categórica: dourado e azul do haras primeiro, o resto derivado deles */
const SERIE = ['#CA9703','#7FA8C4','#4CC38A','#E0B341','#4E7FA0','#F07A7A',
               '#A87D02','#B6C8D6','#2A7F62','#8E6BA8','#C98A3F','#5D93B8'];

const MES_ABR = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
const fmt = {
  int:v=>v==null?'—':Math.round(v).toLocaleString('pt-BR'),
  num:(v,d=2)=>v==null?'—':v.toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}),
  pct:v=>v==null?'—':(v>=0?'+':'')+v.toFixed(1).replace('.',',')+'%',
  mi:v=>v==null?'—':(v/1e6).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+' Mi',
  rs:v=>v==null?'—':'R$ '+v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}),
  rs0:v=>v==null?'—':'R$ '+v.toLocaleString('pt-BR',{maximumFractionDigits:0}),
  br:d=>{if(!d)return'—';const[y,m,dd]=d.split('-');return dd+'/'+m+'/'+y;},
  mesano:d=>{if(!d)return'—';const[y,m]=d.split('-');return MES_ABR[+m-1]+'/'+y.slice(2);}
};
const cls = v => v==null?'':v>=0?'pos':'neg';
const kpi = (l,v,s,c)=>`<div class="card kpi"><div class="label">${l}</div><div class="val ${c||''}">${v}</div><div class="delta ${c||''}">${s||'&nbsp;'}</div></div>`;

/* ---- ECharts base ---- */
function baseOpt(){return {
  backgroundColor:'transparent',
  color:SERIE,
  textStyle:{fontFamily:'Segoe UI, system-ui, sans-serif',color:C.ink2},
  grid:{left:72,right:24,top:34,bottom:56},
  tooltip:{trigger:'axis',backgroundColor:'#072B49',borderColor:C.line,textStyle:{color:C.ink},
    axisPointer:{lineStyle:{color:C.ink3}}},
  legend:{textStyle:{color:C.ink2},top:2,icon:'roundRect',itemWidth:12,itemHeight:12},
};}
/* Arthur não quer linha de grade nos gráficos — splitLine desligado por padrão. */
const axis = extra => Object.assign({axisLine:{lineStyle:{color:C.line}},axisLabel:{color:C.ink3},
  splitLine:{show:false},axisTick:{show:false}},extra||{});
function zoom(start,end){return [
  {type:'slider',height:20,bottom:14,borderColor:C.line,fillerColor:'rgba(202,151,3,.16)',
   start:start==null?0:start,end:end==null?100:end,
   handleStyle:{color:C.amber},moveHandleStyle:{color:C.amber},
   dataBackground:{lineStyle:{color:C.ink3},areaStyle:{color:'rgba(182,200,214,.15)'}},
   selectedDataBackground:{lineStyle:{color:C.amber},areaStyle:{color:'rgba(202,151,3,.25)'}},
   textStyle:{color:C.ink3},labelFormatter:''},
];}
const charts=[];
function mkChart(el,opt){if(!el)return null;const c=echarts.init(el,null,{renderer:'canvas'});c.setOption(opt);charts.push(c);return c;}
window.addEventListener('resize',()=>charts.forEach(c=>c.resize()));
function clearCharts(){while(charts.length)charts.pop().dispose();}

/* ---- helpers ---- */
/* Converte {cols,rows} (formato compacto que o build_data grava) em objetos.
   O JSON fica pequeno porque não repete o nome da coluna em cada linha. */
function objify(t){ return t.rows.map(r=>{const o={};t.cols.forEach((c,i)=>o[c]=r[i]);return o;}); }
const uniq = (arr,k)=>[...new Set(arr.map(x=>x[k]).filter(v=>v!=null&&v!==''))].sort();
function selOpts(list,val,todos){
  return (todos?[`<option value="">${todos}</option>`]:[])
    .concat(list.map(v=>`<option ${v===val?'selected':''}>${v}</option>`)).join('');
}
function seg(id,opts,val){
  return `<div class="seg" id="${id}">`+opts.map(o=>`<button data-v="${o}" class="${o===val?'on':''}">${o}</button>`).join('')+`</div>`;
}
function bindSeg(id,cb){const g=document.getElementById(id);if(!g)return;
  g.onclick=e=>{const b=e.target.closest('button');if(!b)return;
    g.querySelectorAll('button').forEach(x=>x.classList.toggle('on',x===b));cb();};}
const segVal=id=>{const b=document.querySelector('#'+id+' button.on');return b?b.dataset.v:null;};

/* Tabela ordenável. specs = [[rótulo, chave, formatador, alinhaEsquerda?]] */
function sortTable(host, specs, rows, opts){
  opts = opts||{};
  let sk = opts.sort || specs[0][1], sd = opts.dir || 1;
  const head = () => specs.map(s=>`<th data-k="${s[1]}">${s[0]}${s[1]===sk?`<span class="sort">${sd>0?'▲':'▼'}</span>`:''}</th>`).join('');
  const draw = () => {
    const sorted=[...rows].sort((a,b)=>{
      const x=a[sk], y=b[sk];
      if(x==null&&y==null)return 0; if(x==null)return 1; if(y==null)return -1;
      return (typeof x==='number'&&typeof y==='number') ? (x-y)*sd : String(x).localeCompare(String(y),'pt-BR')*sd;
    });
    host.innerHTML=`<div class="tbl-wrap" style="max-height:${opts.maxH||460}px"><table class="data">
      <thead><tr>${head()}</tr></thead>
      <tbody>${sorted.map(r=>`<tr>${specs.map(s=>`<td>${s[2]?s[2](r[s[1]],r):(r[s[1]]==null?'—':r[s[1]])}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
    host.querySelectorAll('th').forEach(th=>th.onclick=()=>{
      const k=th.dataset.k; if(k===sk) sd=-sd; else {sk=k; sd=1;} draw();
    });
  };
  draw();
}

/* ---- rotas ---- */
const ICON = {
  home:'M3 11l9-8 9 8M5 10v10h5v-6h4v6h5V10',
  semanal:'M3 5h18v16H3zM3 9h18M8 3v4M16 3v4M8 14h3M8 17h6',
  comite:'M4 20V10M10 20V4M16 20v-7M22 20H2',
  plantel:'M4 20V8l8-5 8 5v12M9 20v-6h6v6'
};
const ROUTES = [
  {id:'', title:'Início', sub:'Hub do Haras Pao Grande', icon:'home', render:renderHome},
  {id:'semanal', title:'Atualização Semanal', sub:'Fechamento da semana — plantel, produção e receptoras', icon:'semanal', render:renderSemanal},
  {id:'comite', title:'Comitê Mensal', sub:'Plantel e patrimônio por mês de referência', icon:'comite', render:renderComite},
  {id:'plantel', title:'Plantel / Movimentação', sub:'Cascata de valor e movimentos do mês', icon:'plantel', render:renderPlantel},
];
/* Aba só entra na nav se o usuário pode ver E o dado dela chegou — ninguém abre
   painel vazio porque o build ainda não rodou. Offline, o semanal é o iframe
   local, então só some se o build_semanal.py nunca tiver rodado. */
function temDado(id){
  if(id==='comite')  return !!window.COMITE_DATA;
  if(id==='plantel') return !!window.PLANTEL_DATA;
  return true;
}
function allowed(){
  const ok=(window.HUB&&window.HUB.dashboards)||[];
  return ROUTES.filter(r=>!r.id||(ok.includes(r.id)&&temDado(r.id)));
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
  clearCharts();
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
      <div class="card-title"><svg class="ico" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="${C.amber}" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${ICON[r.icon]}"/></svg><h3 style="margin:0">${r.title}</h3></div>
      <div class="desc">${r.sub}</div></a>`).join('');
  const faltando=ROUTES.filter(r=>r.id&&!allowed().some(a=>a.id===r.id));
  el.innerHTML=`<div class="hero"><h1>Haras Pao Grande</h1>
    <p>Hub dos painéis do haras: o fechamento semanal que já roda no pipeline, o comitê
    mensal e a movimentação do plantel, num lugar só.</p></div>
    <div class="grid g-3">${cards}</div>
    ${faltando.length?`<div class="banner" style="margin-top:22px">Sem dado carregado para:
      ${faltando.map(r=>r.title).join(', ')}. Rode <code>python hub/tools/build_data.py</code>.</div>`:''}`;
}

/* ---- Atualização Semanal ----
   O dashboard_semanal.html já é autocontido e já usa a paleta do haras — não há
   re-skin, só o embed. O build (hub/tools/build_semanal.py) copia a saída do
   pipeline e esconde o cabeçalho interno, que duplicaria o título da aba. */
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

/* ---- Comitê Mensal (base_bi.parquet) ---- */
function renderComite(el){
  const D=window.COMITE_DATA;
  if(!D){el.innerHTML='<div class="empty">Dados não carregados. Rode <code>python hub/tools/build_data.py</code>.</div>';return;}
  const all=objify(D.tabela);
  const meses=D.meses, mesDef=meses[meses.length-1];
  const locais=uniq(all,'local'), cats=uniq(all,'categoria'), status=uniq(all,'status');
  const SLIDE=['Tudo','Slide 1','Slide 2'];

  el.innerHTML=`
    <div class="toolbar">
      <div class="field"><label>Mês de referência</label>
        <select id="mes">${meses.map(m=>`<option value="${m}" ${m===mesDef?'selected':''}>${fmt.mesano(m)}</option>`).join('')}</select></div>
      <div class="field"><label>Local</label><select id="loc">${selOpts(locais,'','Todos')}</select></div>
      <div class="field"><label>Categoria</label><select id="cat">${selOpts(cats,'','Todas')}</select></div>
      <div class="field"><label>Status no plantel</label><select id="st">${selOpts(status,'','Todos')}</select></div>
      <div class="field"><label>Recorte do comitê</label>${seg('slide',SLIDE,'Tudo')}</div>
    </div>
    <div class="grid g-5" id="kpis" style="margin-bottom:16px"></div>
    <div class="grid g-2">
      <div class="card"><div class="card-title"><h2>Patrimônio proporcional por mês</h2><span class="muted" id="evoSub"></span></div>
        <div id="evo" class="chart"></div></div>
      <div class="card"><div class="card-title"><h2>Patrimônio por local</h2><span class="muted" id="locSub"></span></div>
        <div id="donut" class="chart"></div></div>
    </div>
    <div class="card" style="margin-top:16px"><div class="card-title"><h2>Animais por categoria</h2><span class="muted">contagem no mês</span></div>
      <div id="barCat" class="chart"></div></div>
    <div class="card" style="margin-top:16px"><div class="card-title"><h2>Plantel do mês</h2><span class="muted" id="tblSub"></span></div>
      <div id="tbl"></div></div>`;

  const filtros = () => ({
    mes:document.getElementById('mes').value,
    loc:document.getElementById('loc').value,
    cat:document.getElementById('cat').value,
    st:document.getElementById('st').value,
    slide:segVal('slide')
  });
  /* filtro sem o mês: usado na série histórica, senão a linha teria 1 ponto só */
  const match = (r,f) => (!f.loc||r.local===f.loc) && (!f.cat||r.categoria===f.cat)
    && (!f.st||r.status===f.st)
    && (f.slide==='Tudo' || (f.slide==='Slide 1'?r.s1:r.s2));

  const draw=()=>{
    clearCharts();
    const f=filtros();
    const serie=all.filter(r=>match(r,f));
    const mes=serie.filter(r=>r.mes===f.mes);
    const iAtual=meses.indexOf(f.mes);
    const ant=iAtual>0?serie.filter(r=>r.mes===meses[iAtual-1]):[];
    const soma=(a,k)=>a.reduce((s,r)=>s+(r[k]||0),0);

    const pat=soma(mes,'patrimonio'), patAnt=soma(ant,'patrimonio');
    const dPat=ant.length?pat-patAnt:null;
    const dPct=(ant.length&&patAnt)?(pat/patAnt-1)*100:null;
    document.getElementById('kpis').innerHTML=[
      kpi('Animais',fmt.int(mes.length),ant.length?`${mes.length-ant.length>=0?'+':''}${mes.length-ant.length} vs ${fmt.mesano(meses[iAtual-1])}`:'',
          ant.length?cls(mes.length-ant.length):''),
      kpi('Patrimônio proporcional',fmt.mi(pat),fmt.rs0(pat),''),
      kpi('Δ mês',dPat==null?'—':fmt.mi(dPat),fmt.pct(dPct),cls(dPat)),
      kpi('Valor 100%',fmt.mi(soma(mes,'valor_100')),fmt.rs0(soma(mes,'valor_100')),''),
      kpi('Cotas',fmt.num(soma(mes,'cotas'),2),`${mes.filter(r=>r.cotas===1).length} integrais`,''),
    ].join('');

    /* evolução: patrimônio por mês, com os mesmos filtros (menos o mês) */
    const porMes=meses.map(m=>soma(serie.filter(r=>r.mes===m),'patrimonio'));
    document.getElementById('evoSub').textContent=`${fmt.mesano(meses[0])} → ${fmt.mesano(meses[meses.length-1])}`;
    const evo=mkChart(document.getElementById('evo'),Object.assign(baseOpt(),{
      legend:{show:false},grid:{left:72,right:20,top:20,bottom:52},
      tooltip:Object.assign(baseOpt().tooltip,{valueFormatter:v=>fmt.rs(v)}),
      xAxis:axis({type:'category',data:meses,boundaryGap:false,axisLabel:{color:C.ink3,formatter:fmt.mesano}}),
      yAxis:axis({type:'value',scale:true,axisLabel:{color:C.ink3,formatter:v=>(v/1e6).toFixed(0)+' Mi'}}),
      series:[{name:'Patrimônio',type:'line',smooth:true,symbol:'none',data:porMes,
        lineStyle:{color:C.amber,width:2.4},
        areaStyle:{color:new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:'rgba(202,151,3,.28)'},{offset:1,color:'rgba(202,151,3,0)'}])},
        markLine:{silent:true,symbol:'none',data:[{xAxis:f.mes}],
          lineStyle:{color:C.teal,type:'dashed'},label:{show:false}}}]
    }));
    if(evo) evo.on('click',p=>{if(p.name){document.getElementById('mes').value=p.name;draw();}});

    /* patrimônio por local no mês */
    const porLocal={};
    mes.forEach(r=>{const k=r.local||'(sem local)';porLocal[k]=(porLocal[k]||0)+(r.patrimonio||0);});
    const dl=Object.entries(porLocal).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
    document.getElementById('locSub').textContent=fmt.mesano(f.mes);
    mkChart(document.getElementById('donut'),Object.assign(baseOpt(),{
      grid:null,legend:{textStyle:{color:C.ink2},bottom:0,icon:'roundRect',itemWidth:11,itemHeight:11,type:'scroll'},
      tooltip:{trigger:'item',backgroundColor:'#072B49',borderColor:C.line,textStyle:{color:C.ink},
        formatter:p=>`${p.name}<br><b>${fmt.rs(p.value)}</b> · ${p.percent}%`},
      series:[{type:'pie',radius:['46%','72%'],center:['50%','44%'],avoidLabelOverlap:true,
        itemStyle:{borderColor:'#0A3050',borderWidth:2},
        label:{color:C.ink2,formatter:p=>`${p.percent}%`},
        data:dl.map(([n,v])=>({name:n,value:v}))}]
    }));

    /* contagem por categoria no mês */
    const porCat={};
    mes.forEach(r=>{const k=r.categoria||'(sem categoria)';porCat[k]=(porCat[k]||0)+1;});
    const dc=Object.entries(porCat).sort((a,b)=>b[1]-a[1]);
    mkChart(document.getElementById('barCat'),Object.assign(baseOpt(),{
      legend:{show:false},grid:{left:72,right:24,top:18,bottom:70},
      tooltip:Object.assign(baseOpt().tooltip,{valueFormatter:v=>fmt.int(v)+' animais'}),
      xAxis:axis({type:'category',data:dc.map(d=>d[0]),axisLabel:{color:C.ink3,rotate:32,fontSize:11}}),
      yAxis:axis({type:'value',axisLabel:{color:C.ink3}}),
      series:[{type:'bar',data:dc.map(d=>d[1]),itemStyle:{color:C.amber,borderRadius:[3,3,0,0]},
        barMaxWidth:44,label:{show:true,position:'top',color:C.ink3,formatter:p=>fmt.int(p.value)}}]
    }));

    document.getElementById('tblSub').textContent=`${mes.length} animais · ${fmt.mesano(f.mes)}`;
    sortTable(document.getElementById('tbl'),[
      ['Animal','nome',v=>v||'—'],
      ['Categoria','categoria'],
      ['Sexo','sexo',v=>v||'—'],
      ['Local','local',v=>v||'—'],
      ['Status','status',v=>v||'—'],
      ['Nota','nota',v=>v||'—'],
      ['Idade','idade',v=>v==null?'—':fmt.num(v,1)],
      ['Cotas','cotas',v=>v==null?'—':fmt.num(v,4)],
      ['Valor 100%','valor_100',fmt.rs0],
      ['Patrimônio','patrimonio',fmt.rs0],
    ],mes,{sort:'patrimonio',dir:-1,maxH:520});
  };

  ['mes','loc','cat','st'].forEach(id=>document.getElementById(id).onchange=draw);
  bindSeg('slide',draw);
  draw();
}

/* ---- Plantel / Movimentação (PlantelHPG) ---- */
/* Cascata: barra invisível de apoio + a variação por cima, no mesmo stack.
   ECharts não tem waterfall nativo; esse é o jeito canônico. */
const MOV = [
  ['compra','Compra'], ['producao','Produção'], ['venda','Venda'], ['morte','Morte'],
  ['doacao','Doação'], ['reaval','Reavaliação'], ['transferencia','Transferência'],
  ['saiu_controle','Saiu do controle'],
];
function renderPlantel(el){
  const D=window.PLANTEL_DATA;
  if(!D){el.innerHTML='<div class="empty">Dados não carregados. Rode <code>python hub/tools/build_data.py</code>.</div>';return;}
  const casc=objify(D.cascata), det=objify(D.detalhe);
  const meses=casc.map(c=>c.mes), mesDef=meses[meses.length-1];
  const tipos=uniq(det,'tipo');

  el.innerHTML=`
    <div class="toolbar">
      <div class="field"><label>Mês</label>
        <select id="mes">${meses.map(m=>`<option value="${m}" ${m===mesDef?'selected':''}>${fmt.mesano(m)}</option>`).join('')}</select></div>
      <div class="field"><label>Tipo de movimento</label><select id="tipo">${selOpts(tipos,'','Todos')}</select></div>
    </div>
    <div class="grid g-4" id="kpis" style="margin-bottom:16px"></div>
    <div class="card"><div class="card-title"><h2>Cascata do mês</h2><span class="muted" id="cascSub"></span></div>
      <div id="casc" class="chart tall"></div></div>
    <div class="card" style="margin-top:16px"><div class="card-title"><h2>Saldo do plantel</h2><span class="muted">valor de plantel ao fim de cada mês</span></div>
      <div id="serie" class="chart"></div></div>
    <div class="card" style="margin-top:16px"><div class="card-title"><h2>Movimentos do mês</h2><span class="muted" id="detSub"></span></div>
      <div id="tbl"></div></div>`;

  const draw=()=>{
    clearCharts();
    const m=document.getElementById('mes').value, tipo=document.getElementById('tipo').value;
    const row=casc.find(c=>c.mes===m)||{};
    const i=meses.indexOf(m);

    const dPct=row.saldo_ini?((row.saldo_fim/row.saldo_ini)-1)*100:null;
    document.getElementById('kpis').innerHTML=[
      kpi('Saldo inicial',fmt.mi(row.saldo_ini),fmt.rs0(row.saldo_ini),''),
      kpi('Saldo final',fmt.mi(row.saldo_fim),fmt.rs0(row.saldo_fim),''),
      kpi('Variação',fmt.mi(row.d_total),fmt.pct(dPct),cls(row.d_total)),
      kpi('Conferência',Math.abs(row.check_dif||0)<0.01?'OK':fmt.rs(row.check_dif),
          Math.abs(row.check_dif||0)<0.01?'saldo fecha com a cascata':'cascata não fecha com o saldo',
          Math.abs(row.check_dif||0)<0.01?'pos':'neg'),
    ].join('');

    /* cascata */
    const steps=MOV.map(([k,l])=>[l,row[k]||0]).filter(s=>s[1]!==0);
    const cats=['Saldo inicial',...steps.map(s=>s[0]),'Saldo final'];
    const apoio=[0], sobe=['-'], desce=['-'], total=[row.saldo_ini||0];
    let run=row.saldo_ini||0;
    for(const [,v] of steps){
      total.push('-');
      if(v>=0){apoio.push(run);sobe.push(v);desce.push('-');}
      else{apoio.push(run+v);sobe.push('-');desce.push(-v);}
      run+=v;
    }
    apoio.push(0); sobe.push('-'); desce.push('-'); total.push(row.saldo_fim||0);
    document.getElementById('cascSub').textContent=`${fmt.mesano(m)} · ${steps.length} componentes`;
    const barLabel={show:true,position:'top',color:C.ink3,fontSize:11,
      formatter:p=>p.value==='-'||p.value==null?'':fmt.mi(p.value)};
    mkChart(document.getElementById('casc'),Object.assign(baseOpt(),{
      legend:{show:false},grid:{left:76,right:24,top:24,bottom:76},
      tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'#072B49',borderColor:C.line,
        textStyle:{color:C.ink},
        formatter:ps=>{const p=ps.find(x=>x.value!=='-'&&x.seriesName!=='apoio');
          if(!p)return '';
          const v=p.seriesName==='Redução'?-p.value:p.value;
          return `${p.name}<br><b>${fmt.rs(v)}</b>`;}},
      xAxis:axis({type:'category',data:cats,axisLabel:{color:C.ink3,rotate:28,fontSize:11}}),
      yAxis:axis({type:'value',scale:true,axisLabel:{color:C.ink3,formatter:v=>(v/1e6).toFixed(1)+' Mi'}}),
      series:[
        {name:'apoio',type:'bar',stack:'wf',silent:true,itemStyle:{color:'transparent'},data:apoio,barMaxWidth:46},
        {name:'Saldo',type:'bar',stack:'wf',itemStyle:{color:C.teal},data:total,barMaxWidth:46,label:barLabel},
        {name:'Aumento',type:'bar',stack:'wf',itemStyle:{color:C.pos},data:sobe,barMaxWidth:46,label:barLabel},
        {name:'Redução',type:'bar',stack:'wf',itemStyle:{color:C.neg},data:desce,barMaxWidth:46,
         label:Object.assign({},barLabel,{position:'bottom',formatter:p=>p.value==='-'||p.value==null?'':fmt.mi(-p.value)})},
      ]
    }));

    /* série de saldo */
    const s=mkChart(document.getElementById('serie'),Object.assign(baseOpt(),{
      legend:{show:false},grid:{left:76,right:24,top:20,bottom:56},
      dataZoom:zoom(Math.max(0,(1-24/meses.length)*100),100),
      tooltip:Object.assign(baseOpt().tooltip,{valueFormatter:v=>fmt.rs(v)}),
      xAxis:axis({type:'category',data:meses,boundaryGap:false,axisLabel:{color:C.ink3,formatter:fmt.mesano}}),
      yAxis:axis({type:'value',scale:true,axisLabel:{color:C.ink3,formatter:v=>(v/1e6).toFixed(0)+' Mi'}}),
      series:[{name:'Saldo final',type:'line',smooth:true,symbol:'none',data:casc.map(c=>c.saldo_fim),
        lineStyle:{color:C.amber,width:2.4},
        areaStyle:{color:new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:'rgba(202,151,3,.26)'},{offset:1,color:'rgba(202,151,3,0)'}])},
        markLine:{silent:true,symbol:'none',data:[{xAxis:m}],lineStyle:{color:C.teal,type:'dashed'},label:{show:false}}}]
    }));
    if(s) s.on('click',p=>{if(p.name){document.getElementById('mes').value=p.name;draw();}});

    const linhas=det.filter(d=>d.mes===m&&(!tipo||d.tipo===tipo));
    document.getElementById('detSub').textContent=`${linhas.length} movimento(s) · ${fmt.mesano(m)}`;
    if(!linhas.length){
      document.getElementById('tbl').innerHTML='<div class="empty">Nenhum movimento nesse recorte.</div>';
    }else{
      sortTable(document.getElementById('tbl'),[
        ['Animal','nome',v=>v||'—'],
        ['Categoria','categoria',v=>v||'—'],
        ['Tipo','tipo'],
        ['Status anterior','status_ant',v=>v||'—'],
        ['Status atual','status_atual',v=>v||'—'],
        ['Valor anterior','valor_ant',fmt.rs0],
        ['Valor atual','valor_atual',fmt.rs0],
        ['Efeito','efeito',(v)=>`<span class="${cls(v)}">${fmt.rs0(v)}</span>`],
      ],linhas,{sort:'efeito',dir:-1,maxH:520});
    }
    void i;
  };
  document.getElementById('mes').onchange=draw;
  document.getElementById('tipo').onchange=draw;
  draw();
}

/* ---- sidebar recolhível ---- */
(function collapse(){
  const btn=document.getElementById('collapseBtn'), app=document.querySelector('.app');
  if(localStorage.getItem('hpg-collapsed')==='1')app.classList.add('collapsed');
  btn.onclick=()=>{app.classList.toggle('collapsed');
    localStorage.setItem('hpg-collapsed',app.classList.contains('collapsed')?'1':'0');
    charts.forEach(c=>c.resize());};
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
