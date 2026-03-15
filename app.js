// ── State ──────────────────────────────────────────────
let launches=[], financials={}, mosaic={}, activeFilter='all';

// ── Tab Navigation ─────────────────────────────────────
function showTab(id){
  document.querySelectorAll('section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  event.target.classList.add('active');
}

// ── Data Loader ────────────────────────────────────────
async function load(path){const r=await fetch(path);return r.json();}

async function init(){
  [launches, financials, mosaic] = await Promise.all([
    load('data/launches.json').then(d=>d.launches),
    load('data/financials.json'),
    load('data/mosaic.json')
  ]);
  renderStats();
  renderLaunchTable(launches);
  renderOverviewCharts();
  renderFinancial();
  renderMosaic();
}

// ── Stats Bar ──────────────────────────────────────────
function renderStats(){
  const suc=launches.filter(l=>l.outcome==='success').length;
  const rate=Math.round(suc/launches.length*100);
  document.getElementById('s-total').textContent=launches.length;
  document.getElementById('s-suc').textContent=suc;
  document.getElementById('s-rate').textContent=rate+'%';
  const last=financials.annual.at(-1);
  document.getElementById('s-rev').textContent='$'+last.revenue+'M';
  document.getElementById('s-backlog').textContent='$'+financials.key_metrics.backlog_2024+'M';
}

// ── Launch Table ───────────────────────────────────────
function renderLaunchTable(data){
  const tb=document.getElementById('launch-tbody');
  const badge=o=>o==='success'?'<span class="badge badge-s">성공</span>':
                  o==='failure'?'<span class="badge badge-f">실패</span>':
                  '<span class="badge badge-p">부분</span>';
  tb.innerHTML=data.map(l=>`
    <tr>
      <td style="color:var(--txt3)">${l.id}</td>
      <td style="font-weight:600">${l.name}</td>
      <td style="color:var(--txt2)">${l.date}</td>
      <td>${badge(l.outcome)}</td>
      <td style="color:var(--txt2)">${l.orbit||'—'}</td>
      <td style="color:var(--txt2);font-size:.76rem">${l.payload}</td>
      <td>${l.customer}</td>
      <td style="color:var(--txt3);font-size:.76rem">${l.notes}</td>
    </tr>`).join('');
  document.getElementById('launch-count').textContent=`${data.length}개 미션 표시`;
}

// ── Filter ─────────────────────────────────────────────
function setFilter(f,btn){
  activeFilter=f;
  document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  filterLaunches();
}

function filterLaunches(){
  const q=document.getElementById('search').value.toLowerCase();
  let d=launches;
  if(activeFilter!=='all') d=d.filter(l=>l.outcome===activeFilter||l.site===activeFilter);
  if(q) d=d.filter(l=>[l.name,l.customer,l.payload,l.notes].some(v=>v.toLowerCase().includes(q)));
  renderLaunchTable(d);
}

// ── Bar Chart Helper ───────────────────────────────────
function barChart(el, rows, color='var(--acc)'){
  const max=Math.max(...rows.map(r=>Math.abs(r.val)));
  el.innerHTML=rows.map(r=>{
    const pct=Math.round(Math.abs(r.val)/max*100);
    const c=r.color||color;
    return `<div class="bar-row">
      <div class="bar-label">${r.label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:${c}"></div></div>
      <div class="bar-val">${r.display||r.val}</div>
    </div>`;
  }).join('');
}

// ── Overview Charts ────────────────────────────────────
function renderOverviewCharts(){
  // launches by year
  const byYear={};
  launches.forEach(l=>{const y=l.date.slice(0,4);byYear[y]=(byYear[y]||0)+1;});
  barChart(document.getElementById('launch-by-year'),
    Object.entries(byYear).map(([y,c])=>({label:y,val:c,display:c+'회'})));

  // customer dist (top 6)
  const byC={};
  launches.forEach(l=>{
    l.customer.split('/').forEach(c=>{const k=c.trim();byC[k]=(byC[k]||0)+1;});
  });
  const top=Object.entries(byC).sort((a,b)=>b[1]-a[1]).slice(0,6);
  barChart(document.getElementById('customer-dist'),
    top.map(([k,v])=>({label:k.slice(0,10),val:v,display:v+'회'})),
    'var(--acc2)');

  // revenue sparkline as bar
  barChart(document.getElementById('rev-chart'),
    financials.annual.map(a=>({label:a.year,val:a.revenue,display:'$'+a.revenue+'M'})),
    'var(--suc)');

  // recent timeline
  const tl=document.getElementById('recent-timeline');
  tl.innerHTML=launches.slice(-10).reverse().map(l=>`
    <div class="tl-item">
      <div class="tl-dot ${l.outcome==='success'?'s':'f'}"></div>
      <div class="tl-date">${l.date}</div>
      <div class="tl-name">${l.name}</div>
      <div class="tl-meta">${l.customer} · ${l.orbit||'실패'}</div>
    </div>`).join('');
}

// ── Financial ──────────────────────────────────────────
function renderFinancial(){
  const last=financials.annual.at(-1);
  const prev=financials.annual.at(-2);
  const revGrowth=Math.round((last.revenue-prev.revenue)/prev.revenue*100);
  const km=financials.key_metrics;

  const kpis=[
    {label:'2024 매출',val:'$'+last.revenue+'M',sub:'YoY +'+revGrowth+'%',cls:'up'},
    {label:'2024 매출총이익',val:'$'+last.gross_profit+'M',sub:'GP마진 '+Math.round(last.gross_profit/last.revenue*100)+'%',cls:'up'},
    {label:'2024 순손실',val:'-$'+Math.abs(last.net_loss)+'M',sub:'전년 대비 개선',cls:'flat'},
    {label:'수주잔고',val:'$'+km.backlog_2024+'M',sub:'매출 '+Math.round(km.backlog_2024/last.revenue)+'배',cls:'up'},
    {label:'상장가',val:'$'+financials.ipo_price,sub:financials.ipo_date,cls:'flat'},
    {label:'임직원',val:km.employees_2024.toLocaleString(),sub:'2024 기준',cls:'flat'},
  ];

  document.getElementById('fin-kpis').innerHTML=kpis.map(k=>`
    <div class="fin-card">
      <h4>${k.label}</h4>
      <div class="val ${k.cls}">${k.val}</div>
      <div class="sub">${k.sub}</div>
    </div>`).join('');

  barChart(document.getElementById('fin-rev-chart'),
    financials.annual.map(a=>({label:a.year,val:a.revenue,display:'$'+a.revenue+'M'})),
    'var(--suc)');

  barChart(document.getElementById('fin-gp-chart'),
    financials.annual.map(a=>({label:a.year,val:Math.max(a.gross_profit,0),
      display:a.gross_profit<0?'-$'+Math.abs(a.gross_profit)+'M':'$'+a.gross_profit+'M',
      color:a.gross_profit>=0?'var(--suc)':'var(--fail)'})));

  barChart(document.getElementById('fin-nl-chart'),
    financials.annual.map(a=>({label:a.year,val:Math.abs(a.net_loss),
      display:'-$'+Math.abs(a.net_loss)+'M'})),
    'var(--fail)');
}

// ── Mosaic ─────────────────────────────────────────────
function renderMosaic(){
  document.getElementById('mosaic-score').textContent=mosaic.mosaic_score.toFixed(1);

  const trendCls={up:'trend-up',down:'trend-down',neutral:'trend-neutral'};
  document.getElementById('signals-grid').innerHTML=mosaic.signals.map(s=>`
    <div class="signal-card">
      <div><span class="signal-cat">${s.category}</span></div>
      <div class="signal-body">
        <h4 class="${trendCls[s.trend]}">${s.signal}</h4>
        <p>${s.implication}</p>
        <div class="weight-dots">${[1,2,3,4,5].map(i=>`<div class="dot${i<=s.weight?' on':''}"></div>`).join('')}</div>
      </div>
      <div style="font-size:.72rem;color:var(--txt3)">${s.source}</div>
    </div>`).join('');

  document.getElementById('verdict-box').innerHTML=
    `<strong style="color:var(--acc2)">종합 판단</strong><br>${mosaic.verdict}`;
}

// ── Boot ───────────────────────────────────────────────
init();
