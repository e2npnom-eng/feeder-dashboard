/* ตรรกะแดชบอร์ด — ปกติไม่ต้องแก้ (ค่าตั้งต้นอยู่ที่ config.js, สีอยู่ที่ style.css) */

const A={t:['dateoutage','date','datetime','วันที่','วันเวลา'],f:['equiptment','equipment','feeder','อุปกรณ์','ฟีดเดอร์'],k:['tripreclose','type','ประเภท'],r:['reason','สาเหตุ','remark']};
const KEY={};for(const k in A)A[k].forEach(a=>KEY[a]=k);
const pad=n=>String(n).padStart(2,'0');
const MON=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
function when(v){let y,mo,d,hh=0,mi=0;
  if(v instanceof Date){v=new Date(Math.round(v/6e4)*6e4);y=v.getFullYear();mo=v.getMonth()+1;d=v.getDate();hh=v.getHours();mi=v.getMinutes()}
  else{const m=/^\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:\D+(\d{1,2}):(\d{2}))?/.exec(String(v??''));if(!m)return null;d=+m[1];mo=+m[2];y=+m[3];hh=+(m[4]||0);mi=+(m[5]||0)}
  if(y>2400)y-=543;if(mo<1||mo>12||d<1||d>31||hh>23||mi>59)return null;
  return`${y}-${pad(mo)}-${pad(d)}T${pad(hh)}:${pad(mi)}`}
const kind=(k,r)=>{for(const x of[String(k??''),String(r??'').slice(0,4)]){if(/t\s*\/\s*l|lock/i.test(x))return'T/L';if(/t\s*\/\s*r|reclos/i.test(x))return'T/R'}return null};
function parseWB(wb,fname){const fb=(/Event_Data_(.+?)\.xlsx?$/i.exec(fname)||[])[1]||'';const rows=[];let skipped=0;
  for(const n of wb.SheetNames){const[hd,...rs]=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null});if(!hd)continue;
    const cols=hd.map(x=>KEY[String(x??'').trim().toLowerCase()]);if(!cols.includes('t'))continue;
    for(const r of rs){const o={};cols.forEach((k,i)=>{if(k&&r[i]!=null)o[k]=r[i]});if(!Object.keys(o).length)continue;
      const t=when(o.t),k=kind(o.k,o.r),f=String(o.f||fb).trim();
      if(!(t&&k&&f)){skipped++;continue}
      rows.push([t,f,k,String(o.r??'').replace(/\s+/g,' ').trim()])}}
  return{rows,skipped}}
const merge=(a,b)=>{const m=new Map();for(const r of a.concat(b))m.set(r[1]+'|'+r[0]+'|'+r[2],r);return[...m.values()].sort((x,y)=>x[0]<y[0]?1:x[0]>y[0]?-1:x[1]<y[1]?1:-1)};
const iso=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
const ago=n=>{const d=new Date();d.setDate(d.getDate()-n);return iso(d)};
const be=s=>`${s.slice(8,10)}/${s.slice(5,7)}/${+s.slice(0,4)+543} ${s.slice(11,16)}`;
const beD=s=>`${+s.slice(8,10)} ${MON[+s.slice(5,7)-1]} ${+s.slice(0,4)+543}`;
const dayN=s=>Date.UTC(+s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10))/864e5;
const dstr=n=>new Date(n*864e5).toISOString().slice(0,10);
const cause=r=>{for(const[n,rx]of CAUSE)if(rx.test(r))return n;return'อื่นๆ'};
const cnt=(rows,k)=>rows.reduce((a,r)=>a+(r[2]===k),0);
const stn=f=>((/^[A-Za-z]{2,}/.exec(f)||[f])[0]).toUpperCase(); // สถานี = ตัวอักษรหน้ารหัสฟีดเดอร์ เช่น NNA01R-03 → NNA
function nice(max){for(const p of[1,10,100,1000,1e4])for(const f of[1,2,5]){const st=f*p;if(max/st<=5){const n=Math.ceil(max/st);return{max:n*st,ticks:Array.from({length:n+1},(_,i)=>i*st)}}}}
const rt=(x,y,w,h,r)=>{const q=r?Math.min(4,h,w/2):0;return`M${x},${y+h}V${y+q}Q${x},${y} ${x+q},${y}H${x+w-q}Q${x+w},${y} ${x+w},${y+q}V${y+h}Z`};
const rr=(x,y,w,h,r)=>{const q=r?Math.min(4,w,h/2):0;return`M${x},${y}H${x+w-q}Q${x+w},${y} ${x+w},${y+q}V${y+h-q}Q${x+w},${y+h} ${x+w-q},${y+h}H${x}Z`};

const $=id=>document.getElementById(id);
const NS='http://www.w3.org/2000/svg';
const h=(t,a,...c)=>{const e=document.createElement(t);for(const k in a||{})k==='class'?e.className=a[k]:k.startsWith('on')?e[k]=a[k]:e.setAttribute(k,a[k]);e.append(...c.flat().filter(x=>x!=null));return e};
const s=(t,a,...c)=>{const e=document.createElementNS(NS,t);for(const k in a||{})e.setAttribute(k,a[k]);e.append(...c.flat().filter(x=>x!=null));return e};
const num=n=>n.toLocaleString('en-US');
const dFrom=n=>new Date(n*864e5).toISOString().slice(0,10);
const rowsRL=(r,l)=>[['b1','T/R',r],['b2','T/L',l]];
const typ=k=>h('span',{class:'ty'},h('i',{class:'sw '+(k==='T/R'?'b1':'b2')}),k);
const empty=()=>h('div',{class:'empty'},'ไม่มีเหตุการณ์ในช่วงที่เลือก');
const ST=STATUS.map((n,i)=>[n,['✓','▲','✖'][i],['--good','--warn','--crit'][i]]);
const lvl=(r,l)=>l>=CFG.alarm.tl||r>=CFG.alarm.tr?2:l>=CFG.warn.tl||r>=CFG.warn.tr?1:0;

const START=()=>new Date().getFullYear()+'-01-01'; // วันเริ่มต้นเริ่มต้น = 1 ม.ค. ของปีปัจจุบัน (เปลี่ยนเองทุกปี)
let D=window.DATA||{updated:'',rows:[]},shown=50,msgT,D0='',D1='',DS='';
const X={cause:'',hour:-1},hrOf=r=>Math.floor(+r[0].slice(11,13)/3);
const clk=(e,f)=>{e.classList.add('ck');e.tabIndex=0;e.onclick=f;e.onkeydown=v=>{if(v.key==='Enter'||v.key===' '){v.preventDefault();f()}};return e};
const go=()=>{shown=50;renderAll()};
const setType=t=>{$('ty').value=$('ty').value===t?'':t;go()};
const setStation=n=>{const on=$('sn').value===n;$('sn').value=on?'':n;$('fd').value='';fillFilters();go()};
const setCause=c=>{X.cause=X.cause===c?'':c;go()};
const setHour=i=>{X.hour=X.hour===i?-1:i;go()};
try{const c=JSON.parse(localStorage.getItem('ev-data')||'null');if(c&&c.rows&&c.updated>(D.updated||''))D=c}catch(e){}
const feeders=sn=>[...new Set(D.rows.filter(r=>!sn||stn(r[1])===sn).map(r=>r[1]))].sort();
const stations=()=>[...new Set(D.rows.map(r=>stn(r[1])))].sort();
const msg=t=>{$('msg').textContent=t;clearTimeout(msgT);msgT=setTimeout(()=>$('msg').textContent='',9000)};

function fillFilters(){const sc=$('sn').value,ss=stations();
  $('sn').replaceChildren(h('option',{value:''},'ทุกสถานี'),...ss.map(x=>h('option',{value:x},x)));$('sn').value=ss.includes(sc)?sc:'';
  const cur=$('fd').value,fs=feeders($('sn').value);
  $('fd').replaceChildren(h('option',{value:''},'ทุกฟีดเดอร์'),...fs.map(f=>h('option',{value:f},f)));$('fd').value=fs.includes(cur)?cur:''}
function setRange(){let a='',b='';for(const r of D.rows){const d=r[0].slice(0,10);if(!a||d<a)a=d;if(!b||d>b)b=d}
  [$('d1'),$('d2')].forEach(e=>{e.min=a;e.max=b});{const y=START()<=b?START():b.slice(0,4)+'-01-01';DS=y>a?y:a}$('d1').value=DS;$('d2').value=b;D0=a;D1=b}
function pickFeeder(f,t){const on=$('fd').value===f;if(!on)$('sn').value=stn(f);if(t)$('ty').value=on?'':t;fillFilters();$('fd').value=on?'':f;shown=50;renderAll()}

function slice(){
  const sn=$('sn').value,fd=$('fd').value,ty=$('ty').value,a=$('d1').value||'0000-01-01',b=$('d2').value||'9999-12-31';
  const scope=D.rows.filter(r=>(!sn||stn(r[1])===sn)&&(!fd||r[1]===fd)),base=scope.filter(r=>!ty||r[2]===ty);
  const inR=(r,x,y)=>{const d=r[0].slice(0,10);return d>=x&&d<=y};
  const ok=D0&&a>=D0&&b>=a,span=ok?dayN(b)-dayN(a)+1:1,pb=ok?dFrom(dayN(a)-1):'',pa=ok?dFrom(dayN(a)-span):'';
  const okC=r=>!X.cause||cause(r[3])===X.cause,okH=r=>X.hour<0||hrOf(r)===X.hour,rng=base.filter(r=>inR(r,a,b));
  return{sn,fd,ty,a,b,span,scope,base,cur:rng.filter(r=>okC(r)&&okH(r)),forCause:rng.filter(okH),forHour:rng.filter(okC),prev:ok?base.filter(r=>inR(r,pa,pb)&&okC(r)&&okH(r)):[],hasPrev:!!ok&&D0<=pa}}

function renderHead(){let a='',b='';D.rows.forEach(r=>{const d=r[0];if(!a||d<a)a=d;if(!b||d>b)b=d});
  $('sub').replaceChildren('การไฟฟ้าส่วนภูมิภาคจังหวัดนครพนม — ',h('b',{},a?`ข้อมูล ${beD(a)} – ${beD(b)}`:'ยังไม่มีข้อมูล'),' | ',h('b',{},`อัพเดทข้อมูล ${D.updated?be(D.updated):'—'}`),' | จัดทำโดย นายศักดิ์อนันต์ แสนเมือง')}

function renderKPI(S){
  const R=cnt(S.cur,'T/R'),L=cnt(S.cur,'T/L'),T=R+L,pR=cnt(S.prev,'T/R'),pL=cnt(S.prev,'T/L'),pc=v=>T?Math.round(v/T*100)+'% ของทั้งหมด':'—';
  const dl=(c,p)=>{if(!S.hasPrev)return null;const d=c-p;return h('span',{},h('b',{class:d>0?'up':d<0?'dn':''},d>0?'▲':d<0?'▼':'–'),` ${d>0?'+':''}${d} เทียบ ${S.span} วันก่อนหน้า (${p})`)};
  let lt='';for(const r of S.scope)if(r[2]==='T/L'&&r[0]>lt)lt=r[0];
  const ago_=lt?dayN(iso(new Date()))-dayN(lt):null,mo=Math.max(1,S.span/30.4375);
  const tile=(lab,val,l1,l2,c,t)=>{const e=h('div',{class:'k'+(t&&S.ty===t?' act':''),style:`--c:${c}`},h('div',{class:'kl'},lab),h('div',{class:'kv'},String(val)),h('div',{class:'ks'},l1),h('div',{class:'ks'},l2||'\u00a0'));return t==null?e:clk(e,()=>setType(t))};
  $('kpi').replaceChildren(
    tile('เหตุการณ์ทั้งหมด',num(T),'รายการตามตัวกรอง',dl(T,pR+pL),'var(--a1)',''),
    tile(LABEL.tr,num(R),pc(R),dl(R,pR),'var(--s1)','T/R'),
    tile(LABEL.tl,num(L),pc(L),dl(L,pL),'var(--s2)','T/L'),
    tile('เฉลี่ยต่อเดือน',T?(T/mo).toFixed(1):'—','ครั้ง/เดือน',lt?`T/L ล่าสุด ${ago_<=0?'วันนี้':ago_+' วันที่แล้ว'} (${beD(lt)})`:'ไม่มี T/L ในข้อมูล','var(--a2)'))}

const tip=$('tip');
function showTip(x,y,title,rows){tip.replaceChildren(h('div',{class:'tt'},title),...rows.map(([c,n,v])=>h('div',{class:'tr'},h('i',{class:'lk '+c}),h('b',{},String(v)),h('span',{},n))));
  tip.style.display='block';tip.style.left=Math.max(8,Math.min(x+12,innerWidth-tip.offsetWidth-8))+'px';tip.style.top=Math.max(8,y-tip.offsetHeight-12)+'px'}
const hideTip=()=>tip.style.display='none';
function hov(el,title,rows,g=el){const on=e=>{g.classList.add('on');const b=e.type==='focus'?el.getBoundingClientRect():{left:e.clientX,top:e.clientY};showTip(b.left,b.top,title,rows)};
  ['pointerenter','pointermove','focus'].forEach(t=>el.addEventListener(t,on));['pointerleave','blur'].forEach(t=>el.addEventListener(t,()=>{g.classList.remove('on');hideTip()}))}
const tbl=(hd,rows)=>h('table',{class:'mini'},h('thead',{},h('tr',{},...hd.map(x=>h('th',{},x)))),h('tbody',{},...rows.map(r=>h('tr',{},...r.map(c=>h('td',{},String(c)))))));
const tdet=(id,hd,rows)=>$(id).replaceChildren(tbl(hd,rows));

function hbars(el,items){
  const mx=Math.max(1,...items.map(i=>i.r+i.l));
  el.replaceChildren(...(items.length?items.map(it=>{const t=it.r+it.l;
    const row=h('div',{class:'hr',tabindex:'0','aria-label':`${it.name}: T/R ${it.r}, T/L ${it.l}`},h('span',{class:'hl',title:it.name},it.name),
      h('div',{},h('div',{class:'bar',style:`width:${Math.max(t/mx*100,1.5)}%`},it.r?h('i',{class:'b1',style:`flex:${it.r}`}):null,it.l?h('i',{class:'b2',style:`flex:${it.l}`}):null)),h('span',{class:'hv'},num(t)));
    hov(row,it.name,it.only?[['b2','T/L',it.l]]:rowsRL(it.r,it.l));
    if(it.pick)clk(row,it.pick);if(it.act)row.classList.add('act');
    return row}):[empty()]))}

function colChart(el,o){
  el.replaceChildren();
  const n=o.labels.length,W=Math.max(el.clientWidth||320,260),H=o.h||230,m={t:o.lab?22:12,r:8,b:26,l:34},iw=W-m.l-m.r,ih=H-m.t-m.b,yb=m.t+ih;
  const nc=nice(Math.max(1,...o.r.map((v,i)=>v+o.l[i]))),step=iw/n,bw=Math.min(24,Math.max(3,step-3)),Y=v=>v/nc.max*ih,k=Math.ceil(n*(o.lw||46)/iw);
  const svg=s('svg',{viewBox:`0 0 ${W} ${H}`,width:W,height:H,role:'img','aria-label':o.aria});
  nc.ticks.forEach(t=>svg.append(s('line',{x1:m.l,x2:W-m.r,y1:yb-Y(t),y2:yb-Y(t),class:t?'gl':'ax'}),s('text',{x:m.l-6,y:yb-Y(t)+4,class:'tk','text-anchor':'end'},t>=1000?t/1000+'K':String(t))));
  for(let i=0;i<n;i++){
    const x0=m.l+step*i,cx=x0+step/2,x=cx-bw/2,hr=o.r[i]?Math.max(2,Y(o.r[i])):0,hl=o.l[i]?Math.max(2,Y(o.l[i])):0,g=s('g',{class:'col'});
    if(hr)g.append(s('path',{d:rt(x,yb-hr,bw,hr,!hl),class:'f1'}));
    if(hl)g.append(s('path',{d:rt(x,yb-hr-(hr?2:0)-hl,bw,hl,true),class:'f2'}));
    svg.append(g);
    if(o.lab&&step>=24&&hr+hl)svg.append(s('text',{x:cx,y:yb-hr-hl-(hr&&hl?2:0)-5,class:'val','text-anchor':'middle'},String(o.r[i]+o.l[i])));
    if(i%k===0)svg.append(s('text',{x:cx,y:H-8,class:'tk','text-anchor':'middle'},o.short[i]));
    const hit=s('rect',{x:x0,y:m.t,width:step,height:ih,class:'hit',tabindex:'0','aria-label':`${o.labels[i]}: T/R ${o.r[i]}, T/L ${o.l[i]}`});
    if(o.pick)clk(hit,()=>o.pick(i));hov(hit,o.labels[i],rowsRL(o.r[i],o.l[i]),g);svg.append(hit)}
  el.append(svg)}

function donut(el,r,l){
  const T=r+l,R=62,C=2*Math.PI*R,gap=r&&l?3:0,svg=s('svg',{viewBox:'0 0 180 180',width:180,height:180,role:'img','aria-label':`สัดส่วน T/R ${r} ครั้ง และ T/L ${l} ครั้ง`});
  svg.append(s('circle',{cx:90,cy:90,r:R,fill:'none',stroke:'var(--grid)','stroke-width':24}));
  let off=0;
  [[r,'st1','T/R','b1'],[l,'st2','T/L','b2']].forEach(([v,cls,name,b])=>{if(!v)return;const len=v/T*C;
    const c=s('circle',{cx:90,cy:90,r:R,fill:'none','stroke-width':24,class:cls,'stroke-dasharray':`${Math.max(1,len-gap)} ${C}`,'stroke-dashoffset':-off,transform:'rotate(-90 90 90)',tabindex:'0','aria-label':`${name} ${v} ครั้ง`});
    clk(c,()=>setType(name));hov(c,name,[[b,name,v]]);svg.append(c);off+=len});
  svg.append(s('text',{x:90,y:92,class:'ctr','text-anchor':'middle'},num(T)),s('text',{x:90,y:110,class:'cts','text-anchor':'middle'},'ครั้ง'));
  el.replaceChildren(svg,h('div',{class:'lg'},...[['b1','T/R',r],['b2','T/L',l]].map(([c,n,v])=>clk(h('span',{},h('i',{class:'sw '+c}),n+': ',h('b',{},num(v)),` (${T?Math.round(v/T*100):0}%)`),()=>setType(n)))))}

function renderWatch(S){
  const c30=ago(CFG.days),fs=S.fd?[S.fd]:feeders(S.sn);$('wsub').textContent=`${CFG.days} วันล่าสุด`;
  const L=fs.map(f=>{const rec=D.rows.filter(r=>r[1]===f&&r[0]>=c30),r=cnt(rec,'T/R'),l=cnt(rec,'T/L');return{f,r,l,lv:lvl(r,l)}}).sort((a,b)=>b.lv-a.lv||b.l-a.l||b.r-a.r||(a.f<b.f?-1:1));
  const c=[0,1,2].map(i=>L.filter(x=>x.lv===i).length);
  $('c-watch').replaceChildren(h('div',{class:'ws',style:'margin-bottom:4px'},`วิกฤต ${c[2]} · เฝ้าระวัง ${c[1]} · ปกติ ${c[0]} (จาก ${L.length} ฟีดเดอร์)`),
    ...L.slice(0,10).map(x=>{const[lab,ic,col]=ST[x.lv];return h('button',{class:'wr',onclick:()=>pickFeeder(x.f)},h('div',{},h('div',{class:'wn'},x.f),h('div',{class:'ws'},`T/R ${x.r} · T/L ${x.l}`)),h('span',{class:'st'},h('b',{style:`color:var(${col})`},ic),lab))}),
    L.length>10?h('div',{class:'ws',style:'margin-top:4px'},`แสดง 10 อันดับแรกจาก ${L.length} ฟีดเดอร์`):'',
    h('div',{class:'ws',style:'margin-top:8px'},`เกณฑ์: เฝ้าระวัง T/L ≥ ${CFG.warn.tl} หรือ T/R ≥ ${CFG.warn.tr} · วิกฤต T/L ≥ ${CFG.alarm.tl} หรือ T/R ≥ ${CFG.alarm.tr}`))}

function renderCharts(S){
  const T=S.cur.length;
  if(!T){['c-cause','c-st','c-trend','c-top','c-topl','c-hour'].forEach(i=>$(i).replaceChildren(empty()));$('c-donut').replaceChildren();['t-cause','t-trend','t-hour'].forEach(i=>$(i).replaceChildren());$('tsub').textContent='';return}
  const grp=(keyf,namef,src=S.cur)=>{const m=new Map();src.forEach(e=>{const k=keyf(e),o=m.get(k)||{name:namef?namef(k):k,r:0,l:0};o[e[2]==='T/R'?'r':'l']++;m.set(k,o)});return[...m.entries()].map(([k,v])=>({...v,k}))};
  const tot=a=>a.r+a.l;
  // สาเหตุ
  const causes=grp(e=>cause(e[3]),null,S.forCause).sort((a,b)=>tot(b)-tot(a));hbars($('c-cause'),causes.map(i=>({...i,pick:()=>setCause(i.k),act:X.cause===i.k})));
  tdet('t-cause',['หมวด','T/R','T/L','รวม'],causes.map(i=>[i.name,i.r,i.l,tot(i)]));
  // donut + สถานี
  donut($('c-donut'),cnt(S.cur,'T/R'),cnt(S.cur,'T/L'));
  hbars($('c-st'),grp(e=>stn(e[1])).sort((a,b)=>tot(b)-tot(a)).slice(0,8).map(i=>({...i,name:'สถานี '+i.name,pick:()=>setStation(i.k),act:S.sn===i.k})));
  // แนวโน้ม
  const daily=S.span<=45,labels=[],short=[],r=[],l=[],idx={},keys=[];
  if(daily){for(let n=dayN(S.a);n<=dayN(S.b);n++){const key=dFrom(n);idx[key]=labels.length;keys.push(key);labels.push(beD(key));short.push(`${+key.slice(8)} ${MON[+key.slice(5,7)-1]}`);r.push(0);l.push(0)}}
  else{let y=+S.a.slice(0,4),mo=+S.a.slice(5,7)-1;const ey=+S.b.slice(0,4),em=+S.b.slice(5,7)-1;
    while(y<ey||(y===ey&&mo<=em)){idx[`${y}-${pad(mo+1)}`]=labels.length;keys.push(`${y}-${pad(mo+1)}`);labels.push(`${MON[mo]} ${y+543}`);short.push(`${MON[mo]} ${String(y+543).slice(2)}`);r.push(0);l.push(0);if(++mo>11){mo=0;y++}}}
  S.cur.forEach(e=>{const i=idx[e[0].slice(0,daily?10:7)];if(i!=null)(e[2]==='T/R'?r:l)[i]++});
  $('tsub').textContent=`ต่อ${daily?'วัน':'เดือน'}`;
  const pickT=i=>{const k=keys[i],a0=daily?k:k+'-01',b0=daily?k:dFrom(Date.UTC(+k.slice(0,4),+k.slice(5,7),0)/864e5),A=a0<D0?D0:a0,B=b0>D1?D1:b0;
    if($('d1').value===A&&$('d2').value===B)setRange();else{$('d1').value=A;$('d2').value=B}go()};
  colChart($('c-trend'),{labels,short,r,l,lab:true,h:300,pick:pickT,aria:`กราฟแท่งจำนวนเหตุการณ์ต่อ${daily?'วัน':'เดือน'} รวม ${T} ครั้ง`});
  tdet('t-trend',[daily?'วัน':'เดือน','T/R','T/L','รวม'],labels.map((x,i)=>[x,r[i],l[i],r[i]+l[i]]));
  // TOP 10
  const fm=grp(e=>e[1]);
  hbars($('c-top'),fm.slice().sort((a,b)=>tot(b)-tot(a)||(a.name<b.name?-1:1)).slice(0,10).map(i=>({...i,pick:()=>pickFeeder(i.k),act:S.fd===i.k})));
  hbars($('c-topl'),fm.filter(i=>i.l).sort((a,b)=>b.l-a.l||(a.name<b.name?-1:1)).slice(0,10).map(i=>({name:i.name,r:0,l:i.l,only:true,pick:()=>pickFeeder(i.k,'T/L'),act:S.fd===i.k})));
  // ช่วงเวลา
  const hb=grp(hrOf,null,S.forHour),hm=new Map(hb.map(x=>[x.k,x]));
  const bands=Array.from({length:8},(_,i)=>({name:`${pad(i*3)}:00–${pad(i*3+2)}:59`,r:(hm.get(i)||{}).r||0,l:(hm.get(i)||{}).l||0,pick:()=>setHour(i),act:X.hour===i}));
  hbars($('c-hour'),bands);tdet('t-hour',['ช่วงเวลา','T/R','T/L','รวม'],bands.map(i=>[i.name,i.r,i.l,tot(i)]))}

function renderLog(S){
  const q=$('q').value.trim().toLowerCase(),rows=S.cur.filter(r=>!q||(r[1]+' '+r[3]+' '+cause(r[3])).toLowerCase().includes(q));
  $('lc').textContent=`${num(rows.length)} รายการ`;
  $('lb').replaceChildren(...(rows.length?rows.slice(0,shown).map(r=>h('tr',{},h('td',{class:'nw'},be(r[0])),clk(h('td',{class:'nw'},r[1]),()=>pickFeeder(r[1])),clk(h('td',{class:'nw'},typ(r[2])),()=>setType(r[2])),clk(h('td',{class:'nw'},cause(r[3])),()=>setCause(cause(r[3]))),h('td',{class:'rs'},r[3]))):[h('tr',{},h('td',{class:'empty',colspan:'5'},'ไม่พบข้อมูล'))]));
  $('more').style.display=rows.length>shown?'':'none';$('more').textContent=`แสดงเพิ่ม (${shown}/${num(rows.length)})`}

function renderChips(){const c=[];
  if(X.cause)c.push(['สาเหตุ: '+X.cause,()=>{X.cause='';go()}]);
  if(X.hour>=0)c.push([`ช่วงเวลา: ${pad(X.hour*3)}:00–${pad(X.hour*3+2)}:59`,()=>{X.hour=-1;go()}]);
  if(D0&&($('d1').value!==DS||$('d2').value!==D1))c.push([`ช่วงวันที่: ${beD($('d1').value)} – ${beD($('d2').value)}`,()=>{setRange();go()}]);
  $('chips').replaceChildren(...c.map(([t,f])=>h('button',{class:'chip',title:'คลิกเพื่อล้าง',onclick:f},t+' ✕')))}
const renderAll=()=>{const S=slice();renderHead();renderChips();renderKPI(S);renderWatch(S);renderCharts(S);renderLog(S)};

async function onFiles(files){
  if(typeof XLSX==='undefined'){msg('โหลดไลบรารีอ่าน Excel ไม่ได้ (ต้องเชื่อมต่ออินเทอร์เน็ต)');return}
  let add=[],skip=0,bad=[];
  for(const f of files){try{const p=parseWB(XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true}),f.name);add=add.concat(p.rows);skip+=p.skipped}catch(e){bad.push(f.name)}}
  const before=D.rows.length;D={updated:iso(new Date()),rows:merge(D.rows,add)};
  try{localStorage.setItem('ev-data',JSON.stringify(D))}catch(e){}
  fillFilters();setRange();shown=50;renderAll();
  msg(`อ่าน ${files.length-bad.length} ไฟล์ · ${num(add.length)} แถว · เพิ่มใหม่ ${num(D.rows.length-before)}`+(skip?` · ข้าม ${skip} แถวที่อ่านวันที่/ประเภทไม่ได้`:'')+(bad.length?` · อ่านไม่ได้: ${bad.join(', ')}`:'')+' — กด "⬇ data.js" แล้ว commit เพื่อให้ทุกคนเห็น')}
$('up').onclick=()=>$('f').click();
$('f').onchange=e=>{onFiles([...e.target.files]);e.target.value=''};
addEventListener('dragover',e=>{e.preventDefault();$('drop').style.display='flex'});
addEventListener('dragleave',e=>{if(!e.relatedTarget)$('drop').style.display='none'});
addEventListener('drop',e=>{e.preventDefault();$('drop').style.display='none';if(e.dataTransfer.files.length)onFiles([...e.dataTransfer.files])});
$('dl').onclick=()=>{const t=`window.DATA={"updated":${JSON.stringify(D.updated)},"rows":[\n${D.rows.map(r=>JSON.stringify(r)).join(',\n')}\n]};\n`,a=h('a',{href:URL.createObjectURL(new Blob([t],{type:'text/javascript'})),download:'data.js'});a.click()};
$('rs').onclick=()=>{try{localStorage.removeItem('ev-data')}catch(e){}location.reload()};

$('sn').onchange=()=>{$('fd').value='';fillFilters();shown=50;renderAll()};
['fd','ty','d1','d2'].forEach(i=>$(i).onchange=()=>{shown=50;renderAll()});
$('rf').onclick=()=>{X.cause='';X.hour=-1;$('sn').value='';$('ty').value='';fillFilters();$('fd').value='';$('q').value='';setRange();shown=50;renderAll()};
$('q').oninput=()=>{shown=50;renderLog(slice())};
$('more').onclick=()=>{shown+=50;renderLog(slice())};
let rz;addEventListener('resize',()=>{clearTimeout(rz);rz=setTimeout(()=>renderCharts(slice()),150)});
fillFilters();setRange();renderAll();
