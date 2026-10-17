(()=>{
const M=['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'],
 WD=['อา','จ','อ','พ','พฤ','ศ','ส'],p=n=>String(n).padStart(2,'0'),
 fmt=v=>v?`${v.slice(8)}/${v.slice(5,7)}/${+v.slice(0,4)+543}`:'วว/ดด/พพพพ',
 desc=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value'),
 pop=document.createElement('div');let cur=null;
pop.className='tdp';pop.hidden=true;document.body.append(pop);
const close=()=>{pop.hidden=true;cur=null};
function draw(){const{inp,y,m}=cur,v=inp.value,mn=inp.min,mx=inp.max,f=new Date(y,m,1).getDay(),n=new Date(y,m+1,0).getDate();
 let h=`<div class="tdh"><button data-n="-12">«</button><button data-n="-1">‹</button><b>${M[m]} ${y+543}</b><button data-n="1">›</button><button data-n="12">»</button></div><div class="tdg">`+WD.map(w=>`<i>${w}</i>`).join('')+'<span></span>'.repeat(f);
 for(let d=1;d<=n;d++){const s=`${y}-${p(m+1)}-${p(d)}`;h+=`<button data-d="${s}"${(mn&&s<mn)||(mx&&s>mx)?' disabled':''}${s===v?' class="on"':''}>${d}</button>`}
 pop.innerHTML=h+'</div>'}
pop.onclick=e=>{const b=e.target.closest('button');if(!b||!cur)return;
 if(b.dataset.n){const t=cur.m+ +b.dataset.n;cur.y+=Math.floor(t/12);cur.m=(t%12+12)%12;draw()}
 else if(b.dataset.d){cur.inp.value=b.dataset.d;cur.inp.dispatchEvent(new Event('change'));close()}};
document.addEventListener('click',e=>{if(!pop.contains(e.target))close()});
['d1','d2'].forEach(id=>{const inp=document.getElementById(id),btn=document.createElement('button'),ref=()=>btn.textContent=fmt(inp.value);
 btn.type='button';btn.className='tdb';inp.hidden=true;inp.after(btn);
 Object.defineProperty(inp,'value',{get(){return desc.get.call(this)},set(x){desc.set.call(this,x);ref()}});
 btn.onclick=e=>{e.stopPropagation();if(cur&&cur.inp===inp)return close();
  const v=inp.value||inp.max||new Date().toISOString().slice(0,10);
  cur={inp,y:+v.slice(0,4),m:+v.slice(5,7)-1};draw();pop.hidden=false;
  const r=btn.getBoundingClientRect();pop.style.left=Math.max(8,Math.min(r.left,innerWidth-pop.offsetWidth-8))+'px';pop.style.top=r.bottom+4+'px'};
 ref()});
})();
