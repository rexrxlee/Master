(function(root){
 const money=n=>Math.round(Number(n)*100)/100;
 const today=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Asia/Singapore'});
 function date(s){if(!/^\d{4}-\d{2}-\d{2}$/.test(s))throw Error('Enter a valid date');const d=new Date(s+'T12:00:00Z');if(d.toISOString().slice(0,10)!==s)throw Error('Invalid date');return d;}
 function year(s){const d=date(s),y=d.getUTCFullYear()+1,m=d.getUTCMonth(),day=d.getUTCDate();return new Date(Date.UTC(y,m,Math.min(day,new Date(Date.UTC(y,m+1,0)).getUTCDate()))).toISOString().slice(0,10);}
 function complete(p){return !!(String(p.number||'').trim()&&p.start&&p.due&&p.premium!==''&&p.premium!=null&&Number.isFinite(+p.premium)&&+p.premium>0);}
 function cash(p){return !p.funding||p.funding==='Cash';}
 function validate(p){for(const k of ['start','due','end'])if(p[k])date(p[k]);if(!String(p.description||'').trim())throw Error('Description is required.');if(p.funding&&!['Cash','CPF MediSave'].includes(p.funding))throw Error('Choose Cash or CPF MediSave.');if(p.premium!==''&&p.premium!=null&&(!Number.isFinite(+p.premium)||+p.premium<=0))throw Error('Premium must be positive or left blank.');if(!Number.isFinite(+p.increment)||+p.increment<0)throw Error('Increment must be zero or positive.');if(p.start&&(p.due&&p.due<p.start||p.end&&p.end<p.start))throw Error('Dates must follow the policy start.');}
 function amount(p,due){const a=date(p.start),b=date(due);let months=(b.getUTCFullYear()-a.getUTCFullYear())*12+b.getUTCMonth()-a.getUTCMonth();if(b.getUTCDate()<a.getUTCDate())months--;return money(+p.premium+Math.max(0,months)*(+p.increment||0));}
 function first(p,now=today()){validate(p);const q={...p};if(!complete(q))return q;while(q.due<now&&q.annual)q.due=year(q.due);if(q.due<now&&!q.annual)q.finished=true;return q;}
 function events(state,until){const out=[];for(const p of state.policies){if(p.finished||!cash(p)||!complete(p))continue;let due=p.due;for(let i=0;i<50&&due<=until;i++){if(!p.end||due<=p.end)out.push({id:p.id,due,amount:amount(p,due),description:p.description});if(!p.annual||p.end&&due>=p.end)break;due=year(due);}}return out.sort((a,b)=>a.due.localeCompare(b.due));}
 function batch(state){const e=events(state,'2099-12-31'),due=e[0]?.due||'';return {due,amount:money(e.filter(x=>x.due.slice(0,7)===due.slice(0,7)).reduce((s,x)=>s+x.amount,0))};}
 // Reserve every policy's next payment by the earliest deadline. This deliberately
 // funds later bills early rather than depending on one month's future salary.
 function plan(state){const next=events(state,'2099-12-31').filter((e,i,all)=>all.findIndex(x=>x.id===e.id)===i);return {due:next[0]?.due||'',amount:money(next.reduce((sum,e)=>sum+e.amount,0))};}
 // Subsequent renewals accrue across the 12 months after their preceding renewal.
 function reserve(state,month){const key=month.slice(0,7),index=s=>+s.slice(0,4)*12+(+s.slice(5,7)-1);let cents=0;for(const p of state.policies){if(!cash(p)||!complete(p)||p.finished||!p.annual)continue;let previous=p.due;for(let i=0;i<50;i++){const due=year(previous);if(p.end&&due>p.end)break;const offset=index(key)-index(previous);if(offset>=0&&offset<12){const total=Math.round(amount(p,due)*100);cents+=Math.floor(total/12)+(offset<total%12?1:0);break;}if(index(previous)>index(key))break;previous=due;}}return cents/100;}
 function parse(v){if(!v)return {version:1,revision:0,linkedName:'',policies:[],history:[]};const s=JSON.parse(v);if(s.version!==1||!Array.isArray(s.policies))throw Error('Insurance data cannot be read');return s;}
 function mergePolicies(state,incoming,now=today()){
   if(!Array.isArray(incoming)||!incoming.length||incoming.length>200)throw Error('Import 1–200 policies.');
   const next=JSON.parse(JSON.stringify(state)),seen=new Set();
   for(const raw of incoming){
     const p={};for(const k of ['owner','insurer','description','number','funding','start','end','due','notes'])p[k]=String(raw[k]??'').trim();
     p.funding=p.funding||'Cash';p.premium=raw.premium==null||raw.premium===''?'':Number(raw.premium);p.increment=Number(raw.increment||0);p.annual=raw.annual!==false;
     const key=p.number?p.insurer.toLowerCase()+'|'+p.number:p.owner.toLowerCase()+'|'+p.description.toLowerCase();
     if(seen.has(key))throw Error('Duplicate policy in import');seen.add(key);validate(p);
     const match=next.policies.find(x=>p.number?x.number===p.number&&(x.insurer||'').toLowerCase()===p.insurer.toLowerCase():!x.number&&(x.owner||'').toLowerCase()===p.owner.toLowerCase()&&x.description.toLowerCase()===p.description.toLowerCase());
     if(match)throw Error('Policy already exists: '+p.description+'. Edit it instead of importing again.');
     p.historicalPaid={};for(const y of ['2024','2025','2026']){const v=raw.historicalPaid?.[y];if(v!=null&&v!==''){if(!Number.isFinite(+v)||+v<0)throw Error('Invalid historical payment');p.historicalPaid[y]=+v;}}
     p.id=crypto.randomUUID();next.policies.push(first(p,now));
   }return next;
 }
 function registerRows(state){return [['Insurance register — edit policies in the app. This sheet updates on every save.'],['Owner','Insurer','Description','Policy number','Funding','Annual premium (SGD)','Cover start','Policy end','Next renewal','Monthly increase (SGD)','Paid 2024','Paid 2025','Paid 2026','Notes','Status'],...state.policies.map(p=>[p.owner||'',p.insurer||'',p.description,p.number||'',p.funding||'Cash',p.premium??'',p.start||'',p.end||'',p.due||'',p.increment||0,p.historicalPaid?.['2024']??'',p.historicalPaid?.['2025']??'',p.historicalPaid?.['2026']??'',p.notes||'',!complete(p)?'Missing details':!cash(p)?'MediSave — excluded from cash goals':p.finished?'Complete':'Cash savings'])];}
 root.FinInsurance={money,today,year,validate,amount,first,events,batch,plan,reserve,parse,complete,cash,mergePolicies,registerRows};
})(typeof window==='undefined'?globalThis:window);
