(function(root){
 const money=n=>Math.round(Number(n)*100)/100;
 const today=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Asia/Singapore'});
 function date(s){if(!/^\d{4}-\d{2}-\d{2}$/.test(s))throw Error('Enter a valid date');const d=new Date(s+'T12:00:00Z');if(d.toISOString().slice(0,10)!==s)throw Error('Invalid date');return d;}
 function year(s){const d=date(s),y=d.getUTCFullYear()+1,m=d.getUTCMonth(),day=d.getUTCDate();return new Date(Date.UTC(y,m,Math.min(day,new Date(Date.UTC(y,m+1,0)).getUTCDate()))).toISOString().slice(0,10);}
 function validate(p){date(p.start);date(p.due);if(p.end)date(p.end);if(!p.number.trim()||!p.description.trim()||!Number.isFinite(+p.premium)||+p.premium<=0||!Number.isFinite(+p.increment)||+p.increment<0)throw Error('Policy number, description and positive premium are required; increment must be zero or positive.');if(p.due<p.start||p.end&&p.end<p.start)throw Error('Dates must follow the policy start.');}
 function amount(p,due){const a=date(p.start),b=date(due);let months=(b.getUTCFullYear()-a.getUTCFullYear())*12+b.getUTCMonth()-a.getUTCMonth();if(b.getUTCDate()<a.getUTCDate())months--;return money(+p.premium+Math.max(0,months)*(+p.increment||0));}
 function first(p,now=today()){validate(p);const q={...p};while(q.due<now&&q.annual)q.due=year(q.due);if(q.due<now&&!q.annual)q.finished=true;return q;}
 function events(state,until){const out=[];for(const p of state.policies){if(p.finished)continue;let due=p.due;for(let i=0;i<50&&due<=until;i++){if(!p.end||due<=p.end)out.push({id:p.id,due,amount:amount(p,due),description:p.description});if(!p.annual||p.end&&due>=p.end)break;due=year(due);}}return out.sort((a,b)=>a.due.localeCompare(b.due));}
 function batch(state){const e=events(state,'2099-12-31'),due=e[0]?.due||'';return {due,amount:money(e.filter(x=>x.due.slice(0,7)===due.slice(0,7)).reduce((s,x)=>s+x.amount,0))};}
 function parse(v){if(!v)return {version:1,revision:0,linkedName:'',policies:[],history:[]};const s=JSON.parse(v);if(s.version!==1||!Array.isArray(s.policies))throw Error('Insurance data cannot be read');return s;}
 root.FinInsurance={money,today,year,validate,amount,first,events,batch,parse};
})(typeof window==='undefined'?globalThis:window);
