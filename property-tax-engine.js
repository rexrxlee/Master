(function(root){
  const money=n=>Math.round((n+Number.EPSILON)*100)/100;
  const today=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Asia/Singapore'});
  function number(v,label){if(v===''||v==null||!Number.isFinite(+v)||+v<0)throw Error('Enter '+label+' (zero is allowed).');return +v;}
  function dueDate(s){if(!/^\d{4}-\d{2}-\d{2}$/.test(s)||!Number.isFinite(Date.parse(s+'T12:00:00Z'))||new Date(s+'T12:00:00Z').toISOString().slice(0,10)!==s)throw Error('Enter a valid deadline.');return s;}
  function calculate(r){
    if(!Number.isInteger(+r.year)||+r.year<2026||+r.year>2100)throw Error('Choose a tax year from 2026 to 2100.');
    if(r.actual!==''&&r.actual!=null)return {grossTax:null,rebate:0,netTax:money(number(r.actual,'the net amount on your bill')),actual:true};
    if((r.av===''||r.av==null)&&r.carryEstimate!=null)return {grossTax:null,rebate:0,netTax:money(number(r.carryEstimate,'the provisional amount')),provisional:true};
    if(!r.ownerOccupied)throw Error('This calculator covers owner-occupied HDB flats. Enter the actual bill for other cases.');
    const av=number(r.av,'the Annual Value from IRAS');
    const bands=[[12000,0],[40000,.04],[50000,.06],[75000,.10],[85000,.14],[100000,.20],[140000,.26],[Infinity,.32]];
    let previous=0,gross=0;for(const [limit,rate] of bands){gross+=Math.max(0,Math.min(av,limit)-previous)*rate;previous=limit;if(av<=limit)break;}
    const grossTax=money(gross),rebate=+r.year===2026?money(grossTax*.15):0;
    return {grossTax,rebate,netTax:money(grossTax-rebate),actual:false};
  }
  function newRecord(year){return {year:+year,av:'',actual:'',ownerOccupied:true,dueOverride:'',firstBill:false,paid:false};}
  function deadline(r){if(r.dueOverride)return dueDate(r.dueOverride);if(r.firstBill)return '';return r.year+'-01-31';}
  function goalName(y){return 'Property Tax '+y;}
  function completeYear(state,now=today()){
    const next=JSON.parse(JSON.stringify(state)),r=next.years[next.activeYear];
    if(r.paid)throw Error('This year is already paid.');
    r.finalEstimate=calculate(r).netTax;r.paid=true;r.paidDate=now;
    next.activeYear=+next.activeYear+1;
    next.years[next.activeYear]={...newRecord(next.activeYear),av:r.av,ownerOccupied:r.ownerOccupied,...((r.av===''||!r.ownerOccupied)?{av:'',carryEstimate:r.finalEstimate}:{})};
    return next;
  }
  root.FinPropertyTax={money,today,dueDate,calculate,newRecord,deadline,goalName,completeYear};
})(typeof window==='undefined'?globalThis:window);
