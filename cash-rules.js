// Shared cash protections for Budget and Goals. All amounts use integer cents.
const CashRules = (() => {
  const key=v=>String(v??'').trim().toLowerCase();
  const cents=v=>Math.round((Number(String(v??0).replace(/[$,]/g,''))||0)*100);
  function cardDebt(accounts,transactions,parseDate) {
    const savings=new Set(accounts.filter(a=>key(a.type||'Savings')==='savings').map(a=>key(a.name)));
    const configured=accounts.filter(a=>key(a.type)==='credit card').map(a=>key(a.name));
    const names=[...new Set(transactions.map(r=>key(r.Account)).filter(Boolean))].filter(n=>configured.length?configured.includes(n):!savings.has(n));
    return names.reduce((total,name)=>{
      const rows=transactions.filter(r=>key(r.Account)===name), opening=rows.find(r=>key(r.Transaction)==='opening balance');
      const date=opening?parseDate(opening.Date):null;
      let owed=opening?cents(opening.Amount):0;
      for(const r of rows){
        if(key(r.Transaction)==='opening balance')continue;
        if(date){const d=parseDate(r.Date);if(!d||d<date)continue;}
        const main=key(r['Main Category']),sub=key(r['Sub Category']),amount=cents(r.Amount);
        if(main==='income')owed-=Math.abs(amount);
        else if(main==='transfer'){
          if(['transfer in','cc payment in'].includes(sub))owed-=Math.abs(amount);
          else if(['transfer out','cc payment out'].includes(sub))owed+=Math.abs(amount);
        } else owed+=amount;
      }
      // A credit on one card cannot settle a different card's debt.
      return total+Math.max(0,owed);
    },0)/100;
  }
  function protectsBill(account,selected,accounts) {
    if(!key(account))return true; // Unknown funding must not silently release cash.
    return selected.some(n=>key(n)===key(account)) || accounts.some(a=>key(a.name)===key(account)&&key(a.type)==='credit card');
  }
  // Household obligations stay protected even when paid from an unselected account.
  function billReserve(rows,selected,accounts){return rows.reduce((sum,r)=>sum+Math.max(0,cents(r.balance)),0)/100;}
  function planningTotal(cash,claims){return (cents(cash)+cents(claims))/100;}

  function monthlyReserve(rows){return Math.max(0,rows.reduce((sum,r)=>sum+cents(r.balance),0))/100;}
  function pendingClaims(rows){
    return rows.reduce((sum,r)=>{
      const field=n=>r[Object.keys(r).find(k=>key(k).replace(/[^a-z]/g,'')===n)];
      if(key(field('claimstatus'))!=='pending')return sum;
      const amount=Math.abs(cents(field('amount'))),claim=Math.abs(cents(field('claimamount')));
      return sum+Math.min(amount,claim||amount);
    },0)/100;
  }
  return {planningTotal,cents,cardDebt,protectsBill,billReserve,monthlyReserve,pendingClaims};
})();
