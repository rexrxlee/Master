const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const c=vm.createContext({});vm.runInContext(fs.readFileSync('cash-rules.js','utf8')+';this.rules=CashRules;',c);const R=c.rules;
const accounts=[{name:'Bank',type:'Savings'},{name:'Restricted',type:'Savings'},{name:'Card',type:'Credit Card'},{name:'Other card',type:'Credit Card'}];
const rows=[{Account:'Card',Transaction:'Opening Balance',Date:'2026-01-01',Amount:1000},{Account:'Card',Date:'2026-02-01',Amount:10,'Main Category':'Income'},{Account:'Card',Date:'2026-02-02',Amount:-20,'Main Category':'Bills'},{Account:'Other card',Transaction:'Opening Balance',Amount:-100}];
assert.equal(R.cardDebt(accounts,rows,v=>v?new Date(v):null),970);
const bills=[{balance:200,fundingAccount:'Restricted'},{balance:40,fundingAccount:'Card'},{balance:25,fundingAccount:''},{balance:-5,fundingAccount:'Bank'}];
assert.equal(R.billReserve(bills,['Bank'],accounts),65);assert.equal(R.billReserve(bills,['Bank','Restricted'],accounts),265);
assert.equal(R.monthlyReserve([{balance:100.01},{balance:-20}]),80.01);
assert.equal(R.pendingClaims([{'Claim Status':'Pending',Amount:200,'Claim Amount':50,'Claim Account':'Restricted'},{'Claim Status':'Claimed',Amount:100}]),50);
// Both page entry points must call the same engine.
for(const [file,name,scope] of [['goals.js','computeCCOwed','allAccounts'],['budget.js','computeBudgetCreditCardOwed','budgetAccountTypes']]){
 const text=fs.readFileSync(file,'utf8'),start=text.indexOf('function '+name+'('),end=text.indexOf('\nfunction ',start+1);
 c[scope]=accounts;c.budgetTransactions=rows;c.parseBudgetDate=c.parseExcelDate=v=>v?new Date(v):null;
 vm.runInContext(text.slice(start,end),c);assert.equal(c[name](rows),970);
}
console.log('Shared card credits, refunds, per-card debt floors, bill account scope, pooled monthly reserves and pending claims passed.');

assert.equal(R.planningTotal(2100.57,400.32),2500.89);
