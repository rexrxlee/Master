const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const context = vm.createContext({console,Date,Intl});
for (const name of ['tax-engine.js','tax-storage.js']) vm.runInContext(fs.readFileSync(path.join(__dirname,'..',name),'utf8'),context);
const {FinTax:T,FinTaxStore:S} = context;
const input = (employment,extras={}) => ({...T.newRecord(2026).inputs,employment,...extras});
for (const [income,tax] of [[0,0],[20000,0],[30000,200],[40000,550],[80000,3350],[120000,7950],[160000,13950],[200000,21150],[240000,28750],[280000,36550],[320000,44550],[500000,84150],[1000000,199150],[1000100,199174]]) {
  assert.equal(T.calculate(input(income)).netTax,tax,`Bracket ${income}`);
}
assert.equal(T.calculate(input(100000,{cpf:90000})).chargeable,20000);
assert.equal(T.calculate(input(40250,{donations:250})).netTax,550);
assert.equal(T.calculate(input(40000,{ptr:1000})).netTax,0);
assert.equal(T.calculate(input(40000,{ptr:1000})).rebateRemaining,450);
assert.equal(T.calculate(input(40000,{trade:-10000})).netTax,200);
assert.throws(() => T.calculate(input('')),/employment income/);
assert.throws(() => T.calculate(input(-1)),/valid annual amount/);
assert.throws(() => T.calculate(input(Infinity)),/valid annual amount/);
assert.equal(T.dueDate('2026-01-31'),'2026-02-28');
assert.equal(T.dueDate('2028-01-31'),'2028-02-29');
assert.equal(T.dueDate('2026-12-31'),'2027-01-31');
assert.equal(T.dueDate(''),'');
assert.throws(() => T.dueDate('2026-02-30'),/valid tax bill date/);

const state = S.parse('');
state.years[2026].inputs = input(92400,{cpf:18480,earned:1000,ptr:100});
state.years[2026].confirmed = true;
state.years[2026].billDate = '2026-10-02';
const rows = Array.from({length:19},()=>[]);
rows[0] = ['Emergency Fund',50000,100,0,'2026-01-01','2029-01-01','High'];
const before = JSON.stringify(rows);
const created = S.prepare(state,rows,false,'2026-10-03');
assert.equal(JSON.stringify(rows),before);
assert.equal(created.index,1);
assert.equal(created.goal[0],'Personal Tax YA 2026');
assert.equal(created.goal[1],2754.4);
assert.equal(created.goal[5],'2026-11-02');
rows[1] = created.goal;
rows[1][2] = 300;
const rollover = S.prepare(created.state,rows,true,'2026-10-03');
assert.equal(rollover.state.activeYear,2027);
assert.equal(rollover.goal[0],'Personal Tax YA 2027');
assert.equal(rollover.goal[2],0);
assert.equal(rollover.goal[5],'');
assert.equal(rollover.state.years[2026].paid,true);
assert.equal(rollover.state.years[2027].confirmed,false);
assert.equal(rollover.state.years[2027].inputs.ptr,0);
assert.equal(rollover.state.years[2027].inputs.cpfTopup,0);
assert.throws(() => T.completeYear(rollover.state,'2026-10-03'),/Confirm/);
assert.throws(() => S.prepare(state,Array.from({length:19},(_,i)=>['Goal '+i]),false,'2026-10-03'),/19 goal slots/);

(async () => {
  let stored = S.parse(''), storedRows = Array.from({length:19},()=>[]), failAt = -1;
  context.readBudgetSetupRange = async () => ({values:storedRows.map((row,i)=>{
    const wide = Array.from({length:19},(_,j)=>row[j]??'');
    if(i===0) wide[18]=JSON.stringify(stored);
    return wide;
  })});
  context.writeExcelRanges = async updates => {
    for (let i=0;i<updates.length;i++) {
      if(i===failAt) throw new Error('network interrupted');
      const u=updates[i];
      if(u.rangeAddress==='AK2') stored=JSON.parse(u.values[0][0]);
      else storedRows[Number(u.rangeAddress.match(/^S(\d+)/)[1])-2]=Array.from(u.values[0]);
    }
  };
  const draft=JSON.parse(JSON.stringify(state));
  failAt=1;
  await assert.rejects(S.save(draft,0,false),/network interrupted/);
  assert.ok(stored.pendingGoal);
  failAt=-1;
  const recovered=await S.recover();
  assert.equal(recovered.pendingGoal,undefined);
  assert.equal(storedRows[0][0],'Personal Tax YA 2026');
  await assert.rejects(S.save(draft,0,false),/changed in another page/);
  const paid=await S.save(recovered,1,true);
  assert.equal(paid.activeYear,2027);
  assert.equal(storedRows[0][0],'Personal Tax YA 2027');
  assert.equal(storedRows[0][2],0);
  console.log('Tax brackets, relief cap, PTR, deadlines, goal linking, rollover, stale saves and interrupted-save recovery passed.');
})().catch(error=>{console.error(error);process.exitCode=1;});
