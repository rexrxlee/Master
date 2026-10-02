const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
function context() {
  const storage = new Map();
  const ctx = vm.createContext({ console, setTimeout, clearTimeout, performance,
    localStorage: { getItem: key => storage.get(key) ?? null, setItem: (key, value) => storage.set(key, value), removeItem: key => storage.delete(key) },
    document: { readyState: 'loading', addEventListener() {} },
    CONFIG: { filePath: 'test.xlsx', sheetName: 'Transactions' },
    msal: { PublicClientApplication: class {} } });
  ctx.window = ctx;
  ctx.addEventListener = () => {};
  return ctx;
}
function load(ctx, file) { vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), ctx); }

async function main() {
  const ctx = context();
  load(ctx, 'business.js');
  load(ctx, 'goals.js');
  load(ctx, 'goals-compact.js');
  vm.runInContext(`
    const todayTest = new Date();
    const monthTest = todayTest.getMonth();
    const yearTest = todayTest.getFullYear();
    const dateTest = (offset) => toDateInputValue(new Date(yearTest, monthTest + offset, 1));
    allAccounts = [{ name: 'Savings', type: 'Savings' }, { name: 'Card', type: 'Credit Card' }];
    goalSavingsAccts = ['Savings'];
    savingsBalances = { Savings: 8000 }; ccOwed = 1200;
    budgetSummary = { billsRows: [{category:'Rent', allocated:1000, fundingAccount:'Savings'}], monthlyRows:[{category:'Food',allocated:1000}], billsTotal:1000, monthlyTotal:1000 };
    historicalStats.avgMonthlyIncome = 6000;
    goalsData = ['Tax', 'Trip', 'Emergency'].map((name, i) => ({name, target: (i+1)*3000, manualSaved:500, monthlyAlloc:0, startDate:dateTest(0), endDate:dateTest(12+i*6), urgency:'Medium', priority:0, goalBuffer:5}));
    allTxForGoals = Array.from({length:3000}, (_, i) => ({Date:dateTest(-1-i%12), Account:'Savings', Amount:10, 'Main Category':'Monthly Expenses', 'Sub Category':'Food'}));
    allTxForGoals.push({Date:dateTest(1), Account:'Savings', Amount:6000, 'Main Category':'Income', 'Sub Category':'Salary'});
    function runForecasts() {
      const outputs = [];
      for (let i=0;i<8;i++) {
        goalsData[0].manualSaved = i*100;
        outputs.push(buildGoalProjectionModel(18,60));
      }
      return JSON.stringify(outputs, (key, value) => key === 'today' ? undefined : value);
    }
  `, ctx);
  let start = performance.now();
  const baseline = vm.runInContext('runForecasts()', ctx);
  const before = performance.now() - start;
  start = performance.now();
  const cached = vm.runInContext('withGoalCalculationCache(runForecasts)', ctx);
  const after = performance.now() - start;
  assert.equal(cached, baseline, 'Every forecast value must match with caching enabled');
  assert.equal(vm.runInContext('goalCalculationCache', ctx), null);
  vm.runInContext(`try { withGoalCalculationCache(() => { throw new Error('test'); }); } catch {}`, ctx);
  assert.equal(vm.runInContext('goalCalculationCache', ctx), null, 'Clear on exceptions');
  const oldCash = vm.runInContext('withGoalCalculationCache(() => computeDeployableBalance().deployable)', ctx);
  vm.runInContext('savingsBalances.Savings += 10000', ctx);
  assert.ok(vm.runInContext('withGoalCalculationCache(() => computeDeployableBalance().deployable)', ctx) > oldCash, 'Next refresh must see changed inputs');
  console.log('Forecast equality passed; eight projections: ' + before.toFixed(0) + 'ms -> ' + after.toFixed(0) + 'ms');

  vm.runInContext(`
    XLSX = { utils: { sheet_to_json: () => [['Emergency Fund',50000,0,0,'2026-10-01','2029-01-01','High','', '#000',0,0]] } };
    const reloadedEmergency = readGoalsFromSheet({AJ2:{v:12}})[0];
  `, ctx);
  assert.equal(vm.runInContext('reloadedEmergency.target', ctx), 50000, 'Saved amount must not be overwritten by legacy months');
  console.log('Emergency goal amount preserved');

  const plan = context();
  load(plan, 'dashboard.js');
  vm.runInContext(`
    const chartTest = {innerHTML:''};
    document.getElementById = () => chartTest;
    financePlanEmergencyTarget = 50000;
    financePlanEmergencyDeadline = '2029-01-01';
    renderPlanMilestones('');
  `, plan);
  assert.match(vm.runInContext('chartTest.innerHTML', plan), /Current age required/);
  vm.runInContext("renderPlanMilestones('32')", plan);
  const chartHtml = vm.runInContext('chartTest.innerHTML', plan);
  assert.match(chartHtml, /<span>32<\/span><span>33<\/span>/);
  assert.match(chartHtml, /50,000/);
  assert.match(chartHtml, /2029/);
  console.log('Milestone annual age ticks, target label and missing-age state passed');

  const io = context();
  load(io, 'onedrive.js');
  vm.runInContext(`
    let patches = [], rebuilds = 0, stored = null, failAt = -1;
    let fixture = { Sheets: { 'Budget Setup': {} } };
    withExcelBusy = async (_, work) => work();
    getToken = async () => 'test';
    getEncodedExcelPath = () => 'test.xlsx';
    readExcelDownloadCache = async () => JSON.stringify(fixture);
    invalidateExcelDownloadCache = async () => { excelCacheGeneration++; stored = null; };
    storeExcelDownloadCache = async value => { stored = value; };
    graphPatch = async (url, token, body) => { patches.push(body); if (patches.length === failAt) throw new Error('save failed'); return body; };
    XLSX = { read: JSON.parse, write: workbook => { rebuilds++; return JSON.stringify(workbook); }, utils: { sheet_add_aoa: (sheet, values, options) => { sheet[options.origin] = values; } } };
    const updatesTest = ['S2:AC20','AG2:AG2','AF2:AF2'].map(rangeAddress => ({sheetName:'Budget Setup',rangeAddress,values:[[rangeAddress]]}));
  `, io);
  await vm.runInContext('writeExcelRanges(updatesTest)', io);
  assert.equal(vm.runInContext('patches.length', io), 3);
  assert.equal(vm.runInContext('rebuilds', io), 1);
  assert.equal(Object.keys(JSON.parse(vm.runInContext('stored', io)).Sheets['Budget Setup']).length, 3);
  vm.runInContext('patches = []; failAt = 2; stored = null;', io);
  await assert.rejects(vm.runInContext('writeExcelRanges(updatesTest)', io), /save failed/);
  assert.equal(vm.runInContext('stored', io), null);
  vm.runInContext('failAt = -1; fixture.Sheets["Budget Setup"].A1 = {f:"SUM(B1:B2)"};', io);
  await vm.runInContext('writeExcelRanges(updatesTest)', io);
  assert.equal(vm.runInContext('stored', io), null, 'Formula workbooks require server refresh');
  console.log('Save batching, failure propagation, queue recovery and formula invalidation passed');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
