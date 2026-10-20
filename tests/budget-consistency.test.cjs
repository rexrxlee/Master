const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
function extract(file, name) {
  const source = fs.readFileSync(file, 'utf8');
  const start = source.indexOf('function ' + name + '(');
  assert.ok(start >= 0, name);
  const end = source.indexOf('\nfunction ', start + 1);
  return source.slice(start, end < 0 ? undefined : end);
}
const context = vm.createContext({
  clean: value => String(value ?? '').trim(),
  formatCurrency: value => value.toFixed(2), escapeHtml: value => value,
  budgetTransactions: [], configuredAccounts: [], monthlyExpenseBudgetRows: [],
  parseExcelDate: value => new Date(value), isBusinessTransaction: row => !!row.business,
  getAmount: value => Number(value) || 0,
  getRowValue: (row, key) => row[key],
  document: { getElementById: () => null }
});
for (const name of ['summariseBudgetRows', 'calculateBillCoverage', 'renderBudgetOverspendBanner', 'budgetCategoryKey', 'computeBudgetSectionRows', 'updateBudgetCards']) {
  vm.runInContext(extract('budget.js', name), context);
}
for (const name of ['isClaimableRow', 'getClaimAmount', 'getMonthlyBudgetStatus', 'computeMonthlyExpenseBudgetPosition']) {
  vm.runInContext(extract('dashboard.js', name), context);
}
const row = (category, allocated, spent) => ({ category, allocated, spent, balance: allocated - spent });
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} != ${expected}`);
test('screenshot reconciles category overruns, monthly deficit, bills and total', () => {
  const monthly = [row('Others', 0, 255.89), row('Pets', 0, 119.73), row('Child', 0, 23.05), row('Other allocations', 2250, 1965.72)];
  const bills = [row('Netflix', 35, 37.96), row('Other bills', 1666.23, 978.84)];
  const m = context.summariseBudgetRows(monthly), b = context.summariseBudgetRows(bills);
  near(m.categoryOver, 398.67); near(m.positiveBalance, 284.28); near(m.over, 114.39);
  near(b.balance + m.balance, 570.04);
  const coverage = context.calculateBillCoverage(bills, monthly);
  near(coverage.billsLeft, 687.39); near(coverage.makeupAmount, 117.35);
  const html = context.renderBudgetOverspendBanner(b, m);
  for (const number of ['398.67', '284.28', '114.39']) assert.ok(html.includes(number));
});
test('bill top-up handles monthly surplus, bill overruns, empty and fully spent budgets', () => {
  for (const [bills, monthly, expected] of [
    [[row('B', 100, 0)], [row('M', 100, 50)], [100, 0, 0]],
    [[row('B', 100, 150), row('C', 200, 0)], [row('M', 100, 120)], [200, 20, 70]],
    [[], [], [0, 0, 0]],
    [[row('B', 100, 100)], [row('M', 100, 120)], [0, 20, 20]]
  ]) Object.values(context.calculateBillCoverage(bills, monthly)).forEach((value, i) => near(value, expected[i]));
});
test('remaining budget offsets category overruns without hiding any category', () => {
  const m = context.summariseBudgetRows([...Array.from({ length: 7 }, (_, i) => row('Category' + i, 0, 1)), row('Unused', 20, 0)]);
  near(m.over, 0);
  const html = context.renderBudgetOverspendBanner({}, m);
  assert.ok(html.includes('Category6')); assert.ok(html.includes('<strong>0.00</strong>'));
});
test('budget includes deleted and uncategorised categories without changing saved setup', () => {
  context.budgetTransactions = [{ 'Main Category': 'Monthly Expenses', 'Sub Category': 'Deleted' }, { 'Main Category': 'Monthly Expenses', 'Sub Category': '' }];
  context.computeBudgetRow = item => ({ ...item, spent: 10, balance: item.allocated - 10 });
  const input = [{ category: 'Food', allocated: 100 }];
  const rows = context.computeBudgetSectionRows(input, 'Monthly Expenses');
  assert.equal(rows.length, 3); assert.equal(rows[1].isUnbudgeted, true); assert.equal(input.length, 1);
  near(context.summariseBudgetRows(rows).spent, 30);
});
test('dashboard retains personal portion of partial claims, excludes business, floors daily allowance', () => {
  context.monthlyExpenseBudgetRows = [{ category: 'Food', allocated: 10 }];
  const today = new Date();
  const tx = (amount, claim, business = false) => ({ Date: today, 'Main Category': 'Monthly Expenses', 'Sub Category': 'Food', Amount: amount, Claimable: claim == null ? '' : 'Yes', 'Claim Amount': claim, business });
  const result = context.computeMonthlyExpenseBudgetPosition([tx(100, 40), tx(50, 50), tx(100, null, true)]);
  near(result.spent, 60); near(result.balance, -50); near(result.leftPerDay, 0);
});
test('food allowance never exceeds pooled monthly allowance or goes negative', () => {
  const values = {};
  context.setCurrencyValue = (id, value) => { values[id] = value; };
  context.setTextValue = () => {};
  context.getDaysRemainingInMonth = () => 10;
  for (const [foodSpent, otherSpent, expected] of [[0, 150, 5], [110, 0, 0], [0, 250, 0]]) {
    context.updateBudgetCards([], [row('Food', 100, foodSpent), row('Other', 100, otherSpent)]);
    near(values.foodPerDay, expected);
  }
});
