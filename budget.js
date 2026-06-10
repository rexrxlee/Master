let budgetTransactions = [];
let billsBudget = [];
let monthlyBudget = [];

const BUDGET_SHEET = "Budget Setup";

async function loadBudgetPage() {
  try {
    clearOutput();

    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();

    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const transactionSheet = workbook.Sheets[CONFIG.sheetName];
    const budgetSheet = workbook.Sheets[BUDGET_SHEET];

    if (!transactionSheet) throw new Error("Sheet not found: " + CONFIG.sheetName);
    if (!budgetSheet)      throw new Error("Sheet not found: " + BUDGET_SHEET);

    budgetTransactions = readTransactionSheet(transactionSheet);
    billsBudget        = readBudgetSection(budgetSheet, "A2:B13", "Bills");
    monthlyBudget      = readBudgetSection(budgetSheet, "F2:G13", "Monthly Expenses");
    accountsList       = readAccountsSection(budgetSheet, "J2:J10");

    renderBudget();
    log("Budget page loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function clean(value) {
  return String(value ?? "").trim();
}

function readTransactionSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  const headers = rows[0].map(h => clean(h));

  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });
}

function readBudgetSection(sheet, range, type) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, blankrows: false });
  return rows
    .map(row => ({ type, category: clean(row[0]), allocated: toNumber(row[1]) }))
    .filter(item => item.category !== "");
}

function addBudgetItem() {
  const type      = document.getElementById("budgetType").value;
  const category  = clean(document.getElementById("subCategoryInput").value);
  const allocated = Number(document.getElementById("allocatedInput").value);

  if (!category || isNaN(allocated)) { alert("Please enter Sub Category and Allocated."); return; }

  const list = type === "Bills" ? billsBudget : monthlyBudget;
  if (list.length >= 12) { alert("Maximum 12 rows allowed for this section."); return; }

  list.push({ type, category, allocated });
  document.getElementById("subCategoryInput").value = "";
  document.getElementById("allocatedInput").value   = "";
  renderBudget();
}

function deleteBudgetItem(type, index) {
  (type === "Bills" ? billsBudget : monthlyBudget).splice(index, 1);
  renderBudget();
}

function updateBudgetCategory(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].category = clean(value);
  renderBudget();
}

function updateBudgetAllocated(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].allocated = Number(value) || 0;
  renderBudget();
}

function renderBudget() {
  const computedBills   = billsBudget.map(item => computeBudgetRow(item));
  const computedMonthly = monthlyBudget.map(item => computeBudgetRow(item));

  renderBudgetTable("billsTable",   "Bills",            computedBills);
  renderBudgetTable("monthlyTable", "Monthly Expenses", computedMonthly);
  updateBudgetCards(computedBills, computedMonthly);
  renderBudgetVisualPanel(computedBills, computedMonthly);
  renderBudgetProjectionPanel(computedMonthly);
  renderBudgetPressurePanel(computedBills, computedMonthly);
}

function computeBudgetRow(item) {
  const spent = calculateSpentForCurrentMonth(item.type, item.category);
  return { ...item, spent, balance: item.allocated - spent };
}

/**
 * Calculates spending for the current month.
 * Claimable rows are excluded from spending budgets entirely.
 * They still affect account/card balances elsewhere until the reimbursement arrives.
 */
function calculateSpentForCurrentMonth(mainCategory, subCategory) {
  const today        = new Date();
  const currentYear  = today.getFullYear();
  const currentMonth = today.getMonth();

  return budgetTransactions
    .filter(row => {
      const rowDate = parseBudgetDate(row["Date"]);
      if (!rowDate) return false;

      const isCurrentMonth =
        rowDate.getFullYear() === currentYear &&
        rowDate.getMonth()    === currentMonth;

      return (
        isCurrentMonth &&
        clean(row["Main Category"]).toLowerCase() === clean(mainCategory).toLowerCase() &&
        clean(row["Sub Category"]).toLowerCase()  === clean(subCategory).toLowerCase()
      );
    })
    .reduce((sum, row) => sum + getBudgetImpactAmount(row), 0);
}

function renderBudgetTable(tableId, type, rows) {
  const table = document.getElementById(tableId);

  table.innerHTML = `
    <tr>
      <th>${type}</th>
      <th>Allocated</th>
      <th>Spent</th>
      <th>Balance</th>
      <th></th>
    </tr>`;

  let totalAllocated = 0, totalSpent = 0, totalBalance = 0;

  rows.forEach((row, index) => {
    totalAllocated += row.allocated;
    totalSpent     += row.spent;
    totalBalance   += row.balance;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeHtml(row.category)}" onchange="updateBudgetCategory('${type}', ${index}, this.value)"></td>
      <td><input type="number" step="0.01" value="${row.allocated}" onchange="updateBudgetAllocated('${type}', ${index}, this.value)"></td>
      <td>${formatCurrency(row.spent)}</td>
      <td style="color:${row.balance < 0 ? '#c0392b' : 'inherit'}">${formatCurrency(row.balance)}</td>
      <td><button onclick="deleteBudgetItem('${type}', ${index})">Delete</button></td>`;
    table.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.className = "total-row";
  totalRow.innerHTML = `
    <td><strong>Total</strong></td>
    <td><strong>${formatCurrency(totalAllocated)}</strong></td>
    <td><strong>${formatCurrency(totalSpent)}</strong></td>
    <td style="color:${totalBalance < 0 ? '#c0392b' : 'inherit'}"><strong>${formatCurrency(totalBalance)}</strong></td>
    <td></td>`;
  table.appendChild(totalRow);
}

function renderBudgetVisualPanel(billsRows, monthlyRows) {
  const container = document.getElementById("budgetVisualPanel");
  if (!container) return;

  const allRows = [...billsRows, ...monthlyRows];
  const bills = summariseBudgetRows(billsRows);
  const monthly = summariseBudgetRows(monthlyRows);
  const total = summariseBudgetRows(allRows);
  const billCoveragePlan = buildBillReallocationPlan(bills, monthly);
  const billCoverageMap = buildBillCoverageMap(billCoveragePlan);
  const coveredFromMonthly = billCoveragePlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const topOver = allRows
    .map(row => ({ ...row, over: Math.max(0, row.spent - row.allocated) }))
    .filter(row => row.over > 0)
    .sort((a, b) => b.over - a.over || b.spent - a.spent);

  const alertHtml = topOver.length
    ? `<span class="bv-alert">${escapeHtml(topOver[0].category)} is ${formatCurrency(topOver[0].over)} over</span>`
    : `<span class="bv-alert ok">No overspent categories</span>`;

  container.innerHTML = `
    <div class="budget-visual-panel">
      <div class="bv-header">
        <div>
          <h2>Budget Visuals</h2>
          <p>Allocated, spent, reserved, and balance by category. Amber means monthly balance is absorbing bill overspend.</p>
        </div>
        ${alertHtml}
      </div>
      <div class="bv-summary-grid">
        ${renderBudgetSummaryMeter("Total Budget", total)}
        ${renderBudgetSummaryMeter("Bills", bills)}
        ${renderBudgetSummaryMeter("Monthly Expenses", monthly, coveredFromMonthly)}
      </div>
      ${renderBillCoverageBanner(billCoveragePlan)}
      <div class="bv-legend">
        <span><i class="spent"></i> Spent inside budget</span>
        <span><i class="reserve"></i> Reserved for bill overspend</span>
        <span><i class="balance"></i> Balance left</span>
        <span><i class="over"></i> Overspent</span>
      </div>
      <div class="bv-section-grid">
        <div>
          <div class="bv-section-title">Bills</div>
          <div class="bv-rows">${renderBudgetVisualRows(billsRows)}</div>
        </div>
        <div>
          <div class="bv-section-title">Monthly Expenses</div>
          <div class="bv-rows">${renderBudgetVisualRows(monthlyRows, billCoverageMap)}</div>
        </div>
      </div>
    </div>`;
}

function renderBudgetSummaryMeter(label, summary, reserved = 0) {
  const balanceClass = summary.balance < 0 ? "red" : "green";
  const reserveText = reserved > 0
    ? `<span class="bv-reserve-text">${formatCurrency(reserved)} reserved</span>`
    : "";
  return `
    <div class="bv-summary">
      <div class="bv-summary-top">
        <span class="bv-summary-label">${escapeHtml(label)}</span>
        <strong class="bv-summary-value ${balanceClass}">${formatCurrency(summary.balance)}</strong>
      </div>
      ${renderBudgetMeter(summary, reserved)}
      <div class="bv-summary-detail">
        <span>${formatCurrency(summary.spent)} spent</span>
        <span>${formatCurrency(summary.allocated)} allocated</span>
        ${reserveText}
      </div>
    </div>`;
}

function renderBudgetVisualRows(rows, reserveMap = {}) {
  if (!rows.length) return `<div class="bp-muted">No budget rows yet.</div>`;

  return [...rows]
    .map(row => ({ ...row, over: Math.max(0, row.spent - row.allocated) }))
    .sort((a, b) => b.over - a.over || (b.spent / Math.max(b.allocated, 1)) - (a.spent / Math.max(a.allocated, 1)))
    .map(row => {
      const reserved = Math.min(Math.max(0, reserveMap[budgetCategoryKey(row.category)] || 0), Math.max(0, row.balance));
      const freeBalance = Math.max(0, row.balance - reserved);
      const balanceClass = row.balance < 0 ? "red" : "green";
      const rowClass = row.balance < 0 ? "bv-row over" : "bv-row";
      const reserveDetail = reserved > 0
        ? `<span>${formatCurrency(reserved)} reserved for bills · ${formatCurrency(freeBalance)} free</span>`
        : `<span>${formatCurrency(row.allocated)} allocated</span>`;
      return `
        <div class="${rowClass}">
          <div class="bv-row-top">
            <span class="bv-row-name">${escapeHtml(row.category)}</span>
            <strong class="bv-row-balance ${balanceClass}">${formatCurrency(row.balance)}</strong>
          </div>
          ${renderBudgetMeter(row, reserved)}
          <div class="bv-row-detail">
            <span>${formatCurrency(row.spent)} spent</span>
            ${reserveDetail}
          </div>
        </div>`;
    }).join("");
}

function renderBudgetMeter(row, reserved = 0) {
  const spentInside = Math.min(Math.max(0, row.spent), Math.max(0, row.allocated));
  const balanceLeft = Math.max(0, row.allocated - row.spent);
  const reserve = Math.min(Math.max(0, reserved), balanceLeft);
  const freeBalance = Math.max(0, balanceLeft - reserve);
  const overspent = Math.max(0, row.spent - row.allocated);
  const scale = Math.max(spentInside + balanceLeft + overspent, row.allocated, row.spent, 1);
  const spentPct = (spentInside / scale) * 100;
  const reservePct = (reserve / scale) * 100;
  const balancePct = (freeBalance / scale) * 100;
  const overPct = (overspent / scale) * 100;
  const emptyClass = row.allocated <= 0 && row.spent <= 0 ? " empty" : "";

  return `
    <div class="bv-meter${emptyClass}" aria-label="${escapeHtml(row.category || "Budget")} budget meter">
      <span class="bv-seg spent" style="width:${spentPct.toFixed(2)}%;"></span>
      <span class="bv-seg reserve" style="width:${reservePct.toFixed(2)}%;"></span>
      <span class="bv-seg balance" style="width:${balancePct.toFixed(2)}%;"></span>
      <span class="bv-seg over" style="width:${overPct.toFixed(2)}%;"></span>
    </div>`;
}

function renderBillCoverageBanner(plan) {
  const covered = plan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const totalNeeded = covered + plan.uncovered;
  if (totalNeeded <= 0) return "";

  const cutRows = plan.cuts.length
    ? plan.cuts.map(cut => `
        <div class="bv-cover-row">
          <span>${escapeHtml(cut.category)}</span>
          <strong>${formatCurrency(cut.cut)}</strong>
        </div>`).join("")
    : `<div class="bp-muted">No monthly balance available to reserve.</div>`;

  return `
    <div class="bv-cover-banner ${plan.uncovered > 0 ? "danger" : ""}">
      <div class="bv-cover-main">
        <span>Bills overspent</span>
        <strong>${formatCurrency(totalNeeded)}</strong>
      </div>
      <div class="bv-cover-main">
        <span>Reserved from monthly balance</span>
        <strong>${formatCurrency(covered)}</strong>
      </div>
      <div class="bv-cover-list">
        ${cutRows}
        ${plan.uncovered > 0 ? `
          <div class="bv-cover-row danger">
            <span>Still uncovered</span>
            <strong>${formatCurrency(plan.uncovered)}</strong>
          </div>` : ""}
      </div>
    </div>`;
}

function renderBudgetProjectionPanel(monthlyRows) {
  const container = document.getElementById("budgetProjectionPanel");
  if (!container) return;

  const projection = buildMonthlyProjection(monthlyRows);
  const projectedBalanceClass = projection.totalProjectedBalance >= 0 ? "green" : "red";
  const badgeClass = projection.totalProjectedBalance >= 0 ? "" : " over";
  const badgeText = projection.totalProjectedBalance >= 0
    ? `Projected ${formatCurrency(projection.totalProjectedBalance)} under`
    : `Projected ${formatCurrency(Math.abs(projection.totalProjectedBalance))} over`;

  const rowsHtml = projection.rows.length
    ? projection.rows.map(renderProjectionRow).join("")
    : `<div class="bp-muted">No monthly expense budget rows yet.</div>`;

  container.innerHTML = `
    <div class="budget-projection-panel">
      <div class="bproj-header">
        <div>
          <h2>Monthly Expense Projection</h2>
          <p>Expected month-end spend using average daily spending so far this month.</p>
        </div>
        <span class="bproj-badge${badgeClass}">${badgeText}</span>
      </div>
      <div class="bproj-grid">
        <div class="bproj-summary">
          <span>Projected Spend</span>
          <strong>${formatCurrency(projection.totalProjectedSpend)}</strong>
        </div>
        <div class="bproj-summary">
          <span>Monthly Allocation</span>
          <strong>${formatCurrency(projection.totalAllocated)}</strong>
        </div>
        <div class="bproj-summary">
          <span>Projected Balance</span>
          <strong class="${projectedBalanceClass}">${formatCurrency(projection.totalProjectedBalance)}</strong>
        </div>
      </div>
      <div class="bproj-rows">${rowsHtml}</div>
    </div>`;
}

function buildMonthlyProjection(rows) {
  const today = new Date();
  const dayOfMonth = Math.max(1, today.getDate());
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const projectedRows = rows.map(row => {
    const dailyAverage = row.spent / dayOfMonth;
    const projectedSpend = dailyAverage * daysInMonth;
    const projectedBalance = row.allocated - projectedSpend;
    return { ...row, dailyAverage, projectedSpend, projectedBalance };
  }).sort((a, b) => a.projectedBalance - b.projectedBalance || b.projectedSpend - a.projectedSpend);

  return {
    rows: projectedRows,
    dayOfMonth,
    daysInMonth,
    totalAllocated: projectedRows.reduce((sum, row) => sum + row.allocated, 0),
    totalProjectedSpend: projectedRows.reduce((sum, row) => sum + row.projectedSpend, 0),
    totalProjectedBalance: projectedRows.reduce((sum, row) => sum + row.projectedBalance, 0)
  };
}

function renderProjectionRow(row) {
  const balanceClass = row.projectedBalance >= 0 ? "green" : "red";
  const scale = Math.max(row.allocated, row.projectedSpend, 1);
  const fillPct = Math.min(100, (row.projectedSpend / scale) * 100);
  const budgetPct = Math.min(100, (row.allocated / scale) * 100);
  return `
    <div class="bproj-row">
      <div class="bproj-row-top">
        <span class="bproj-row-name">${escapeHtml(row.category)}</span>
        <strong class="bproj-row-val ${balanceClass}">${formatCurrency(row.projectedBalance)}</strong>
      </div>
      <div class="bproj-track">
        <span class="bproj-fill ${row.projectedBalance < 0 ? "over" : ""}" style="width:${fillPct.toFixed(2)}%;"></span>
        <span class="bproj-budget-line" style="left:${budgetPct.toFixed(2)}%;"></span>
      </div>
      <div class="bproj-row-detail">
        <span>${formatCurrency(row.spent)} spent · ${formatCurrency(row.dailyAverage)}/day</span>
        <span>${formatCurrency(row.projectedSpend)} projected vs ${formatCurrency(row.allocated)}</span>
      </div>
    </div>`;
}

async function saveBudgetSetupToExcel() {
  try {
    log("Saving Budget Setup to Excel...");
    await writeBudgetSetupRange("A2:B13", buildSaveValues(billsBudget));
    await writeBudgetSetupRange("F2:G13", buildSaveValues(monthlyBudget));
    alert("Budget saved to Excel.");
    log("Budget saved.");
  } catch (err) {
    log("SAVE ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function buildSaveValues(list) {
  const values = list.map(item => [item.category, item.allocated]);
  while (values.length < 12) values.push(["", ""]);
  return values;
}

function updateBudgetCards(billsRows, monthlyRows) {
  const rows = [...billsRows, ...monthlyRows];
  const bills = summariseBudgetRows(billsRows);
  const monthly = summariseBudgetRows(monthlyRows);
  const billReserve = buildBillReallocationPlan(bills, monthly)
    .cuts
    .reduce((sum, cut) => sum + cut.cut, 0);
  const monthlyFreeBalance = monthly.balance - billReserve;
  const totalAllocated = rows.reduce((s, r) => s + r.allocated, 0);
  const totalSpent     = rows.reduce((s, r) => s + r.spent,     0);
  const totalBalance   = rows.reduce((s, r) => s + r.balance,   0);

  const foodRow     = monthlyRows.find(r => clean(r.category).toLowerCase() === "food");
  const foodBalance = foodRow ? foodRow.balance : 0;
  const daysLeft    = getDaysRemainingInMonth();

  setCurrencyValue("totalAllocated", totalAllocated);
  setCurrencyValue("totalSpent", totalSpent, "red");
  setCurrencyValue("totalBalance", totalBalance, totalBalance < 0 ? "red" : "green");
  setCurrencyValue("foodPerDay", foodBalance / daysLeft, foodBalance < 0 ? "red" : "green");
  setCurrencyValue("monthlyPerDay", monthlyFreeBalance / daysLeft, monthlyFreeBalance < 0 ? "red" : "green");
  setTextValue(
    "monthlyPerDayNote",
    billReserve > 0
      ? `Monthly expenses only; ${formatCurrency(billReserve)} reserved for bills.`
      : "Monthly expenses only; bills excluded."
  );
}

function setCurrencyValue(id, value, tone = "") {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerText = formatCurrency(value);
  el.style.color = tone === "red" ? "var(--red)" : tone === "green" ? "var(--green)" : "";
}

function setTextValue(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerText = value;
}

function renderBudgetPressurePanel(billsRows, monthlyRows) {
  const container = document.getElementById("budgetPressurePanel");
  if (!container) return;

  const bills = summariseBudgetRows(billsRows);
  const monthly = summariseBudgetRows(monthlyRows);
  const total = {
    allocated: bills.allocated + monthly.allocated,
    spent: bills.spent + monthly.spent,
    balance: bills.balance + monthly.balance
  };
  total.over = Math.max(0, total.spent - total.allocated);

  const isOverBudget = total.balance < 0;
  const overBills = bills.rows.filter(row => row.over > 0).slice(0, 3);
  const overMonthly = monthly.rows.filter(row => row.over > 0).slice(0, 4);
  const billsCategoryOver = bills.rows.reduce((sum, row) => sum + row.over, 0);
  const monthlyCategoryOver = monthly.rows.reduce((sum, row) => sum + row.over, 0);
  const topOver = [
    ...overBills.map(row => ({ ...row, group: "Bills" })),
    ...overMonthly.map(row => ({ ...row, group: "Monthly" }))
  ].sort((a, b) => b.over - a.over || b.spent - a.spent).slice(0, 5);

  const overRowsHtml = topOver.length
    ? topOver.map(row => `
        <div class="bp-over-row">
          <span>${escapeHtml(row.category)} <small>${row.group}</small></span>
          <strong>${formatCurrency(row.over)} over</strong>
        </div>`).join("")
    : `<div class="bp-muted">No categories are currently over budget.</div>`;

  const monthlyNames = overMonthly.map(row => row.category).slice(0, 3).join(", ");
  const reallocationPlan = buildBillReallocationPlan(bills, monthly);
  const coveredFromMonthly = reallocationPlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const billCoverageTotal = coveredFromMonthly + reallocationPlan.uncovered;
  const billReserveHtml = billCoverageTotal > 0
    ? `<div class="bp-reserve-strip ${reallocationPlan.uncovered > 0 ? "danger" : ""}">
        <div>
          <span>Bills overspent</span>
          <strong class="red">${formatCurrency(billCoverageTotal)}</strong>
        </div>
        <div>
          <span>Reserved from monthly balance</span>
          <strong>${formatCurrency(coveredFromMonthly)}</strong>
        </div>
        <div>
          <span>Still uncovered</span>
          <strong class="${reallocationPlan.uncovered > 0 ? "red" : "green"}">${formatCurrency(reallocationPlan.uncovered)}</strong>
        </div>
      </div>`
    : "";
  const actions = [];

  if (billsCategoryOver > 0) {
    const billAction = bills.balance < 0
      ? "Top up the bill budget or reduce goal allocations before cutting required payments."
      : "Move allocation from under-used bill lines, or top up the bill budget if this is a permanent increase.";
    const reallocationHtml = reallocationPlan.cuts.length
      ? `<div class="bp-cut-list">
          ${reallocationPlan.cuts.map(cut => `
            <div class="bp-cut-row">
              <span>${escapeHtml(cut.category)} <small>least spent so far</small></span>
              <strong>Cut ${formatCurrency(cut.cut)}</strong>
            </div>`).join("")}
          ${reallocationPlan.uncovered > 0 ? `
            <div class="bp-cut-row">
              <span>Still uncovered</span>
              <strong>${formatCurrency(reallocationPlan.uncovered)}</strong>
            </div>` : ""}
        </div>`
      : "";
    actions.push(`
      <div class="bp-action danger">
        <strong>Protect bills first.</strong>
        Bill categories are ${formatCurrency(billsCategoryOver)} over. These are fixed commitments. ${billAction}
        ${reallocationHtml}
      </div>`);
  } else {
    actions.push(`
      <div class="bp-action ok">
        <strong>Bills are covered.</strong>
        Bills still have ${formatCurrency(Math.max(0, bills.balance))} left. Keep that reserved before sending extra money to goals.
      </div>`);
  }

  if (monthlyCategoryOver > 0) {
    actions.push(`
      <div class="bp-action warn">
        <strong>Trim flexible spend next.</strong>
        Monthly expense categories are ${formatCurrency(monthlyCategoryOver)} over${monthlyNames ? `, led by ${escapeHtml(monthlyNames)}` : ""}. Put a short cap on those categories for the rest of the month.
      </div>`);
  } else {
    actions.push(`
      <div class="bp-action ok">
        <strong>Monthly expenses are within plan.</strong>
        Flexible spending still has ${formatCurrency(Math.max(0, monthly.balance))} left.
      </div>`);
  }

  if (isOverBudget) {
    actions.push(`
      <div class="bp-action danger">
        <strong>You are drawing from savings or future card payment capacity.</strong>
        The month is ${formatCurrency(total.over)} over budget. Pause non-urgent goal top-ups until this is absorbed.
      </div>`);
  }

  container.innerHTML = `
    <div class="budget-pressure-panel ${isOverBudget ? "danger" : "ok"}">
      <div class="budget-pressure-header">
        <h2>Budget Pressure</h2>
        <span>Current month risk and next actions</span>
      </div>
      <div class="bp-grid">
        <div class="bp-summary">
          <span class="bp-label">Budget Balance</span>
          <strong class="${isOverBudget ? "red" : "green"}">${formatCurrency(total.balance)}</strong>
          <span class="bp-detail">${formatCurrency(total.spent)} spent vs ${formatCurrency(total.allocated)} allocated</span>
        </div>
        <div class="bp-summary">
          <span class="bp-label">Bills Balance</span>
          <strong class="${bills.balance < 0 ? "red" : "green"}">${formatCurrency(bills.balance)}</strong>
          <span class="bp-detail">Fixed commitments</span>
        </div>
        <div class="bp-summary">
          <span class="bp-label">Monthly Expenses Balance</span>
          <strong class="${monthly.balance < 0 ? "red" : "green"}">${formatCurrency(monthly.balance)}</strong>
          <span class="bp-detail">More flexible categories</span>
        </div>
      </div>
      ${billReserveHtml}
      <div class="bp-content">
        <div>
          <div class="bp-section-title">Over-budget categories</div>
          ${overRowsHtml}
        </div>
        <div class="bp-actions">
          <div class="bp-section-title">What to do</div>
          ${actions.join("")}
        </div>
      </div>
    </div>`;
}

function buildBillReallocationPlan(bills, monthly) {
  let remaining = bills.rows.reduce((sum, row) => sum + row.over, 0);
  const cuts = [];
  if (remaining <= 0) return { cuts, uncovered: 0 };

  monthly.rows
    .filter(row => row.balance > 0)
    .sort((a, b) => a.spent - b.spent || b.balance - a.balance)
    .forEach(row => {
      if (remaining <= 0) return;
      const cut = Math.min(row.balance, remaining);
      if (cut <= 0) return;
      cuts.push({ category: row.category, cut });
      remaining = Math.max(0, remaining - cut);
    });

  return { cuts, uncovered: remaining };
}

function buildBillCoverageMap(plan) {
  return plan.cuts.reduce((map, cut) => {
    const key = budgetCategoryKey(cut.category);
    map[key] = (map[key] || 0) + cut.cut;
    return map;
  }, {});
}

function budgetCategoryKey(category) {
  return clean(category).toLowerCase();
}

function summariseBudgetRows(rows) {
  const sortedRows = [...rows]
    .map(row => ({ ...row, over: Math.max(0, row.spent - row.allocated) }))
    .sort((a, b) => b.over - a.over || b.spent - a.spent);
  const allocated = sortedRows.reduce((sum, row) => sum + row.allocated, 0);
  const spent = sortedRows.reduce((sum, row) => sum + row.spent, 0);
  const balance = allocated - spent;
  return {
    rows: sortedRows,
    allocated,
    spent,
    balance,
    over: Math.max(0, spent - allocated)
  };
}

function getDaysRemainingInMonth() {
  const today   = new Date();
  const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  return Math.max(lastDay.getDate() - today.getDate() + 1, 1);
}

function parseBudgetDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    return new Date(d.y, d.m - 1, d.d);
  }
  const s = String(value).trim();
  // DD/MM/YYYY — our stored format
  const ddmmyyyy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyy) return new Date(Number(ddmmyyyy[3]), Number(ddmmyyyy[2])-1, Number(ddmmyyyy[1]));
  // YYYY-MM-DD ISO — parse parts manually, never pass to new Date() to avoid UTC shift
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3]));
  return null;
}

function getAmount(value) {
  const n = Number(String(value).replace(/[$,]/g, ""));
  return isNaN(n) ? 0 : Math.abs(n);
}

function getBudgetImpactAmount(row) {
  const amount = getAmount(row["Amount"]);
  const claimableKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claimable") || "Claimable";
  const claimable = clean(row[claimableKey]).toLowerCase();
  if (claimable !== "yes") return amount;
  return 0;
}

function toNumber(value) {
  const n = Number(String(value).replace(/[$,]/g, ""));
  return isNaN(n) ? 0 : n;
}

function formatCurrency(value) {
  return value.toLocaleString("en-SG", {
    style: "currency", currency: "SGD",
    minimumFractionDigits: 2, maximumFractionDigits: 2
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&","&amp;").replaceAll("<","&lt;")
    .replaceAll(">","&gt;").replaceAll('"',"&quot;");
}

let accountsList = [];

function readAccountsSection(sheet, range) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, blankrows: false });
  return rows.map(row => clean(row[0])).filter(name => name !== "");
}
