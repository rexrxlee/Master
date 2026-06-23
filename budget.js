let budgetTransactions = [];
let billsBudget = [];
let monthlyBudget = [];
let budgetAutoSaveTimer = null;
let budgetAutoSaveInFlight = false;

const BUDGET_SHEET = "Budget Setup";
const BUDGET_PROJECTION_STORAGE_KEY = "fintrackBudgetProjectionAssumptions";
const BUDGET_PROJECTION_RATE_LEVELS = {
  low: { label: "Low", multiplier: 0.65 },
  medium: { label: "Med", multiplier: 1 },
  high: { label: "High", multiplier: 1.35 }
};
let budgetProjectionAssumptions = loadBudgetProjectionAssumptions();

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

function loadBudgetProjectionAssumptions() {
  try {
    const saved = JSON.parse(localStorage.getItem(BUDGET_PROJECTION_STORAGE_KEY) || "{}");
    return saved && typeof saved === "object" ? saved : {};
  } catch (err) {
    console.warn("Projection assumptions could not be loaded.", err);
    return {};
  }
}

function saveBudgetProjectionAssumptions() {
  try {
    localStorage.setItem(BUDGET_PROJECTION_STORAGE_KEY, JSON.stringify(budgetProjectionAssumptions));
  } catch (err) {
    console.warn("Projection assumptions could not be saved.", err);
  }
}

function getBudgetProjectionAssumption(category) {
  const key = budgetCategoryKey(category);
  const saved = budgetProjectionAssumptions[key] || {};
  const rate = BUDGET_PROJECTION_RATE_LEVELS[saved.rate] ? saved.rate : "medium";
  return {
    includeFuture: saved.includeFuture !== false,
    rate
  };
}

function updateProjectionInclude(encodedKey, checked) {
  const key = decodeURIComponent(encodedKey);
  const saved = budgetProjectionAssumptions[key] || {};
  budgetProjectionAssumptions[key] = {
    ...saved,
    includeFuture: Boolean(checked),
    rate: BUDGET_PROJECTION_RATE_LEVELS[saved.rate] ? saved.rate : "medium"
  };
  saveBudgetProjectionAssumptions();
  renderBudgetProjectionPanel(monthlyBudget.map(item => computeBudgetRow(item)));
}

function updateProjectionRate(encodedKey, rate) {
  if (!BUDGET_PROJECTION_RATE_LEVELS[rate]) return;
  const key = decodeURIComponent(encodedKey);
  const saved = budgetProjectionAssumptions[key] || {};
  budgetProjectionAssumptions[key] = {
    ...saved,
    includeFuture: saved.includeFuture !== false,
    rate
  };
  saveBudgetProjectionAssumptions();
  renderBudgetProjectionPanel(monthlyBudget.map(item => computeBudgetRow(item)));
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
  scheduleBudgetAutoSave();
}

function deleteBudgetItem(type, index) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  const [removed] = list.splice(index, 1);
  if (type === "Monthly Expenses" && removed) {
    delete budgetProjectionAssumptions[budgetCategoryKey(removed.category)];
    saveBudgetProjectionAssumptions();
  }
  renderBudget();
  scheduleBudgetAutoSave();
}

function updateBudgetCategory(type, index, value) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  const oldKey = budgetCategoryKey(list[index].category);
  const category = clean(value);
  list[index].category = category;
  if (type === "Monthly Expenses") {
    const newKey = budgetCategoryKey(category);
    if (oldKey && oldKey !== newKey && budgetProjectionAssumptions[oldKey]) {
      if (newKey && !budgetProjectionAssumptions[newKey]) {
        budgetProjectionAssumptions[newKey] = budgetProjectionAssumptions[oldKey];
      }
      delete budgetProjectionAssumptions[oldKey];
      saveBudgetProjectionAssumptions();
    }
  }
  renderBudget();
  scheduleBudgetAutoSave();
}

function updateBudgetAllocated(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].allocated = Number(value) || 0;
  renderBudget();
  scheduleBudgetAutoSave();
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

      return isCurrentMonth && matchesBudgetCategory(row, mainCategory, subCategory);
    })
    .reduce((sum, row) => sum + getBudgetImpactAmount(row), 0);
}

function matchesBudgetCategory(row, mainCategory, subCategory) {
  return (
    clean(row["Main Category"]).toLowerCase() === clean(mainCategory).toLowerCase() &&
    clean(row["Sub Category"]).toLowerCase()  === clean(subCategory).toLowerCase()
  );
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
  const monthlyCoveragePlan = buildMonthlyReallocationPlan(monthly, billCoverageMap);
  const monthlyCoverageMap = mergeReserveMaps(billCoverageMap, buildBillCoverageMap(monthlyCoveragePlan));
  const coveredFromMonthly = billCoveragePlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const coveredMonthlyOverspend = monthlyCoveragePlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
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
          <p>Allocated, spent, reserved, and balance by category. Amber means monthly balance is blocked for overspend elsewhere.</p>
        </div>
        ${alertHtml}
      </div>
      <div class="bv-summary-grid">
        ${renderBudgetSummaryMeter("Total Budget", total)}
        ${renderBudgetSummaryMeter("Bills", bills)}
        ${renderBudgetSummaryMeter("Monthly Expenses", monthly, coveredFromMonthly + coveredMonthlyOverspend)}
      </div>
      ${renderBudgetReserveBanner(billCoveragePlan, monthlyCoveragePlan)}
      <div class="bv-legend">
        <span><i class="spent"></i> Spent inside budget</span>
        <span><i class="reserve"></i> Reserved for overspend</span>
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
          <div class="bv-rows">${renderBudgetVisualRows(monthlyRows, monthlyCoverageMap)}</div>
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
        ? `<span>${formatCurrency(reserved)} reserved for overspend · ${formatCurrency(freeBalance)} free</span>`
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

function renderBudgetReserveBanner(billPlan, monthlyPlan) {
  const billCovered = billPlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const monthlyCovered = monthlyPlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const billNeeded = billCovered + billPlan.uncovered;
  const monthlyNeeded = monthlyCovered + monthlyPlan.uncovered;
  const covered = billCovered + monthlyCovered;
  const uncovered = billPlan.uncovered + monthlyPlan.uncovered;
  const totalNeeded = billNeeded + monthlyNeeded;
  if (totalNeeded <= 0) return "";

  const cuts = [
    ...billPlan.cuts.map(cut => ({ ...cut, reason: cut.reason || "bill overspend" })),
    ...monthlyPlan.cuts.map(cut => ({ ...cut, reason: cut.reason || "monthly overspend" }))
  ];
  const cutRows = cuts.length
    ? cuts.map(cut => `
        <div class="bv-cover-row">
          <span>${escapeHtml(cut.category)} <small>${escapeHtml(cut.reason)}</small></span>
          <strong>${formatCurrency(cut.cut)}</strong>
        </div>`).join("")
    : `<div class="bp-muted">No monthly balance available to reserve.</div>`;

  return `
    <div class="bv-cover-banner ${uncovered > 0 ? "danger" : ""}">
      <div class="bv-cover-main">
        <span>Bills overspent</span>
        <strong>${formatCurrency(billNeeded)}</strong>
      </div>
      <div class="bv-cover-main">
        <span>Monthly overspent</span>
        <strong>${formatCurrency(monthlyNeeded)}</strong>
      </div>
      <div class="bv-cover-main">
        <span>Total reserved</span>
        <strong>${formatCurrency(covered)}</strong>
      </div>
      <div class="bv-cover-list">
        ${cutRows}
        ${uncovered > 0 ? `
          <div class="bv-cover-row danger">
            <span>Still uncovered</span>
            <strong>${formatCurrency(uncovered)}</strong>
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
          <p>Expected month-end spend using usual timing or daily pace, adjusted by your category assumptions.</p>
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
    const dailyProjectedSpend = dailyAverage * daysInMonth;
    const timing = buildCategorySpendingTiming(row.type, row.category, today, dayOfMonth);
    const usesHistoricalTiming = timing.months >= 2;
    const baseProjectedSpend = usesHistoricalTiming
      ? Math.max(row.spent, row.spent + timing.medianRemaining)
      : dailyProjectedSpend;
    const baseRemaining = Math.max(0, baseProjectedSpend - row.spent);
    const assumption = getBudgetProjectionAssumption(row.category);
    const rateConfig = BUDGET_PROJECTION_RATE_LEVELS[assumption.rate] || BUDGET_PROJECTION_RATE_LEVELS.medium;
    const adjustedRemaining = assumption.includeFuture ? baseRemaining * rateConfig.multiplier : 0;
    const projectedSpend = row.spent + adjustedRemaining;
    const projectedBalance = row.allocated - projectedSpend;
    return {
      ...row,
      dailyAverage,
      dailyProjectedSpend,
      baseProjectedSpend,
      baseRemaining,
      adjustedRemaining,
      projectionAssumption: assumption,
      projectionRateLabel: rateConfig.label,
      projectionRateMultiplier: rateConfig.multiplier,
      projectedSpend,
      projectedBalance,
      projectionMethod: usesHistoricalTiming ? "history" : "daily",
      historicalMonths: timing.months,
      historicalRemaining: timing.medianRemaining,
      historicalProgressRatio: timing.medianProgressRatio
    };
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

function buildCategorySpendingTiming(mainCategory, subCategory, today, cutoffDay) {
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthsByKey = {};

  budgetTransactions.forEach(row => {
    const rowDate = parseBudgetDate(row["Date"]);
    if (!rowDate || rowDate >= currentMonthStart) return;
    if (!matchesBudgetCategory(row, mainCategory, subCategory)) return;

    const amount = getBudgetImpactAmount(row);
    if (amount <= 0) return;

    const year = rowDate.getFullYear();
    const month = rowDate.getMonth();
    const key = `${year}-${String(month + 1).padStart(2, "0")}`;
    const monthCutoffDay = Math.min(cutoffDay, new Date(year, month + 1, 0).getDate());

    if (!monthsByKey[key]) {
      monthsByKey[key] = {
        monthStart: new Date(year, month, 1),
        total: 0,
        throughCutoff: 0
      };
    }

    monthsByKey[key].total += amount;
    if (rowDate.getDate() <= monthCutoffDay) {
      monthsByKey[key].throughCutoff += amount;
    }
  });

  const months = Object.values(monthsByKey)
    .filter(month => month.total > 0)
    .sort((a, b) => b.monthStart - a.monthStart)
    .slice(0, 12);

  if (!months.length) {
    return { months: 0, medianRemaining: 0, medianProgressRatio: null };
  }

  const remaining = months.map(month => Math.max(0, month.total - month.throughCutoff));
  const progress = months.map(month => Math.min(1, Math.max(0, month.throughCutoff / month.total)));

  return {
    months: months.length,
    medianRemaining: medianNumber(remaining),
    medianProgressRatio: medianNumber(progress)
  };
}

function medianNumber(values) {
  const sorted = values
    .map(Number)
    .filter(value => Number.isFinite(value))
    .sort((a, b) => a - b);

  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[middle]
    : (sorted[middle - 1] + sorted[middle]) / 2;
}

function renderProjectionRateControls(row, encodedKey) {
  const assumption = row.projectionAssumption || getBudgetProjectionAssumption(row.category);
  const disabled = assumption.includeFuture ? "" : " disabled";
  return `
    <div class="bproj-rate-group" role="group" aria-label="Spending rate for ${escapeHtml(row.category)}">
      ${Object.entries(BUDGET_PROJECTION_RATE_LEVELS).map(([rate, config]) => `
        <button
          type="button"
          class="bproj-rate-btn${assumption.rate === rate ? " active" : ""}"
          onclick="updateProjectionRate('${encodedKey}', '${rate}')"
          ${disabled}
        >${config.label}</button>`).join("")}
    </div>`;
}

function getProjectionPaceDetail(row) {
  const assumption = row.projectionAssumption || getBudgetProjectionAssumption(row.category);
  if (!assumption.includeFuture) return "no more projected spend";

  const rateLabel = (row.projectionRateLabel || "Med").toLowerCase();
  if (row.projectionMethod === "history") {
    return `${rateLabel} usual remaining ${formatCurrency(row.adjustedRemaining)} · ${row.historicalMonths} mo`;
  }

  const adjustedDailyAverage = row.dailyAverage * (row.projectionRateMultiplier || 1);
  return `${rateLabel} pace ${formatCurrency(adjustedDailyAverage)}/day`;
}

function renderProjectionRow(row) {
  const assumption = row.projectionAssumption || getBudgetProjectionAssumption(row.category);
  const categoryKey = budgetCategoryKey(row.category);
  const encodedKey = encodeURIComponent(categoryKey);
  const balanceClass = row.projectedBalance >= 0 ? "green" : "red";
  const scale = Math.max(row.allocated, row.projectedSpend, 1);
  const fillPct = Math.min(100, (row.projectedSpend / scale) * 100);
  const budgetPct = Math.min(100, (row.allocated / scale) * 100);
  const paceDetail = getProjectionPaceDetail(row);
  const checked = assumption.includeFuture ? " checked" : "";
  return `
    <div class="bproj-row${assumption.includeFuture ? "" : " paused"}">
      <div class="bproj-row-top">
        <label class="bproj-row-toggle">
          <input type="checkbox"${checked} onchange="updateProjectionInclude('${encodedKey}', this.checked)" aria-label="Project future spend for ${escapeHtml(row.category)}">
          <span class="bproj-check-box" aria-hidden="true"></span>
          <span class="bproj-row-name">${escapeHtml(row.category)}</span>
        </label>
        <div class="bproj-row-side">
          ${renderProjectionRateControls(row, encodedKey)}
          <strong class="bproj-row-val ${balanceClass}">${formatCurrency(row.projectedBalance)}</strong>
        </div>
      </div>
      <div class="bproj-track">
        <span class="bproj-fill ${row.projectedBalance < 0 ? "over" : ""}" style="width:${fillPct.toFixed(2)}%;"></span>
        <span class="bproj-budget-line" style="left:${budgetPct.toFixed(2)}%;"></span>
      </div>
      <div class="bproj-row-detail">
        <span>${formatCurrency(row.spent)} spent · ${paceDetail}</span>
        <span>${formatCurrency(row.projectedSpend)} projected vs ${formatCurrency(row.allocated)}</span>
      </div>
    </div>`;
}

function scheduleBudgetAutoSave(delay = 700) {
  clearTimeout(budgetAutoSaveTimer);
  setBudgetAutoSaveStatus("Saving soon...");
  budgetAutoSaveTimer = setTimeout(() => {
    saveBudgetSetupToExcel({ silent: true });
  }, delay);
}

function setBudgetAutoSaveStatus(message, tone = "") {
  const el = document.getElementById("budgetAutosaveStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "autosave-status" + (tone ? " " + tone : "");
}

async function saveBudgetSetupToExcel(options = {}) {
  if (budgetAutoSaveInFlight) {
    scheduleBudgetAutoSave(1000);
    return;
  }

  const silent = options.silent === true;
  budgetAutoSaveInFlight = true;
  try {
    setBudgetAutoSaveStatus("Saving...");
    log("Saving Budget Setup to Excel...");
    await writeBudgetSetupRange("A2:B13", buildSaveValues(billsBudget));
    await writeBudgetSetupRange("F2:G13", buildSaveValues(monthlyBudget));
    setBudgetAutoSaveStatus("Saved to Excel", "ok");
    if (!silent) alert("Budget saved to Excel.");
    log("Budget saved.");
  } catch (err) {
    setBudgetAutoSaveStatus("Save failed", "error");
    log("SAVE ERROR: " + err.message);
    if (!silent) alert(err.message);
    console.error(err);
  } finally {
    budgetAutoSaveInFlight = false;
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
  const billPlan = buildBillReallocationPlan(bills, monthly);
  const billReserveMap = buildBillCoverageMap(billPlan);
  const monthlyPlan = buildMonthlyReallocationPlan(monthly, billReserveMap);
  const totalReserve = [...billPlan.cuts, ...monthlyPlan.cuts]
    .reduce((sum, cut) => sum + cut.cut, 0);
  const monthlyFreeBalance = monthly.balance - totalReserve;
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
    totalReserve > 0
      ? `Monthly expenses only; ${formatCurrency(totalReserve)} reserved for overspend.`
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
  const billCoverageMap = buildBillCoverageMap(reallocationPlan);
  const monthlyReallocationPlan = buildMonthlyReallocationPlan(monthly, billCoverageMap);
  const coveredFromMonthly = reallocationPlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const coveredMonthlyOverspend = monthlyReallocationPlan.cuts.reduce((sum, cut) => sum + cut.cut, 0);
  const billCoverageTotal = coveredFromMonthly + reallocationPlan.uncovered;
  const monthlyCoverageTotal = coveredMonthlyOverspend + monthlyReallocationPlan.uncovered;
  const totalReserved = coveredFromMonthly + coveredMonthlyOverspend;
  const totalUncovered = reallocationPlan.uncovered + monthlyReallocationPlan.uncovered;
  const reserveHtml = (billCoverageTotal + monthlyCoverageTotal) > 0
    ? `<div class="bp-reserve-strip ${totalUncovered > 0 ? "danger" : ""}">
        <div>
          <span>Bills overspent</span>
          <strong class="red">${formatCurrency(billCoverageTotal)}</strong>
        </div>
        <div>
          <span>Monthly overspent</span>
          <strong class="${monthlyCoverageTotal > 0 ? "red" : "green"}">${formatCurrency(monthlyCoverageTotal)}</strong>
        </div>
        <div>
          <span>Reserved from monthly balance</span>
          <strong>${formatCurrency(totalReserved)}</strong>
        </div>
        <div>
          <span>Still uncovered</span>
          <strong class="${totalUncovered > 0 ? "red" : "green"}">${formatCurrency(totalUncovered)}</strong>
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
    const monthlyReserveHtml = monthlyReallocationPlan.cuts.length
      ? `<div class="bp-cut-list">
          ${monthlyReallocationPlan.cuts.map(cut => `
            <div class="bp-cut-row">
              <span>${escapeHtml(cut.category)} <small>available monthly balance</small></span>
              <strong>Block ${formatCurrency(cut.cut)}</strong>
            </div>`).join("")}
          ${monthlyReallocationPlan.uncovered > 0 ? `
            <div class="bp-cut-row">
              <span>Still uncovered</span>
              <strong>${formatCurrency(monthlyReallocationPlan.uncovered)}</strong>
            </div>` : ""}
        </div>`
      : "";
    actions.push(`
      <div class="bp-action warn">
        <strong>Trim flexible spend next.</strong>
        Monthly expense categories are ${formatCurrency(monthlyCategoryOver)} over${monthlyNames ? `, led by ${escapeHtml(monthlyNames)}` : ""}. Put a short cap on those categories for the rest of the month.
        ${monthlyReserveHtml}
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
      ${reserveHtml}
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
      cuts.push({ category: row.category, cut, reason: "bill overspend" });
      remaining = Math.max(0, remaining - cut);
    });

  return { cuts, uncovered: remaining };
}

function buildMonthlyReallocationPlan(monthly, existingReserveMap = {}) {
  let remaining = monthly.rows.reduce((sum, row) => sum + row.over, 0);
  const cuts = [];
  if (remaining <= 0) return { cuts, uncovered: 0 };

  monthly.rows
    .filter(row => row.balance > 0)
    .sort((a, b) => a.spent - b.spent || b.balance - a.balance)
    .forEach(row => {
      if (remaining <= 0) return;
      const key = budgetCategoryKey(row.category);
      const alreadyReserved = Math.max(0, existingReserveMap[key] || 0);
      const availableBalance = Math.max(0, row.balance - alreadyReserved);
      const cut = Math.min(availableBalance, remaining);
      if (cut <= 0) return;
      cuts.push({ category: row.category, cut, reason: "monthly overspend" });
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

function mergeReserveMaps(...maps) {
  return maps.reduce((merged, map) => {
    Object.entries(map || {}).forEach(([key, value]) => {
      merged[key] = (merged[key] || 0) + value;
    });
    return merged;
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
