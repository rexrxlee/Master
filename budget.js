let budgetAccountTypes = [];
let budgetTransactions = [];
let billsBudget = [];
let monthlyBudget = [];
let budgetAutoSaveTimer = null;
let budgetAutoSaveInFlight = false;
let budgetSetupDirty = false;
let extraAllowanceAccounts = [];
let goalSavingsAccountsForBudget = [];

const BUDGET_SHEET = "Budget Setup";
const BUDGET_PROJECTION_STORAGE_KEY = "fintrackBudgetProjectionAssumptions";
const BUDGET_PROJECTION_RATE_LEVELS = {
  low: { label: "Low", multiplier: 0.65 },
  medium: { label: "Med", multiplier: 1 },
  high: { label: "High", multiplier: 1.35 }
};
let budgetProjectionAssumptions = loadBudgetProjectionAssumptions();

async function loadBudgetPage(forceRefresh = false) {
  try {
    clearOutput();

    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile(forceRefresh);

    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const transactionSheet = workbook.Sheets[CONFIG.sheetName];
    const budgetSheet = workbook.Sheets[BUDGET_SHEET];

    if (!transactionSheet) throw new Error("Sheet not found: " + CONFIG.sheetName);
    if (!budgetSheet)      throw new Error("Sheet not found: " + BUDGET_SHEET);

    budgetTransactions = readTransactionSheet(transactionSheet);
    billsBudget        = readBudgetSection(budgetSheet, "A2:B16", "Bills");
    monthlyBudget      = readBudgetSection(budgetSheet, "F2:G13", "Monthly Expenses");
    applyBudgetFundingMap(budgetSheet);
    try {
      const allowanceAccountsResult = await readBudgetSetupRange("AI2:AI2");
      extraAllowanceAccounts = parsePipeList(allowanceAccountsResult?.values?.[0]?.[0]);
    } catch (_) {
      extraAllowanceAccounts = parsePipeList(budgetSheet["AI2"]?.v);
    }
    goalSavingsAccountsForBudget = parsePipeList(budgetSheet["AD2"]?.v);
    accountsList       = readAccountsSection(budgetSheet, "J2:J10");
    budgetAccountTypes = XLSX.utils.sheet_to_json(budgetSheet, {header: 1, range: "J2:K10"}).map(row => ({name: row[0], type: row[1]}));

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
  const canonicalHeaders = [
    "Date",
    "Transaction",
    "Amount",
    "Main Category",
    "Sub Category",
    "Account",
    "Claimable",
    "Claim Status",
    "Claim Amount",
    "Claim Account"
  ];

  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => { if (header) obj[header] = row[index]; });
    canonicalHeaders.forEach((header, index) => {
      if (obj[header] === undefined && row[index] !== undefined) obj[header] = row[index];
    });
    return obj;
  });
}

function readBudgetSection(sheet, range, type) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, blankrows: false });
  return rows
    .map(row => ({ type, category: clean(row[0]), allocated: toNumber(row[1]) }))
    .filter(item => item.category !== "");
}

function budgetFundingKey(type, category) {
  return clean(type).toLowerCase() + "|" + clean(category).toLowerCase();
}

function applyBudgetFundingMap(sheet) {
  let map = {};
  try { map = JSON.parse(String(sheet["AE2"]?.v || "{}")); } catch (_) {}
  [...billsBudget, ...monthlyBudget].forEach(item => {
    const saved = map[budgetFundingKey(item.type, item.category)];
    if (saved && typeof saved === "object") {
      item.fundingAccount = clean(saved.fundingAccount || "");
      item.endDate = clean(saved.endDate || "");
    } else {
      // Backwards compatibility with the original account-only string format.
      item.fundingAccount = clean(saved || "");
      item.endDate = "";
    }
  });
}

function updateBudgetEndDate(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].endDate = clean(value);
  renderBudget();
  markBudgetSetupDirty();
}

function updateBudgetFundingAccount(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].fundingAccount = clean(value);
  renderBudget();
  markBudgetSetupDirty();
}

function renderBudgetFundingOptions(selected) {
  const names = (budgetAccountTypes || []).map(a => clean(a.name)).filter(Boolean);
  return ['<option value="">Select account…</option>']
    .concat(names.map(name => `<option value="${escapeHtml(name)}" ${accountKey(name) === accountKey(selected) ? "selected" : ""}>${escapeHtml(name)}</option>`))
    .join("");
}

function accountKey(name) { return clean(name).toLowerCase().replace(/\s+/g, " "); }
function fieldKey(value) { return clean(value).toLowerCase().replace(/[^a-z0-9]/g, ""); }


function addBudgetItem() {
  const type      = document.getElementById("budgetType").value;
  const category  = clean(document.getElementById("subCategoryInput").value);
  const allocated = Number(document.getElementById("allocatedInput").value);

  if (!category || isNaN(allocated)) { alert("Please enter Sub Category and Allocated."); return; }

  const list = type === "Bills" ? billsBudget : monthlyBudget;
  const maxRows = type === "Bills" ? 15 : 12;
  if (list.length >= maxRows) { alert(`Maximum ${maxRows} rows allowed for this section.`); return; }

  list.push({ type, category, allocated, fundingAccount: "", endDate: "" });
  document.getElementById("subCategoryInput").value = "";
  document.getElementById("allocatedInput").value   = "";
  renderBudget();
  markBudgetSetupDirty();
}

function deleteBudgetItem(type, index) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  const [removed] = list.splice(index, 1);
  if (type === "Monthly Expenses" && removed) {
    delete budgetProjectionAssumptions[budgetCategoryKey(removed.category)];
    saveBudgetProjectionAssumptions();
  }
  renderBudget();
  markBudgetSetupDirty();
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
  markBudgetSetupDirty();
}

function updateBudgetAllocated(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].allocated = Number(value) || 0;
  renderBudget();
  markBudgetSetupDirty();
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
  renderCashLeftPanel();
}

function renderCashLeftPanel() {
  const el = document.getElementById("cashLeftPanel");
  if (!el) return;
  const funding = getExtraAllowanceFundingPosition();
  const accountRows = funding.accounts.length
    ? funding.accounts.map(account => `
        <label class="extra-allowance-account">
          <input type="checkbox" ${account.selected ? "checked" : ""} onchange="toggleExtraAllowanceAccount('${encodeURIComponent(account.name)}', this.checked)">
          <span>${escapeHtml(account.name)}</span>
          <strong>${formatCurrency(account.balance)}</strong>
        </label>`).join("")
    : `<p class="extra-allowance-empty">Add Savings accounts in Accounts & Setup first.</p>`;
  const impactClass = funding.cashLeft > 0 ? "ok" : "danger";
  const impactTitle = funding.cashLeft > 0 ? "Cash left" : "No cash left";
  const impactText = funding.cashLeft > 0
    ? `${formatCurrency(funding.cashLeft)} remains after protecting card payments, current bills, and next month's budget.`
    : "Selected accounts are fully used after protecting card payments, current bills, and next month's budget.";
  el.innerHTML = `
    <div class="extra-allowance-card">
      <div class="extra-allowance-funding">
        <div class="extra-allowance-funding-head">
          <div>
            <strong>Cash left</strong>
            <p>Rolling cash available after card debt, next month's protected budget, and pending claims.</p>
          </div>
        </div>
        <div class="extra-allowance-account-list">${accountRows}</div>
      </div>
      <div class="extra-allowance-impact ${impactClass}">
        <strong>${impactTitle}</strong>
        <p>${impactText}</p>
        <div>
          <span>Selected account cash <b>${formatCurrency(funding.available)}</b></span>
          <span>Credit card owed <b>-${formatCurrency(funding.creditCardOwed)}</b></span>
          <span>Current bills protected <b>-${formatCurrency(funding.currentBillsProtected)}</b></span>
          <span>Next month budget block <b>-${formatCurrency(funding.nextMonthBudgetBlock)}</b></span>
          <span>Claims coming back <b>+${formatCurrency(funding.claimsComingBack)}</b></span>
          <span>Cash left <b>${formatCurrency(funding.cashLeft)}</b></span>
        </div>
      </div>
    </div>`;
}

function toggleExtraAllowanceAccount(encodedName, checked) {
  const name = decodeURIComponent(encodedName);
  if (checked) {
    if (!extraAllowanceAccounts.some(account => accountKey(account) === accountKey(name))) {
      extraAllowanceAccounts.push(name);
    }
  } else {
    extraAllowanceAccounts = extraAllowanceAccounts.filter(account => accountKey(account) !== accountKey(name));
  }
  renderBudget();
  markBudgetSetupDirty();
}

function getExtraAllowanceFundingPosition() {
  const balances = computeBudgetAccountBalances();
  const selectedKeys = new Set(extraAllowanceAccounts.map(accountKey));
  const savingsAccounts = (budgetAccountTypes || [])
    .filter(account => clean(account.name) && clean(account.type || "Savings").toLowerCase() === "savings")
    .map(account => {
      const balance = Math.max(0, balances[account.name] || 0);
      const selected = selectedKeys.has(accountKey(account.name));
      return { name: account.name, balance, selected };
    });
  const selected = savingsAccounts.filter(account => account.selected);
  const available = selected.reduce((sum, account) => sum + account.balance, 0);
  const selectedAccountNames = selected.map(account => account.name);
  const creditCardOwed = computeBudgetCreditCardOwed();
  const currentBillsProtected = computeCurrentBillsProtected(selectedAccountNames);
  const nextMonthBudgetBlock = computeNextMonthBudgetBlock(selectedAccountNames);
  const claimsComingBack = computePendingClaimsForAccounts(selectedAccountNames);
  const cashLeft = Math.max(0, available - creditCardOwed - currentBillsProtected - nextMonthBudgetBlock + claimsComingBack);
  return { accounts: savingsAccounts, selected, available, creditCardOwed, currentBillsProtected, nextMonthBudgetBlock, claimsComingBack, cashLeft };
}

function computeCurrentBillsProtected(selectedAccountNames = []) {
  return billsBudget
    .map(item => computeBudgetRow(item))
    .reduce((sum, row) => sum + Math.max(0, row.balance), 0);
}

function computeBudgetCreditCardOwed() {
  const savingsKeys = new Set((budgetAccountTypes || [])
    .filter(account => clean(account.name) && clean(account.type || "Savings").toLowerCase() === "savings")
    .map(account => accountKey(account.name)));
  const creditKeys = new Set((budgetAccountTypes || [])
    .filter(account => clean(account.name) && clean(account.type).toLowerCase() === "credit card")
    .map(account => accountKey(account.name)));
  const accountsInData = [...new Set(budgetTransactions.map(row => clean(row["Account"])).filter(Boolean))];
  const creditAccounts = creditKeys.size
    ? accountsInData.filter(account => creditKeys.has(accountKey(account)))
    : accountsInData.filter(account => !savingsKeys.has(accountKey(account)));

  return creditAccounts.reduce((total, account) => {
    const accountRows = budgetTransactions.filter(row => clean(row["Account"]) === account);
    const openingRow = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? getSignedBudgetAmount(openingRow["Amount"]) : 0;
    const openingDate = openingRow ? parseBudgetDate(openingRow["Date"]) : null;
    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseBudgetDate(row["Date"]);
      return d && d >= openingDate;
    });
    const charges = subsequent
      .filter(row => {
        const cat = clean(row["Main Category"]).toLowerCase();
        return cat !== "income" && cat !== "transfer";
      })
      .reduce((sum, row) => sum + getSignedBudgetAmount(row["Amount"]), 0);
    const transferImpact = subsequent
      .filter(row => clean(row["Main Category"]).toLowerCase() === "transfer")
      .reduce((sum, row) => {
        const sub = clean(row["Sub Category"]).toLowerCase();
        const amount = Math.abs(getSignedBudgetAmount(row["Amount"]));
        if (sub === "transfer in" || sub === "cc payment in") return sum - amount;
        if (sub === "transfer out" || sub === "cc payment out") return sum + amount;
        return sum;
      }, 0);
    const credits = subsequent
      .filter(row => clean(row["Main Category"]).toLowerCase() === "income")
      .reduce((sum, row) => sum + Math.abs(getSignedBudgetAmount(row["Amount"])), 0);
    return total + Math.max(0, openingBalance + charges + transferImpact - credits);
  }, 0);
}

function computeNextMonthBudgetBlock(selectedAccountNames = []) {
  return [...getActiveBillsForNextMonth(selectedAccountNames), ...monthlyBudget]
    .reduce((sum, item) => sum + Math.max(0, toNumber(item.allocated)), 0);
}

function getActiveBillsForNextMonth(selectedAccountNames = []) {
  const selectedKeys = new Set(selectedAccountNames.map(accountKey));
  const today = new Date();
  const nextMonthStart = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  return billsBudget.filter(item => {
    const endDate = parseBudgetDate(item.endDate);
    if (endDate && endDate < nextMonthStart) return false;
    const fundingAccount = clean(item.fundingAccount);
    return !isExternallyFundedNextMonthBill(item, fundingAccount, selectedKeys);
  });
}

function isExternallyFundedNextMonthBill(item, fundingAccount, selectedKeys) {
  if (fundingAccount && selectedKeys.has(accountKey(fundingAccount))) return false;
  const fundingKey = accountKey(fundingAccount);
  const categoryKey = accountKey(item.category);
  return fundingKey.includes("cda") ||
    fundingKey.includes("pphs") ||
    categoryKey.includes("myfirstskool") ||
    categoryKey === "pphs";
}

function isBudgetSalaryIncomeRow(row) {
  const main = clean(row["Main Category"]).toLowerCase();
  const sub = clean(row["Sub Category"]).toLowerCase();
  const transaction = clean(row["Transaction"]).toLowerCase();
  return main === "income" && (sub.includes("salary") || transaction.includes("salary"));
}

function computePendingClaimsForAccounts(accountNames = []) {
  if (!accountNames.length) return 0;
  const accountKeys = new Set(accountNames.map(accountKey));
  return budgetTransactions.reduce((sum, row) => {
    const status = clean(getBudgetRowValue(row, "Claim Status")).toLowerCase();
    if (status !== "pending") return sum;
    const claimAmount = getBudgetClaimAmount(row);
    if (claimAmount <= 0) return sum;
    const claimAccount = getBudgetClaimAccount(row);
    if (claimAccount && !accountKeys.has(accountKey(claimAccount))) return sum;
    return sum + claimAmount;
  }, 0);
}

function getBudgetRowValue(row, field) {
  const wanted = fieldKey(field);
  const key = Object.keys(row).find(k => fieldKey(k) === wanted);
  return key ? row[key] : "";
}

function getBudgetClaimAccount(row) {
  return clean(
    getBudgetRowValue(row, "Claim Account") ||
    getBudgetRowValue(row, "Claim Paid To") ||
    getBudgetRowValue(row, "Claim Paid") ||
    getBudgetRowValue(row, "Paid To")
  );
}

function getBudgetClaimAmount(row) {
  const expenseAmount = Math.abs(getSignedBudgetAmount(row["Amount"]));
  const storedClaimAmount = Math.abs(getSignedBudgetAmount(getBudgetRowValueStrict(row, "Claim Amount")));
  return Math.min(expenseAmount, storedClaimAmount > 0 ? storedClaimAmount : expenseAmount);
}

function getBudgetRowValueStrict(row, field) {
  const wanted = fieldKey(field);
  const key = Object.keys(row).find(k => fieldKey(k) === wanted);
  return key ? row[key] : "";
}

function computeBudgetAccountBalances() {
  const savingsNames = (budgetAccountTypes || [])
    .filter(account => clean(account.name) && clean(account.type || "Savings").toLowerCase() === "savings")
    .map(account => account.name);
  const balances = {};
  savingsNames.forEach(account => {
    const accountRows = budgetTransactions.filter(row => clean(row["Account"]) === account);
    const openingRow = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? getSignedBudgetAmount(openingRow["Amount"]) : 0;
    const openingDate = openingRow ? parseBudgetDate(openingRow["Date"]) : null;
    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseBudgetDate(row["Date"]);
      return d && d >= openingDate;
    });
    balances[account] = openingBalance + subsequent.reduce((sum, row) => sum + getBudgetAccountBalanceImpact(row), 0);
  });
  return balances;
}

function getBudgetAccountBalanceImpact(row) {
  const amount = Math.abs(getSignedBudgetAmount(row["Amount"]));
  const main = clean(row["Main Category"]).toLowerCase();
  const sub = clean(row["Sub Category"]).toLowerCase();
  if (main === "income") return amount;
  if (main === "transfer") {
    if (sub === "transfer in" || sub === "cc payment in") return amount;
    if (sub === "transfer out" || sub === "cc payment out") return -amount;
    return 0;
  }
  return -getSignedBudgetAmount(row["Amount"]);
}

function getSignedBudgetAmount(value) {
  if (typeof value === "number") return value;
  const n = Number(String(value).replace(/[$,]/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function parsePipeList(value) {
  return String(value ?? "").split("|").map(item => item.trim()).filter(Boolean);
}

function computeBudgetRow(item) {
  const spent = calculateSpentForCurrentMonth(item.type, item.category);
  return { ...item, spent, balance: item.allocated - spent };
}

/**
 * Calculates spending for the current month.
 * Claimable rows count for the portion that cannot be reimbursed.
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
  const isBills = type === "Bills";

  table.innerHTML = `
    <tr>
      <th>${type}</th>
      <th>Allocated</th>
      ${isBills ? `<th>Paid from</th><th>Ending date</th>` : ""}
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
      <td data-label="${type}"><input value="${escapeHtml(row.category)}" onchange="updateBudgetCategory('${type}', ${index}, this.value)"></td>
      <td data-label="Allocated"><input type="number" step="0.01" value="${row.allocated}" onchange="updateBudgetAllocated('${type}', ${index}, this.value)"></td>
      ${isBills ? `<td data-label="Paid from"><select onchange="updateBudgetFundingAccount('${type}', ${index}, this.value)">${renderBudgetFundingOptions(row.fundingAccount)}</select></td><td data-label="Ending date"><input type="date" value="${escapeHtml(row.endDate || "")}" onchange="updateBudgetEndDate('${type}', ${index}, this.value)" aria-label="Ending date for ${escapeHtml(row.category)}"></td>` : ""}
      <td data-label="Spent">${formatCurrency(row.spent)}</td>
      <td data-label="Balance" style="color:${row.balance < 0 ? '#c0392b' : 'inherit'}">${formatCurrency(row.balance)}</td>
      <td data-label=""><button type="button" onclick="deleteBudgetItem('${type}', ${index})">Delete</button></td>`;
    table.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.className = "total-row";
  totalRow.innerHTML = `
    <td><strong>Total</strong></td>
    <td><strong>${formatCurrency(totalAllocated)}</strong></td>
    ${isBills ? `<td></td><td></td>` : ""}
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
  const monthlyReservePlan = buildMonthlyReallocationPlan(monthly, {});
  const monthlyReserveMap = buildBillCoverageMap(monthlyReservePlan);
  container.innerHTML = `
    <div class="budget-visual-panel">
      <div class="bv-header">
        <div>
          <h2>Budget Visuals</h2>
          <p>Allocated, spent, balance left, and overspend by category.</p>
        </div>
      </div>
      <div class="bv-summary-grid">
        ${renderBudgetSummaryMeter("Total Budget", total)}
        ${renderBudgetSummaryMeter("Bills", bills)}
        ${renderBudgetSummaryMeter("Monthly Expenses", monthly)}
      </div>
      ${renderBudgetOverspendBanner(bills, monthly)}
      <div class="bv-legend">
        <span><i class="spent"></i> Spent inside budget</span>
        <span><i class="reserve"></i> Reserved</span>
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
          <div class="bv-rows">
            ${renderBudgetVisualRows(monthlyRows, monthlyReserveMap)}
          </div>
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
      const displayBalance = reserved > 0 ? freeBalance : row.balance;
      const balanceClass = displayBalance < 0 ? "red" : "green";
      const rowClass = row.balance < 0 ? "bv-row over" : "bv-row";
      const reserveDetail = reserved > 0
        ? `<span>${formatCurrency(reserved)} reserved · ${formatCurrency(freeBalance)} balance left</span>`
        : `<span>${formatCurrency(row.allocated)} allocated</span>`;
      return `
        <div class="${rowClass}">
          <div class="bv-row-top">
            <span class="bv-row-name">${escapeHtml(row.category)}</span>
            <strong class="bv-row-balance ${balanceClass}">${formatCurrency(displayBalance)}</strong>
          </div>
          ${renderBudgetMeter(row, reserved)}
          <div class="bv-row-detail">
            <span>${formatCurrency(row.spent)} spent</span>
            ${reserveDetail}
          </div>
        </div>`;
    }).join("");
}

function renderBudgetMeter(row, reserved = 0, options = {}) {
  const spentInside = Math.min(Math.max(0, row.spent), Math.max(0, row.allocated));
  const balanceLeft = Math.max(0, row.allocated - row.spent);
  const reserve = options.reserveBeyondBalance
    ? Math.max(0, reserved)
    : Math.min(Math.max(0, reserved), balanceLeft);
  const freeBalance = Math.max(0, balanceLeft - reserve);
  const overspent = options.hideOverWhenReserved ? 0 : Math.max(0, row.spent - row.allocated);
  const scale = Math.max(spentInside + reserve + freeBalance + overspent, row.allocated, row.spent, 1);
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

function renderBudgetOverspendBanner(bills, monthly) {
  const monthlyOver = monthly.rows.reduce((sum, row) => sum + row.over, 0);
  if (monthlyOver <= 0) return "";

  const overRows = monthly.rows
    .filter(row => row.over > 0)
    .slice(0, 5)
    .map(row => `
        <div class="bv-cover-row">
          <span>${escapeHtml(row.category)}</span>
          <strong>${formatCurrency(row.over)}</strong>
        </div>`)
    .join("");
  return `
    <div class="bv-cover-banner danger">
      <div class="bv-cover-main">
        <span>Monthly overspent from allocated budget</span>
        <strong>${formatCurrency(monthlyOver)}</strong>
      </div>
      <div class="bv-cover-list">
        ${overRows}
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
    if (amount === 0) return;

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
    return `${formatCurrency(row.adjustedRemaining)} more expected based on past ${row.historicalMonths} months (${rateLabel})`;
  }

  const adjustedDailyAverage = row.dailyAverage * (row.projectionRateMultiplier || 1);
  return `${formatCurrency(adjustedDailyAverage)}/day expected for rest of month (${rateLabel})`;
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
        <span>${formatCurrency(row.spent)} spent - ${paceDetail}</span>
        <span>${formatCurrency(row.projectedSpend)} projected vs ${formatCurrency(row.allocated)}</span>
      </div>
    </div>`;
}

function scheduleBudgetAutoSave(delay = 700) {
  markBudgetSetupDirty();
}

function markBudgetSetupDirty() {
  budgetSetupDirty = true;
  clearTimeout(budgetAutoSaveTimer);
  setBudgetAutoSaveStatus("Unsaved changes", "warn");
}

function setBudgetAutoSaveStatus(message, tone = "") {
  const el = document.getElementById("budgetAutosaveStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "autosave-status" + (tone ? " " + tone : "");
}

async function saveBudgetSetupToExcel(options = {}) {
  if (budgetAutoSaveInFlight) {
    setBudgetAutoSaveStatus("Save already running...", "warn");
    return;
  }

  const silent = options.silent === true;
  budgetAutoSaveInFlight = true;
  try {
    setBudgetAutoSaveStatus("Saving...");
    log("Saving Budget Setup to Excel...");
    await writeBudgetSetupRange("A2:B16", buildSaveValues(billsBudget, 15));
    await writeBudgetSetupRange("F2:G13", buildSaveValues(monthlyBudget, 12));
    await writeBudgetSetupRange("AE2:AE2", [[JSON.stringify(buildBudgetFundingMap())]]);
    await writeBudgetSetupRange("AI2:AI2", [[extraAllowanceAccounts.join("|")]]);
    budgetSetupDirty = false;
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

function buildBudgetFundingMap() {
  const map = {};
  billsBudget.forEach(item => {
    if (clean(item.category) && (clean(item.fundingAccount) || clean(item.endDate))) {
      map[budgetFundingKey(item.type, item.category)] = {
        fundingAccount: clean(item.fundingAccount),
        endDate: clean(item.endDate || "")
      };
    }
  });
  return map;
}

function buildSaveValues(list, rows = 12) {
  const values = list.map(item => [item.category, item.allocated]);
  while (values.length < rows) values.push(["", ""]);
  return values;
}

function updateBudgetCards(billsRows, monthlyRows) {
  const rows = [...billsRows, ...monthlyRows];
  const bills = summariseBudgetRows(billsRows);
  const monthly = summariseBudgetRows(monthlyRows);
  const monthlyFreeBalance = monthly.balance;
  const totalAllocated = rows.reduce((s, r) => s + r.allocated, 0);
  const totalSpent     = rows.reduce((s, r) => s + r.spent,     0);
  const totalBalance   = rows.reduce((s, r) => s + r.balance,   0);

  const foodRow     = monthlyRows.find(r => clean(r.category).toLowerCase() === "food");
  const foodBalance = foodRow ? foodRow.balance : 0;
  const daysLeft    = getDaysRemainingInMonth();
  const monthlyIsOverspent = monthly.balance < 0;
  const foodPerDay = monthlyIsOverspent ? 0 : foodBalance / daysLeft;
  const monthlyPerDay = monthlyIsOverspent ? 0 : monthlyFreeBalance / daysLeft;

  setCurrencyValue("totalAllocated", totalAllocated);
  setCurrencyValue("totalSpent", totalSpent, "red");
  setCurrencyValue("totalBalance", totalBalance, totalBalance < 0 ? "red" : "green");
  setCurrencyValue("foodPerDay", foodPerDay, foodPerDay <= 0 ? "red" : "green");
  setCurrencyValue("monthlyPerDay", monthlyPerDay, monthlyPerDay <= 0 ? "red" : "green");
  setTextValue(
    "monthlyPerDayNote",
    monthlyIsOverspent
      ? "Monthly expenses are overspent; any more spending eats into cash left."
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
  const overMonthly = monthly.rows.filter(row => row.over > 0).slice(0, 4);
  const billsCategoryOver = bills.rows.reduce((sum, row) => sum + row.over, 0);
  const monthlyCategoryOver = monthly.rows.reduce((sum, row) => sum + row.over, 0);
  const monthlyNames = overMonthly.map(row => row.category).slice(0, 3).join(", ");
  const actions = [];

  if (billsCategoryOver > 0) {
    const billAction = bills.balance < 0
      ? "Top up the bill budget or reduce goal allocations before cutting required payments."
      : "Move allocation from under-used bill lines, or top up the bill budget if this is a permanent increase.";
    actions.push(`
      <div class="bp-action danger">
        <strong>Protect bills first.</strong>
        Bill categories are ${formatCurrency(billsCategoryOver)} over. These are fixed commitments. ${billAction}
      </div>`);
  } else {
    actions.push(`
      <div class="bp-action ok">
        <strong>Bills are covered.</strong>
        Bills still have ${formatCurrency(Math.max(0, bills.balance))} left. Keep that available before sending extra money to goals.
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
        <span>What to do next</span>
      </div>
      <div class="bp-actions compact">
        ${actions.join("")}
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
  return isNaN(n) ? 0 : n;
}

function getBudgetImpactAmount(row) {
  if (isBusinessTransaction(row, budgetAccountTypes)) return 0;
  const amount = getAmount(row["Amount"]);
  const claimableKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claimable") || "Claimable";
  const claimable = clean(row[claimableKey]).toLowerCase();
  if (claimable !== "yes") return amount;

  const claimAmountKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claim amount") || "Claim Amount";
  const storedClaimAmount = getAmount(row[claimAmountKey]);
  const claimAmount = Math.min(amount, storedClaimAmount > 0 ? storedClaimAmount : amount);
  return Math.max(0, amount - claimAmount);
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
