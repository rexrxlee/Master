let allTransactions = [];
let isUpdatingFilters = false;
let savingsAccountNames = [];
let creditCardAccountNames = [];
let configuredAccounts = [];
let monthlyExpenseBudgetRows = [];
let assetAccountSelections = {};
let activeDashboardView = "expenses";
const ASSET_SCOPE_STORAGE_KEY = "fintrack.dashboard.assetAccountSelections.v1";
const NO_FILTER_SELECTION = "__FINTRACK_NONE_SELECTED__";

async function loadDashboard() {
  try {
    clearOutput();
    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();
    log("Reading workbook...");
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const worksheet = workbook.Sheets[CONFIG.sheetName];
    if (!worksheet) throw new Error("Sheet not found: " + CONFIG.sheetName);
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false });
    prepareTransactions(rows);
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (budgetSheet) {
      monthlyExpenseBudgetRows = readDashboardBudgetSection(budgetSheet, "F2:G13");
      const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, blankrows: false });
      configuredAccounts = allRows.slice(1)
        .map(row => ({
          name: clean(row[9]),
          type: clean(row[10] || "Savings")
        }))
        .filter(account => account.name !== "");

      savingsAccountNames = configuredAccounts
        .filter(account => account.type.toLowerCase() === "savings")
        .map(account => account.name);

      creditCardAccountNames = configuredAccounts
        .filter(account => account.type.toLowerCase() === "credit card")
        .map(account => account.name);
    } else {
      monthlyExpenseBudgetRows = [];
    }
    loadAssetAccountSelections();
    setupFilters();
    const filters = getCurrentFilters();
    refreshAllFilters(filters);
    updateDashboard(filters);
    updateActiveFilterBadges(filters);
    log("Dashboard loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function clean(value) { return String(value ?? "").trim(); }
function accountKey(value) { return clean(value).toLowerCase().replace(/\s+/g, " "); }
function fieldKey(value) { return clean(value).toLowerCase().replace(/[^a-z0-9]/g, ""); }

function readDashboardBudgetSection(sheet, range) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, blankrows: false });
  return rows
    .map(row => ({
      category: clean(row[0]),
      allocated: Number(String(row[1] ?? "").replace(/[$,]/g, "")) || 0
    }))
    .filter(row => row.category !== "");
}
function getRowValue(row, fieldName) {
  const wanted = fieldKey(fieldName);
  const key = Object.keys(row || {}).find(k => fieldKey(k) === wanted);
  return key ? row[key] : undefined;
}

function prepareTransactions(rows) {
  const headers = rows[0].map(header => clean(header));
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
  allTransactions = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = row[index];
    });
    canonicalHeaders.forEach((header, index) => {
      if (obj[header] === undefined && row[index] !== undefined) obj[header] = row[index];
    });
    return obj;
  });
}

function loadAssetAccountSelections() {
  try {
    assetAccountSelections = JSON.parse(localStorage.getItem(ASSET_SCOPE_STORAGE_KEY) || "{}") || {};
  } catch {
    assetAccountSelections = {};
  }
}

function saveAssetAccountSelections() {
  try {
    localStorage.setItem(ASSET_SCOPE_STORAGE_KEY, JSON.stringify(assetAccountSelections));
  } catch {
    // Local storage is a convenience only; dashboard still works without it.
  }
}

function isAssetAccountIncluded(account) {
  const key = accountKey(account);
  return assetAccountSelections[key] !== false;
}

function toggleAssetAccount(account, checked) {
  assetAccountSelections[accountKey(account)] = !!checked;
  saveAssetAccountSelections();
  updateFinanceCards(allTransactions);
}

function setupFilters() {
  document.getElementById("fromDate").addEventListener("change", handleFilterChange);
  document.getElementById("toDate").addEventListener("change", handleFilterChange);

  // Auto-collapse open dropdown when another is opened, or click outside
  document.addEventListener("click", function(event) {
    if (!event.target.closest(".filter-box") && !event.target.closest(".date-range-panel")) {
      document.querySelectorAll(".checkbox-menu").forEach(menu => { menu.style.display = "none"; });
      const drp = document.getElementById("dateRangePanel");
      if (drp) drp.style.display = "none";
    }
  });
}

function switchDashboardView(view) {
  if (!["expenses", "income"].includes(view)) return;
  if (activeDashboardView !== view) {
    const current = getCurrentFilters();
    activeDashboardView = view;
    const resetNonDateFilters = {
      mainCategories: [],
      subCategories: [],
      transactions: [],
      fromDate: current.fromDate,
      toDate: current.toDate
    };
    refreshAllFilters(resetNonDateFilters);
  } else {
    syncDashboardViewVisibility();
  }

  const filters = getCurrentFilters();
  updateDashboard(filters);
  updateActiveFilterBadges(filters);
}

function syncDashboardViewVisibility() {
  const isIncome = activeDashboardView === "income";
  const expensesView = document.getElementById("expensesDashboardView");
  const incomeView = document.getElementById("incomeDashboardView");
  const expensesTab = document.getElementById("expensesDashboardTab");
  const incomeTab = document.getElementById("incomeDashboardTab");

  if (expensesView) expensesView.classList.toggle("active", !isIncome);
  if (incomeView) incomeView.classList.toggle("active", isIncome);
  if (expensesTab) expensesTab.classList.toggle("active", !isIncome);
  if (incomeTab) incomeTab.classList.toggle("active", isIncome);
}

function toggleDropdown(menuId) {
  const menu = document.getElementById(menuId);
  const isOpen = menu.style.display === "block";
  // Close all menus first
  document.querySelectorAll(".checkbox-menu").forEach(m => { m.style.display = "none"; });
  const drp = document.getElementById("dateRangePanel");
  if (drp) drp.style.display = "none";
  // Then open this one if it was closed
  if (!isOpen) menu.style.display = "block";
}

function toggleDatePanel() {
  const panel = document.getElementById("dateRangePanel");
  const isOpen = panel.style.display === "block";
  document.querySelectorAll(".checkbox-menu").forEach(m => { m.style.display = "none"; });
  panel.style.display = isOpen ? "none" : "block";
}

function applyDatePreset(preset) {
  const today = new Date();
  let from, to;
  if (preset === "today") {
    from = to = today;
  } else if (preset === "this_week") {
    const day = today.getDay();
    from = new Date(today); from.setDate(today.getDate() - (day === 0 ? 6 : day - 1));
    to = today;
  } else if (preset === "this_month") {
    from = new Date(today.getFullYear(), today.getMonth(), 1);
    to = today;
  } else if (preset === "last_month") {
    from = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    to = new Date(today.getFullYear(), today.getMonth(), 0);
  } else if (preset === "this_year") {
    from = new Date(today.getFullYear(), 0, 1);
    to = today;
  } else if (preset === "last_3m") {
    from = new Date(today); from.setMonth(from.getMonth() - 3);
    to = today;
  }
  document.getElementById("fromDate").value = toDateInputValue(from);
  document.getElementById("toDate").value = toDateInputValue(to);
  handleFilterChange();
}

function toDateInputValue(d) {
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

function getCheckedValues(menuId) {
  const checkboxes = Array.from(document.querySelectorAll(`#${menuId} input[type="checkbox"]:not([data-select-all])`));
  const checked = checkboxes.filter(input => input.checked).map(input => clean(input.value));
  if (checkboxes.length > 0 && checked.length === 0) return [NO_FILTER_SELECTION];
  return checkboxes.length > 0 && checked.length === checkboxes.length ? [] : checked;
}

function hasNoFilterSelection(values) {
  return Array.isArray(values) && values.includes(NO_FILTER_SELECTION);
}

function cleanFilterValues(values) {
  return (values || []).filter(value => value !== NO_FILTER_SELECTION);
}

function filterAllowsValue(values, rowValue) {
  if (hasNoFilterSelection(values)) return false;
  const selected = cleanFilterValues(values);
  return selected.length === 0 || selected.includes(clean(rowValue));
}

function getCurrentFilters() {
  return {
    mainCategories: getCheckedValues("mainCategoryMenu"),
    subCategories: getCheckedValues("subCategoryMenu"),
    transactions: getCheckedValues("transactionMenu"),
    fromDate: document.getElementById("fromDate").value,
    toDate: document.getElementById("toDate").value
  };
}

function handleFilterChange() {
  if (isUpdatingFilters) return;
  const filters = getCurrentFilters();
  updateDashboard(filters);
  updateActiveFilterBadges(filters);
  refreshAllFilters(filters);
}

function isExpenseRow(row) {
  const cat = clean(row["Main Category"]).toLowerCase();
  return cat !== "income" && cat !== "transfer" && cat !== "saving goals";
}

function isClaimableRow(row) {
  return clean(getRowValue(row, "Claimable")).toLowerCase() === "yes";
}

function isSpendAnalyticsRow(row) {
  return isExpenseRow(row) && !isClaimableRow(row);
}

function getClaimAmount(row) {
  const expenseAmount = getAmount(getRowValue(row, "Amount"));
  const storedClaimAmount = getAmount(getRowValue(row, "Claim Amount"));
  const claimAmount = storedClaimAmount > 0 ? storedClaimAmount : expenseAmount;
  return Math.min(expenseAmount, claimAmount);
}

function getPendingClaimReceivableAmount(row) {
  if (!isClaimableRow(row)) return 0;
  if (clean(getRowValue(row, "Claim Status")).toLowerCase() !== "pending") return 0;
  return getClaimAmount(row);
}

function getClaimReceivableAccount(row) {
  return clean(getRowValue(row, "Claim Account")) || "(Claim account not set)";
}

function computePendingClaimsByAccount(rows, shouldIncludeAccount = () => true) {
  const byAccount = {};
  let total = 0;
  let count = 0;

  (rows || []).forEach(row => {
    const amount = getPendingClaimReceivableAmount(row);
    if (amount <= 0) return;

    const account = getClaimReceivableAccount(row);
    if (!shouldIncludeAccount(account)) return;

    byAccount[account] = (byAccount[account] || 0) + amount;
    total += amount;
    count += 1;
  });

  return { byAccount, total, count };
}

function isIncomeRow(row) {
  return clean(row["Main Category"]).toLowerCase() === "income"
    && clean(row["Sub Category"]) !== "Opening Balance";
}

function rowMatchesFilters(row, filters, ignoreField = null) {
  const rowDate = parseExcelDate(row["Date"]);
  if (!rowDate) return false;
  if (ignoreField !== "mainCategory" && !filterAllowsValue(filters.mainCategories, row["Main Category"])) return false;
  if (ignoreField !== "subCategory" && !filterAllowsValue(filters.subCategories, row["Sub Category"])) return false;
  if (ignoreField !== "transaction" && !filterAllowsValue(filters.transactions, row["Transaction"])) return false;
  if (ignoreField !== "date" && filters.fromDate) {
    const [fy, fm, fd] = filters.fromDate.split("-").map(Number);
    if (rowDate < new Date(fy, fm - 1, fd)) return false;
  }
  if (ignoreField !== "date" && filters.toDate) {
    const [ty, tm, td] = filters.toDate.split("-").map(Number);
    if (rowDate > new Date(ty, tm - 1, td, 23, 59, 59)) return false;
  }
  return true;
}

function rowMatchesDateFilters(row, filters) {
  return rowMatchesFilters(row, {
    mainCategories: [],
    subCategories: [],
    transactions: [],
    fromDate: filters.fromDate,
    toDate: filters.toDate
  });
}

function refreshAllFilters(filters = getCurrentFilters()) {
  isUpdatingFilters = true;
  const filterRows = activeDashboardView === "income"
    ? allTransactions.filter(isIncomeRow)
    : allTransactions.filter(isSpendAnalyticsRow);
  rebuildCheckboxMenu("mainCategoryMenu", "Main Category", filterRows.filter(row => rowMatchesFilters(row, filters, "mainCategory")), filters.mainCategories);
  rebuildCheckboxMenu("subCategoryMenu", "Sub Category", filterRows.filter(row => rowMatchesFilters(row, filters, "subCategory")), filters.subCategories);
  rebuildCheckboxMenu("transactionMenu", "Transaction", filterRows.filter(row => rowMatchesFilters(row, filters, "transaction")), filters.transactions);
  isUpdatingFilters = false;
}

function rebuildCheckboxMenu(menuId, columnName, rows, selectedValues) {
  const menu = document.getElementById(menuId);
  const noneSelected = hasNoFilterSelection(selectedValues);
  const selected = cleanFilterValues(selectedValues);
  const values = [...new Set([
    ...rows.map(row => clean(row[columnName])).filter(v => v !== ""),
    ...selected
  ])].sort();
  menu.innerHTML = "";

  const allSelected = !noneSelected && (selected.length === 0 || values.every(value => selected.includes(value)));
  const allLabel = document.createElement("label");
  allLabel.className = "checkbox-option select-all-option";
  const allCheckbox = document.createElement("input");
  allCheckbox.type = "checkbox";
  allCheckbox.dataset.selectAll = "true";
  allCheckbox.checked = allSelected;
  allCheckbox.indeterminate = !noneSelected && selected.length > 0 && !allSelected;
  allCheckbox.addEventListener("change", () => {
    menu.querySelectorAll("input[type='checkbox']:not([data-select-all])").forEach(input => {
      input.checked = allCheckbox.checked;
    });
    allCheckbox.indeterminate = false;
    handleFilterChange();
  });
  allLabel.appendChild(allCheckbox);
  allLabel.appendChild(document.createTextNode(" Select all"));
  menu.appendChild(allLabel);

  values.forEach(value => {
    const label = document.createElement("label");
    label.className = "checkbox-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = value;
    checkbox.checked = !noneSelected && (selected.length === 0 || selected.includes(value));
    checkbox.addEventListener("change", () => {
      const itemCheckboxes = Array.from(menu.querySelectorAll("input[type='checkbox']:not([data-select-all])"));
      const checkedCount = itemCheckboxes.filter(input => input.checked).length;
      allCheckbox.checked = itemCheckboxes.length > 0 && checkedCount === itemCheckboxes.length;
      allCheckbox.indeterminate = checkedCount > 0 && checkedCount < itemCheckboxes.length;
      handleFilterChange();
    });
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(" " + value));
    menu.appendChild(label);
  });
}

function updateActiveFilterBadges(filters = getCurrentFilters()) {
  const badges = [];
  addFilterBadges(badges, "Main", filters.mainCategories);
  addFilterBadges(badges, "Sub", filters.subCategories);
  addFilterBadges(badges, "Tx", filters.transactions);
  if (filters.fromDate) badges.push(`From: ${filters.fromDate}`);
  if (filters.toDate) badges.push(`To: ${filters.toDate}`);

  const container = document.getElementById("activeFilterBadges");
  if (!container) return;
  container.innerHTML = "";
  badges.forEach(badge => {
    const span = document.createElement("span");
    span.className = "filter-badge";
    span.textContent = badge;
    container.appendChild(span);
  });
  const clearBtn = document.getElementById("clearFiltersBtn");
  if (clearBtn) clearBtn.style.display = badges.length > 0 ? "inline-block" : "none";
}

function addFilterBadges(badges, label, values) {
  if (!values || values.length === 0) return;
  if (hasNoFilterSelection(values)) {
    badges.push(`${label}: none selected`);
    return;
  }
  const selected = cleanFilterValues(values);
  if (selected.length > 3) {
    badges.push(`${label}: ${selected.length} selected`);
  } else {
    badges.push(...selected.map(v => `${label}: ${v}`));
  }
}

function updateDashboard(filters = getCurrentFilters()) {
  syncDashboardViewVisibility();

  const filtered = allTransactions.filter(row => isSpendAnalyticsRow(row) && rowMatchesFilters(row, filters));
  const incomeByDate = allTransactions.filter(row => isIncomeRow(row) && rowMatchesDateFilters(row, filters));
  const incomeFiltered = allTransactions.filter(row => {
    if (!isIncomeRow(row)) return false;
    return activeDashboardView === "income"
      ? rowMatchesFilters(row, filters)
      : rowMatchesDateFilters(row, filters);
  });

  const totalExpenses = filtered.reduce((sum, row) => sum + getAmount(row["Amount"]), 0);
  const totalIncome = incomeByDate.reduce((sum, row) => sum + getAmount(row["Amount"]), 0);
  const netSavings = totalIncome - totalExpenses;

  document.getElementById("totalExpenses").innerText = formatKpiCurrency(totalExpenses);
  const incomeEl = document.getElementById("totalIncome");
  if (incomeEl) incomeEl.innerText = formatKpiCurrency(totalIncome);
  const netEl = document.getElementById("netSavings");
  if (netEl) { netEl.innerText = formatKpiCurrency(netSavings); netEl.style.color = netSavings >= 0 ? "#27ae60" : "#c0392b"; }
  const netNote = document.getElementById("netSavingsNote");
  if (netNote) {
    const filteredExpenseView = filters.mainCategories.length || filters.subCategories.length || filters.transactions.length;
    netNote.textContent = filteredExpenseView
      ? "All income in date range minus filtered expenses"
      : "Income minus expenses for the selected date range";
  }
  document.getElementById("transactionCount").innerText = filtered.length;

  updateFinanceCards(allTransactions);
  renderMonthlyExpenseBudgetInsight(computeMonthlyExpenseBudgetPosition(allTransactions));
  if (activeDashboardView === "income") {
    drawIncomeSubCategoryChart(incomeFiltered, filters);
  } else {
    drawMonthlyExpenseChart(filtered, incomeByDate, filters);
    drawSubCategoryMonthlyChart(filtered, filters);
  }
  renderRecentTransactions(filtered);
  renderTop5Transactions(filtered);
  renderMonthlyAvgByCategory(filtered);
  renderIncomeBreakdown(incomeFiltered, filters);
  renderIncomeSummaryCards(incomeFiltered, filters);
  renderIncomeInsightCards(incomeFiltered, allTransactions, filters);
  renderIncomeMonthlyTable(incomeFiltered);
  renderInsightCards(filtered, incomeByDate, allTransactions);
}

// ─── Monthly Expenses Budget Insight ──────────────────────────────────────────

function computeMonthlyExpenseBudgetPosition(data) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  const daysElapsed = Math.max(1, today.getDate());
  const daysLeft = Math.max(1, daysInMonth - today.getDate() + 1);

  const allocatedByCategory = {};
  monthlyExpenseBudgetRows.forEach(row => {
    const category = clean(row.category);
    if (!category) return;
    allocatedByCategory[category] = (allocatedByCategory[category] || 0) + row.allocated;
  });

  const spentByCategory = {};
  (data || []).forEach(row => {
    const rowDate = parseExcelDate(row["Date"]);
    if (!rowDate) return;
    if (rowDate.getFullYear() !== currentYear || rowDate.getMonth() !== currentMonth) return;
    if (clean(row["Main Category"]).toLowerCase() !== "monthly expenses") return;
    if (isClaimableRow(row)) return;

    const category = clean(row["Sub Category"]) || "Uncategorised";
    spentByCategory[category] = (spentByCategory[category] || 0) + getAmount(row["Amount"]);
  });

  const categories = [...new Set([
    ...Object.keys(allocatedByCategory),
    ...Object.keys(spentByCategory)
  ])];

  const rows = categories.map(category => {
    const allocated = allocatedByCategory[category] || 0;
    const spent = spentByCategory[category] || 0;
    const balance = allocated - spent;
    const projectedSpend = spent / daysElapsed * daysInMonth;
    const projectedBalance = allocated - projectedSpend;
    return {
      category,
      allocated,
      spent,
      balance,
      leftPerDay: balance / daysLeft,
      projectedSpend,
      projectedBalance,
      status: getMonthlyBudgetStatus(allocated, spent, balance, projectedBalance)
    };
  }).sort((a, b) => {
    if (a.balance < 0 && b.balance >= 0) return -1;
    if (b.balance < 0 && a.balance >= 0) return 1;
    return a.balance - b.balance;
  });

  const allocated = rows.reduce((sum, row) => sum + row.allocated, 0);
  const spent = rows.reduce((sum, row) => sum + row.spent, 0);
  const balance = allocated - spent;
  const projectedSpend = spent / daysElapsed * daysInMonth;
  const projectedBalance = allocated - projectedSpend;

  return {
    rows,
    allocated,
    spent,
    balance,
    leftPerDay: balance / daysLeft,
    projectedSpend,
    projectedBalance,
    monthLabel: today.toLocaleString("en-SG", { month: "short", year: "2-digit" })
  };
}

function getMonthlyBudgetStatus(allocated, spent, balance, projectedBalance) {
  if (allocated <= 0 && spent > 0) return { label: "Unbudgeted", className: "over" };
  if (balance < 0) return { label: "Over", className: "over" };
  if (projectedBalance < 0) return { label: "Watch", className: "watch" };
  return { label: "On track", className: "ok" };
}

function renderMonthlyExpenseBudgetInsight(summary) {
  setMoneyText("monthlyExpenseBalance", summary.balance, true);
  setMoneyText("monthlyExpenseSpent", summary.spent, false);
  setMoneyText("monthlyExpenseLeftPerDay", summary.leftPerDay, true);
  setMoneyText("monthlyExpenseProjected", summary.projectedBalance, true);

  const projectedNote = document.getElementById("monthlyExpenseProjectedNote");
  if (projectedNote) {
    projectedNote.textContent = `Projected spend ${formatCurrency(summary.projectedSpend)} this month`;
  }

  const badge = document.getElementById("monthlyExpenseBudgetBadge");
  if (badge) {
    badge.textContent = summary.balance >= 0
      ? `${summary.monthLabel}: ${formatCurrency(summary.balance)} left`
      : `${summary.monthLabel}: ${formatCurrency(Math.abs(summary.balance))} over`;
    badge.classList.toggle("warn", summary.balance < 0 || summary.projectedBalance < 0);
  }

  const table = document.getElementById("monthlyExpenseBudgetTable");
  if (!table) return;

  table.innerHTML = `<tr><th>Sub Category</th><th>Allocated</th><th>Spent</th><th>Balance</th><th>Left / Day</th><th>Status</th></tr>`;
  if (!summary.rows.length) {
    table.innerHTML += `<tr><td colspan="6" style="text-align:center;color:#999;">No monthly expense budget rows found</td></tr>`;
    return;
  }

  summary.rows.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.category}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;">${formatCurrency(row.allocated)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;">${formatCurrency(row.spent)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;color:${row.balance >= 0 ? "#16a34a" : "#dc2626"};font-weight:bold;">${formatCurrency(row.balance)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;color:${row.leftPerDay >= 0 ? "#16a34a" : "#dc2626"};">${formatCurrency(row.leftPerDay)}</td>
      <td><span class="budget-status ${row.status.className}">${row.status.label}</span></td>
    `;
    table.appendChild(tr);
  });
}

function setMoneyText(elementId, value, positiveIsGood = true, suffix = "") {
  const element = document.getElementById(elementId);
  if (!element) return;
  element.textContent = formatKpiCurrency(value) + suffix;
  element.style.color = value >= 0
    ? (positiveIsGood ? "#16a34a" : "#1a1d23")
    : "#dc2626";
}

// ─── Finance Cards ─────────────────────────────────────────────────────────────

function updateFinanceCards(data) {
  const savingsAccounts = savingsAccountNames;
  const balances = {};
  const savingsAccountKeys = new Set(savingsAccounts.map(accountKey));

  savingsAccounts.forEach(account => {
    const accountRows = data.filter(row => clean(row["Account"]) === account);

    const openingRow = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? getSignedAmount(openingRow["Amount"]) : 0;
    const openingDate = openingRow ? parseExcelDate(openingRow["Date"]) : null;

    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseExcelDate(row["Date"]);
      return d && d >= openingDate;
    });

    const movement = subsequent.reduce((sum, row) => sum + getAccountBalanceImpact(row), 0);

    balances[account] = openingBalance + movement;
  });

  const includedSavingsAccounts = savingsAccounts.filter(isAssetAccountIncluded);
  const savingsTotal = includedSavingsAccounts.reduce((sum, account) => sum + (balances[account] || 0), 0);
  const excludedSavingsTotal = savingsAccounts
    .filter(account => !isAssetAccountIncluded(account))
    .reduce((sum, account) => sum + (balances[account] || 0), 0);

  // ── Credit card balance ────────────────────────────────────────────────────
  // Outstanding CC balance = opening balance + charges, adjusted by transfer direction.
  // Opening balance row is stored as Income/Opening Balance with a positive amount (what you owe).
  // Expense rows on the CC account add to what you owe.
  // Transfer/CC payment rows into the CC reduce what you owe; transfers out increase it.
  // We know the CC accounts from Budget Setup (type = "Credit Card").
  // Fall back to the old "not savings" inference only if no credit-card accounts are configured.
  const allAccountsInData = [...new Set(data.map(row => clean(row["Account"])).filter(Boolean))];
  const creditCardAccountKeys = new Set(creditCardAccountNames.map(accountKey));
  const ccAccounts = creditCardAccountNames.length > 0
    ? allAccountsInData.filter(a => creditCardAccountKeys.has(accountKey(a)))
    : allAccountsInData.filter(a => !savingsAccountKeys.has(accountKey(a)));

  const ccBalances = {};
  ccAccounts.forEach(account => {
    const accountRows = data.filter(row => clean(row["Account"]) === account);

    // Opening balance row (stored as Income > Opening Balance)
    const openingRow = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? Math.abs(getSignedAmount(openingRow["Amount"])) : 0;
    const openingDate   = openingRow ? parseExcelDate(openingRow["Date"]) : null;

    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseExcelDate(row["Date"]);
      return d && d >= openingDate;
    });

    // Charges: expense rows billed to this CC (add to balance owed).
    // Claimable/claimed rows still hit the credit-card bill; the claim fields only
    // adjust budget impact, not what is owed to the card issuer.
    const charges = subsequent
      .filter(row => {
        const cat = clean(row["Main Category"]).toLowerCase();
        return cat !== "income" && cat !== "transfer";
      })
      .reduce((sum, row) => sum + Math.abs(getSignedAmount(row["Amount"])), 0);

    // Transfers into a CC reduce what is owed; transfers out of a CC increase it.
    const transferImpact = subsequent
      .filter(row => clean(row["Main Category"]).toLowerCase() === "transfer")
      .reduce((sum, row) => sum + getCreditCardTransferOwedImpact(row), 0);

    ccBalances[account] = openingBalance + charges + transferImpact;
    log(`CC [${account}]: opening=${openingBalance.toFixed(2)}, charges=${charges.toFixed(2)}, transferImpact=${transferImpact.toFixed(2)}, total=${ccBalances[account].toFixed(2)}, rows=${subsequent.length}`);
  });

  // Total CC outstanding (what you owe)
  const totalCcOwed = ccAccounts.reduce((sum, a) => sum + (ccBalances[a] || 0), 0);
  log(`CC accounts detected: ${ccAccounts.join(", ") || "(none)"} | Savings: ${savingsAccounts.join(", ")}`);

  const pendingClaims = computePendingClaimsByAccount(data, account => {
    const key = accountKey(account);
    if (savingsAccountKeys.has(key)) return isAssetAccountIncluded(account);
    return true;
  });

  // For backward-compat the HTML still uses uobOneBalance / assetsBalance IDs;
  // show total CC owed and personal net assets (included savings minus CC debt plus pending receivables).
  const assetsBalance = savingsTotal - totalCcOwed + pendingClaims.total;

  // If there's exactly one CC account keep the label; otherwise show total
  const uobOneEl = document.getElementById("uobOneBalance");
  if (uobOneEl) uobOneEl.innerText = formatCurrency(totalCcOwed);
  renderPendingClaimsAssets(pendingClaims);
  updateAssetScopeBadge(savingsAccounts.length - includedSavingsAccounts.length, excludedSavingsTotal);
  document.getElementById("assetsBalance").innerText = formatCurrency(assetsBalance);
  drawSavingsTable(savingsAccounts, balances, savingsTotal);
}

function drawSavingsTable(accounts, balances, total) {
  const table = document.getElementById("savingsTable");
  table.innerHTML = "";
  const header = document.createElement("tr");
  header.className = "accounts-head-row";
  header.innerHTML = `<th>Use</th><th>Account</th><th>Balance</th>`;
  table.appendChild(header);

  accounts.forEach(account => {
    const included = isAssetAccountIncluded(account);
    const tr = document.createElement("tr");
    if (!included) tr.className = "excluded-account";

    const includeTd = document.createElement("td");
    includeTd.className = "asset-include-cell";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = included;
    checkbox.title = included ? "Included in Net Assets" : "Excluded from Net Assets";
    checkbox.addEventListener("change", () => toggleAssetAccount(account, checkbox.checked));
    includeTd.appendChild(checkbox);

    const accountTd = document.createElement("td");
    accountTd.innerText = account;
    const amountTd = document.createElement("td");
    amountTd.innerText = formatCurrency(balances[account]);
    amountTd.className = "amount-cell";

    tr.appendChild(includeTd);
    tr.appendChild(accountTd);
    tr.appendChild(amountTd);
    table.appendChild(tr);
  });

  const totalRow = document.createElement("tr"); totalRow.className = "total-row";
  const totalCount = document.createElement("td"); totalCount.innerText = "";
  const totalLabel = document.createElement("td"); totalLabel.innerText = "Included savings";
  const totalAmount = document.createElement("td"); totalAmount.innerText = formatCurrency(total); totalAmount.className = "amount-cell";
  totalRow.appendChild(totalCount); totalRow.appendChild(totalLabel); totalRow.appendChild(totalAmount);
  table.appendChild(totalRow);
}

function renderPendingClaimsAssets(summary) {
  const section = document.getElementById("pendingClaimsAssets");
  const line = document.getElementById("pendingClaimsLine");
  const balance = document.getElementById("pendingClaimsBalance");

  if (balance) balance.innerText = "+" + formatCurrency(summary.total);
  if (line) line.style.display = summary.total > 0 ? "flex" : "none";
  if (!section) return;

  section.innerHTML = "";
  if (summary.total <= 0) {
    section.style.display = "none";
    return;
  }

  section.style.display = "block";
  const title = document.createElement("div");
  title.className = "pending-claims-title";
  title.innerHTML = `<span>Pending Claims Receivable</span><strong>+${formatCurrency(summary.total)}</strong>`;
  section.appendChild(title);

  Object.entries(summary.byAccount)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .forEach(([account, amount]) => {
      const row = document.createElement("div");
      row.className = "pending-claim-row";
      const label = document.createElement("span");
      label.textContent = account;
      const value = document.createElement("strong");
      value.textContent = "+" + formatCurrency(amount);
      row.appendChild(label);
      row.appendChild(value);
      section.appendChild(row);
    });
}

function updateAssetScopeBadge(excludedCount, excludedTotal) {
  const badge = document.getElementById("assetScopeBadge");
  if (!badge) return;
  if (excludedCount > 0) {
    badge.textContent = `${excludedCount} excluded`;
    badge.title = `${formatCurrency(excludedTotal)} excluded from Net Assets`;
    badge.classList.add("warn");
  } else {
    badge.textContent = "All included";
    badge.title = "All savings accounts are counted in Net Assets";
    badge.classList.remove("warn");
  }
}

// ─── Tables ────────────────────────────────────────────────────────────────────

function renderRecentTransactions(data) {
  const sorted = [...data].filter(row => parseExcelDate(row["Date"])).sort((a,b) => parseExcelDate(b["Date"]) - parseExcelDate(a["Date"])).slice(0,10);
  const table = document.getElementById("recentTransactionsTable");
  table.innerHTML = `<tr><th>Date</th><th>Transaction</th><th>Sub Category</th><th>Account</th><th>Amount</th></tr>`;
  if (sorted.length === 0) { table.innerHTML += `<tr><td colspan="5" style="text-align:center;color:#999;">No transactions found</td></tr>`; return; }
  sorted.forEach(row => {
    const tr = document.createElement("tr");
    const date = parseExcelDate(row["Date"]);
    tr.innerHTML = `<td>${date ? date.toLocaleDateString("en-SG") : ""}</td><td>${clean(row["Transaction"])}</td><td>${clean(row["Sub Category"])}</td><td>${clean(row["Account"])}</td><td style="text-align:right;font-weight:bold;">${formatCurrency(getAmount(row["Amount"]))}</td>`;
    table.appendChild(tr);
  });
}

function renderTop5Transactions(data) {
  const sorted = [...data].filter(row => parseExcelDate(row["Date"])).sort((a,b) => getAmount(b["Amount"]) - getAmount(a["Amount"])).slice(0,5);
  const table = document.getElementById("top5Table");
  table.innerHTML = `<tr><th>Date</th><th>Transaction</th><th>Sub Category</th><th>Amount</th></tr>`;
  if (sorted.length === 0) { table.innerHTML += `<tr><td colspan="4" style="text-align:center;color:#999;">No transactions found</td></tr>`; return; }
  sorted.forEach(row => {
    const tr = document.createElement("tr");
    const date = parseExcelDate(row["Date"]);
    tr.innerHTML = `<td>${date ? date.toLocaleDateString("en-SG") : ""}</td><td>${clean(row["Transaction"])}</td><td>${clean(row["Sub Category"])}</td><td style="text-align:right;font-weight:bold;">${formatCurrency(getAmount(row["Amount"]))}</td>`;
    table.appendChild(tr);
  });
}

function renderMonthlyAvgByCategory(data) {
  const monthlyTotals = {};
  const monthsSet = new Set();
  data.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    if (!date) return;
    const sub = clean(row["Sub Category"]); if (!sub) return;
    const key = date.getFullYear() + "-" + String(date.getMonth()+1).padStart(2,"0");
    monthsSet.add(key);
    if (!monthlyTotals[sub]) monthlyTotals[sub] = {};
    monthlyTotals[sub][key] = (monthlyTotals[sub][key] || 0) + getAmount(row["Amount"]);
  });
  const totalMonths = monthsSet.size || 1;
  const avgByCategory = Object.entries(monthlyTotals).map(([sub, months]) => {
    const total = Object.values(months).reduce((s,v) => s+v, 0);
    return { sub, avg: total / totalMonths, total };
  }).sort((a,b) => b.avg - a.avg);
  const table = document.getElementById("avgCategoryTable");
  table.innerHTML = `<tr><th>Sub Category</th><th>Monthly Avg</th><th>Total Spent</th></tr>`;
  if (avgByCategory.length === 0) { table.innerHTML += `<tr><td colspan="3" style="text-align:center;color:#999;">No data found</td></tr>`; return; }
  avgByCategory.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${item.sub}</td><td style="text-align:right;">${formatCurrency(item.avg)}</td><td style="text-align:right;font-weight:bold;">${formatCurrency(item.total)}</td>`;
    table.appendChild(tr);
  });
}

function renderIncomeBreakdown(data, filters = {}) {
  const table = document.getElementById("incomeBreakdownTable");
  if (!table) return;

  const totals = {};
  const counts = {};
  const months = new Set();

  (data || []).forEach(row => {
    const date = parseExcelDate(row["Date"]);
    if (date) months.add(date.getFullYear() + "-" + String(date.getMonth()+1).padStart(2,"0"));
    const sub = clean(row["Sub Category"]) || "Uncategorised";
    totals[sub] = (totals[sub] || 0) + getAmount(row["Amount"]);
    counts[sub] = (counts[sub] || 0) + 1;
  });

  const totalIncome = Object.values(totals).reduce((sum, amount) => sum + amount, 0);
  const totalEl = document.getElementById("incomeBreakdownTotal");
  if (totalEl) totalEl.innerText = formatCurrency(totalIncome);

  const monthCount = getBreakdownMonthCount(months, filters);
  const rows = Object.entries(totals)
    .filter(([, total]) => total > 0)
    .map(([sub, total]) => ({
      sub,
      count: counts[sub] || 0,
      total,
      share: totalIncome > 0 ? total / totalIncome : 0,
      avg: total / monthCount
    }))
    .sort((a, b) => b.total - a.total);

  table.innerHTML = `<tr><th>Sub Category</th><th>Count</th><th>Total</th><th>Share</th><th>Avg / Month</th></tr>`;
  if (rows.length === 0) {
    table.innerHTML += `<tr><td colspan="5" style="text-align:center;color:#999;">No income data found</td></tr>`;
    return;
  }

  rows.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.sub}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;">${item.count}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;">${formatCurrency(item.total)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;">${(item.share * 100).toFixed(1)}%</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;">${formatCurrency(item.avg)}</td>
    `;
    table.appendChild(tr);
  });
}

function renderIncomeSummaryCards(data, filters = {}) {
  const stats = getIncomeStats(data, filters);
  setMoneyText("incomePeriodTotal", stats.total, true);
  setMoneyText("incomeMonthlyAverage", stats.avgPerMonth, true);

  const sourceCount = document.getElementById("incomeSourceCount");
  if (sourceCount) sourceCount.textContent = stats.sourceCount;

  const bestMonth = document.getElementById("incomeBestMonth");
  const bestMonthNote = document.getElementById("incomeBestMonthNote");
  if (bestMonth) {
    bestMonth.textContent = stats.bestMonth
      ? formatKpiCurrency(stats.bestMonth.total)
      : "$0.00";
    bestMonth.style.color = "#16a34a";
  }
  if (bestMonthNote) {
    bestMonthNote.textContent = stats.bestMonth
      ? formatDashboardMonth(stats.bestMonth.month)
      : "No income in view";
  }
}

function renderIncomeInsightCards(data, allRows, filters = {}) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const lastMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  const lastMonthYear = currentMonth === 0 ? currentYear - 1 : currentYear;

  const scopedIncomeRows = (allRows || [])
    .filter(isIncomeRow)
    .filter(row => matchesIncomeNonDateFilters(row, filters));

  const thisMonthIncome = scopedIncomeRows
    .filter(row => {
      const d = parseExcelDate(row["Date"]);
      return d && d.getFullYear() === currentYear && d.getMonth() === currentMonth;
    })
    .reduce((sum, row) => sum + getAmount(row["Amount"]), 0);

  const lastMonthIncome = scopedIncomeRows
    .filter(row => {
      const d = parseExcelDate(row["Date"]);
      return d && d.getFullYear() === lastMonthYear && d.getMonth() === lastMonth;
    })
    .reduce((sum, row) => sum + getAmount(row["Amount"]), 0);

  const incomeThisMonthCard = document.getElementById("incomeThisMonthCard");
  if (incomeThisMonthCard) {
    incomeThisMonthCard.textContent = formatCurrency(thisMonthIncome);
    incomeThisMonthCard.style.color = "#16a34a";
  }

  const incomeDiff = thisMonthIncome - lastMonthIncome;
  const incomeVsLastMonthCard = document.getElementById("incomeVsLastMonthCard");
  if (incomeVsLastMonthCard) {
    incomeVsLastMonthCard.textContent = incomeDiff >= 0
      ? `${formatCurrency(incomeDiff)} more`
      : `${formatCurrency(Math.abs(incomeDiff))} less`;
    incomeVsLastMonthCard.style.color = incomeDiff >= 0 ? "#16a34a" : "#dc2626";
  }

  const sourceTotals = {};
  (data || []).forEach(row => {
    const source = clean(row["Sub Category"]) || "Uncategorised";
    sourceTotals[source] = (sourceTotals[source] || 0) + getAmount(row["Amount"]);
  });

  const totalIncome = Object.values(sourceTotals).reduce((sum, amount) => sum + amount, 0);
  const topSource = Object.entries(sourceTotals).sort((a, b) => b[1] - a[1])[0];

  const topSourceCard = document.getElementById("incomeTopSourceCard");
  if (topSourceCard) {
    topSourceCard.textContent = topSource
      ? `${topSource[0]} - ${formatCurrency(topSource[1])}`
      : "-";
  }

  const concentrationCard = document.getElementById("incomeConcentrationCard");
  if (concentrationCard) {
    concentrationCard.textContent = topSource && totalIncome > 0
      ? `${(topSource[1] / totalIncome * 100).toFixed(1)}% from top source`
      : "-";
  }

  const countCard = document.getElementById("incomeFilteredCountCard");
  if (countCard) countCard.textContent = String((data || []).length);
}

function renderIncomeMonthlyTable(data) {
  const table = document.getElementById("incomeMonthlyTable");
  if (!table) return;

  const byMonth = {};
  (data || []).forEach(row => {
    const date = parseExcelDate(row["Date"]);
    if (!date) return;
    const key = dashboardMonthKey(date);
    const source = clean(row["Sub Category"]) || "Uncategorised";
    if (!byMonth[key]) byMonth[key] = { total: 0, count: 0, sources: {} };
    byMonth[key].total += getAmount(row["Amount"]);
    byMonth[key].count += 1;
    byMonth[key].sources[source] = (byMonth[key].sources[source] || 0) + getAmount(row["Amount"]);
  });

  const rows = Object.entries(byMonth)
    .map(([month, value]) => {
      const topSource = Object.entries(value.sources).sort((a, b) => b[1] - a[1])[0];
      return {
        month,
        total: value.total,
        count: value.count,
        topSource: topSource ? topSource[0] : "-",
        topSourceAmount: topSource ? topSource[1] : 0
      };
    })
    .sort((a, b) => b.month.localeCompare(a.month));

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const totalEl = document.getElementById("incomeMonthlyTotal");
  if (totalEl) totalEl.textContent = formatCurrency(total);

  table.innerHTML = `<tr><th>Month</th><th>Income</th><th>Entries</th><th>Top Source</th></tr>`;
  if (!rows.length) {
    table.innerHTML += `<tr><td colspan="4" style="text-align:center;color:#999;">No income data found</td></tr>`;
    return;
  }

  rows.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${formatDashboardMonth(row.month)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;color:#16a34a;">${formatCurrency(row.total)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;">${row.count}</td>
      <td style="text-align:left;font-family:'DM Sans', sans-serif;font-weight:500;">${row.topSource} (${formatCurrency(row.topSourceAmount)})</td>
    `;
    table.appendChild(tr);
  });
}

function getIncomeStats(data, filters = {}) {
  const months = new Set();
  const sourceTotals = {};
  const monthTotals = {};
  let total = 0;

  (data || []).forEach(row => {
    const amount = getAmount(row["Amount"]);
    total += amount;

    const source = clean(row["Sub Category"]) || "Uncategorised";
    sourceTotals[source] = (sourceTotals[source] || 0) + amount;

    const date = parseExcelDate(row["Date"]);
    if (!date) return;
    const month = dashboardMonthKey(date);
    months.add(month);
    monthTotals[month] = (monthTotals[month] || 0) + amount;
  });

  const monthCount = getBreakdownMonthCount(months, filters);
  const bestMonth = Object.entries(monthTotals)
    .map(([month, monthTotal]) => ({ month, total: monthTotal }))
    .sort((a, b) => b.total - a.total)[0] || null;

  return {
    total,
    avgPerMonth: total / monthCount,
    sourceCount: Object.keys(sourceTotals).length,
    bestMonth
  };
}

function matchesIncomeNonDateFilters(row, filters = {}) {
  return filterAllowsValue(filters.mainCategories || [], row["Main Category"])
    && filterAllowsValue(filters.subCategories || [], row["Sub Category"])
    && filterAllowsValue(filters.transactions || [], row["Transaction"]);
}

function dashboardMonthKey(date) {
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
}

function formatDashboardMonth(monthKey) {
  const [year, month] = monthKey.split("-").map(Number);
  if (!Number.isFinite(year) || !Number.isFinite(month)) return monthKey;
  return new Date(year, month - 1, 1).toLocaleString("en-SG", { month: "short", year: "2-digit" });
}

function getBreakdownMonthCount(months, filters) {
  if (filters && filters.fromDate && filters.toDate) {
    const [fy, fm] = filters.fromDate.split("-").map(Number);
    const [ty, tm] = filters.toDate.split("-").map(Number);
    if (Number.isFinite(fy) && Number.isFinite(fm) && Number.isFinite(ty) && Number.isFinite(tm)) {
      return Math.max(1, (ty - fy) * 12 + (tm - fm) + 1);
    }
  }
  return months.size || 1;
}

// ─── Insight Cards ─────────────────────────────────────────────────────────────

function renderInsightCards(data, incomeData, allExpenses) {
  allExpenses = (allExpenses || data).filter(isSpendAnalyticsRow);
  const today = new Date();

  const filteredDates = data.map(r => parseExcelDate(r["Date"])).filter(Boolean);
  const refDate = filteredDates.length > 0
    ? filteredDates.sort((a,b) => b - a)[0]
    : today;
  const activeMonth = refDate.getMonth();
  const activeYear  = refDate.getFullYear();
  const lastMonth     = activeMonth === 0 ? 11 : activeMonth - 1;
  const lastMonthYear = activeMonth === 0 ? activeYear - 1 : activeYear;

  const monthlyTotals = {};
  data.forEach(row => {
    const date = parseExcelDate(row["Date"]); if (!date) return;
    const key = date.getFullYear() + "-" + String(date.getMonth()+1).padStart(2,"0");
    monthlyTotals[key] = (monthlyTotals[key] || 0) + getAmount(row["Amount"]);
  });
  const highestMonth = Object.entries(monthlyTotals).sort((a,b) => b[1]-a[1])[0];
  document.getElementById("highestMonthCard").innerText = highestMonth ? highestMonth[0] + " — " + formatCurrency(highestMonth[1]) : "—";

  const thisMonthExpenses = data.filter(row => { const d = parseExcelDate(row["Date"]); return d && d.getMonth()===activeMonth && d.getFullYear()===activeYear; }).reduce((s,row) => s+getAmount(row["Amount"]), 0);

  const lastMonthExpenses = allExpenses.filter(row => { const d = parseExcelDate(row["Date"]); return d && d.getMonth()===lastMonth && d.getFullYear()===lastMonthYear; }).reduce((s,row) => s+getAmount(row["Amount"]), 0);

  const diff = thisMonthExpenses - lastMonthExpenses;
  const vsEl = document.getElementById("vsLastMonthCard");
  vsEl.innerText = diff >= 0 ? "▲ " + formatCurrency(diff) + " more than last month" : "▼ " + formatCurrency(Math.abs(diff)) + " less than last month";
  vsEl.style.color = diff >= 0 ? "#c0392b" : "#27ae60";

  const catTotals = {};
  data.filter(row => { const d = parseExcelDate(row["Date"]); return d && d.getMonth()===activeMonth && d.getFullYear()===activeYear; })
    .forEach(row => { const sub = clean(row["Sub Category"]); if (sub) catTotals[sub] = (catTotals[sub]||0) + getAmount(row["Amount"]); });
  const topCat = Object.entries(catTotals).sort((a,b) => b[1]-a[1])[0];
  document.getElementById("topCategoryCard").innerText = topCat ? topCat[0] + " — " + formatCurrency(topCat[1]) : "—";

  const daysElapsed = activeMonth === today.getMonth() && activeYear === today.getFullYear()
    ? today.getDate()
    : new Date(activeYear, activeMonth + 1, 0).getDate();
  document.getElementById("dailyAvgCard").innerText = formatCurrency(thisMonthExpenses / daysElapsed) + " / day";

  const thisMonthIncome = (incomeData||[]).filter(row => { const d = parseExcelDate(row["Date"]); return d && d.getMonth()===activeMonth && d.getFullYear()===activeYear; }).reduce((s,row) => s+getAmount(row["Amount"]), 0);
  const thisMonthNet = thisMonthIncome - thisMonthExpenses;
  const netEl = document.getElementById("thisMonthNetCard");
  if (netEl) { netEl.innerText = (thisMonthNet >= 0 ? "+" : "") + formatCurrency(thisMonthNet); netEl.style.color = thisMonthNet >= 0 ? "#27ae60" : "#c0392b"; }
}

// ─── Reset Filters ─────────────────────────────────────────────────────────────

function resetFilters() {
  document.querySelectorAll(".checkbox-menu input[type='checkbox']").forEach(input => {
    input.checked = true;
    input.indeterminate = false;
  });
  document.getElementById("fromDate").value = "";
  document.getElementById("toDate").value = "";
  const filters = getCurrentFilters();
  refreshAllFilters(filters);
  updateDashboard(filters);
  updateActiveFilterBadges(filters);
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

function parseExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  // Excel serial number
  if (typeof value === "number") { const date = XLSX.SSF.parse_date_code(value); return new Date(date.y, date.m-1, date.d); }
  const s = String(value).trim();
  // DD/MM/YYYY or D/M/YYYY — our stored format
  const ddmmyyyy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyy) return new Date(Number(ddmmyyyy[3]), Number(ddmmyyyy[2])-1, Number(ddmmyyyy[1]));
  // YYYY-MM-DD — ISO string (safe: parse parts manually, never use new Date(isoString) to avoid UTC shift)
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3]));
  // Do NOT fall through to new Date(string) — JS parses slash-dates as MM/DD/YYYY (American)
  // which would swap month and day for our DD/MM/YYYY strings.
  return null;
}

function getAmount(value) { const n = Number(String(value).replace(/[$,]/g,"")); return isNaN(n) ? 0 : Math.abs(n); }
function getSignedAmount(value) { if (typeof value === "number") return value; const n = Number(String(value).replace(/\$/g,"").replace(/,/g,"").trim()); return isNaN(n) ? 0 : n; }

function getAccountBalanceImpact(row) {
  const amount = Math.abs(getSignedAmount(row["Amount"]));
  const main = clean(row["Main Category"]).toLowerCase();
  const sub = clean(row["Sub Category"]).toLowerCase();

  if (main === "income") return amount;
  if (main === "transfer") {
    if (sub === "transfer in" || sub === "cc payment in") return amount;
    if (sub === "transfer out" || sub === "cc payment out") return -amount;
    return 0;
  }
  return -amount;
}

function getCreditCardTransferOwedImpact(row) {
  const amount = Math.abs(getSignedAmount(row["Amount"]));
  const sub = clean(row["Sub Category"]).toLowerCase();
  if (sub === "transfer in" || sub === "cc payment in") return -amount;
  if (sub === "transfer out" || sub === "cc payment out") return amount;
  return 0;
}
function formatCurrency(value) { return value.toLocaleString("en-SG", { style:"currency", currency:"SGD", minimumFractionDigits:2, maximumFractionDigits:2 }); }

function formatKpiCurrency(value) {
  return value.toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
