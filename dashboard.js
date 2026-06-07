let allTransactions = [];
let isUpdatingFilters = false;
let savingsAccountNames = [];

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
      const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, blankrows: false });
      savingsAccountNames = allRows.slice(1)
        .filter(row => String(row[10] ?? "").trim() === "Savings")
        .map(row => String(row[9] ?? "").trim())
        .filter(Boolean);
    }
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

function prepareTransactions(rows) {
  const headers = rows[0].map(header => clean(header));
  allTransactions = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });
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
  return Array.from(document.querySelectorAll(`#${menuId} input[type="checkbox"]:checked`)).map(input => clean(input.value));
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

function isIncomeRow(row) {
  return clean(row["Main Category"]).toLowerCase() === "income"
    && clean(row["Sub Category"]) !== "Opening Balance";
}

function rowMatchesFilters(row, filters, ignoreField = null) {
  const rowDate = parseExcelDate(row["Date"]);
  if (!rowDate) return false;
  if (ignoreField !== "mainCategory" && filters.mainCategories.length > 0 && !filters.mainCategories.includes(clean(row["Main Category"]))) return false;
  if (ignoreField !== "subCategory" && filters.subCategories.length > 0 && !filters.subCategories.includes(clean(row["Sub Category"]))) return false;
  if (ignoreField !== "transaction" && filters.transactions.length > 0 && !filters.transactions.includes(clean(row["Transaction"]))) return false;
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

function refreshAllFilters(filters = getCurrentFilters()) {
  isUpdatingFilters = true;
  const expenseOnly = allTransactions.filter(isExpenseRow);
  rebuildCheckboxMenu("mainCategoryMenu", "Main Category", expenseOnly.filter(row => rowMatchesFilters(row, filters, "mainCategory")), filters.mainCategories);
  rebuildCheckboxMenu("subCategoryMenu", "Sub Category", expenseOnly.filter(row => rowMatchesFilters(row, filters, "subCategory")), filters.subCategories);
  rebuildCheckboxMenu("transactionMenu", "Transaction", expenseOnly.filter(row => rowMatchesFilters(row, filters, "transaction")), filters.transactions);
  isUpdatingFilters = false;
}

function rebuildCheckboxMenu(menuId, columnName, rows, selectedValues) {
  const menu = document.getElementById(menuId);
  const values = [...new Set([
    ...rows.map(row => clean(row[columnName])).filter(v => v !== ""),
    ...selectedValues
  ])].sort();
  menu.innerHTML = "";
  values.forEach(value => {
    const label = document.createElement("label");
    label.className = "checkbox-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = value;
    checkbox.checked = selectedValues.includes(value);
    checkbox.addEventListener("change", handleFilterChange);
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(" " + value));
    menu.appendChild(label);
  });
}

function updateActiveFilterBadges(filters = getCurrentFilters()) {
  const badges = [];
  if (filters.mainCategories.length > 0) badges.push(...filters.mainCategories.map(v => `Main: ${v}`));
  if (filters.subCategories.length > 0) badges.push(...filters.subCategories.map(v => `Sub: ${v}`));
  if (filters.transactions.length > 0) badges.push(...filters.transactions.map(v => `Tx: ${v}`));
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

function updateDashboard(filters = getCurrentFilters()) {
  const filtered = allTransactions.filter(row => isExpenseRow(row) && rowMatchesFilters(row, filters));
  const incomeFiltered = allTransactions.filter(row => {
      if (!isIncomeRow(row)) return false;
      const rowDate = parseExcelDate(row["Date"]);
      if (!rowDate) return false;
      if (filters.fromDate) {
        const [fy, fm, fd] = filters.fromDate.split("-").map(Number);
        if (rowDate < new Date(fy, fm - 1, fd)) return false;
      }
      if (filters.toDate) {
        const [ty, tm, td] = filters.toDate.split("-").map(Number);
        if (rowDate > new Date(ty, tm - 1, td, 23, 59, 59)) return false;
      }
      return true;
    });

  const totalExpenses = filtered.reduce((sum, row) => sum + getAmount(row["Amount"]), 0);
  const totalIncome = incomeFiltered.reduce((sum, row) => sum + getAmount(row["Amount"]), 0);
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
  drawMonthlyExpenseChart(filtered, incomeFiltered, filters);
  drawSubCategoryMonthlyChart(filtered, filters);
  drawIncomeSubCategoryChart(incomeFiltered, filters);
  renderRecentTransactions(filtered);
  renderTop5Transactions(filtered);
  renderMonthlyAvgByCategory(filtered);
  renderIncomeBreakdown(incomeFiltered, filters);
  renderInsightCards(filtered, incomeFiltered, allTransactions);
}

// ─── Finance Cards ─────────────────────────────────────────────────────────────

function updateFinanceCards(data) {
  const savingsAccounts = savingsAccountNames;
  const balances = {};

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

    const income = subsequent.filter(row => clean(row["Main Category"]).toLowerCase() === "income")
      .reduce((sum, row) => sum + getSignedAmount(row["Amount"]), 0);
    const nonIncome = subsequent.filter(row => clean(row["Main Category"]).toLowerCase() !== "income")
      .reduce((sum, row) => sum + getSignedAmount(row["Amount"]), 0);

    balances[account] = openingBalance + income - nonIncome;
  });

  const savingsTotal = savingsAccounts.reduce((sum, account) => sum + (balances[account] || 0), 0);

  // ── Credit card balance ────────────────────────────────────────────────────
  // Outstanding CC balance = opening balance + charges on the card - payments made to the card.
  // Opening balance row is stored as Income/Opening Balance with a positive amount (what you owe).
  // Expense rows on the CC account add to what you owe.
  // CC-pay rows (Main Category = "Transfer", Account = CC) reduce what you owe.
  const ccAccountNames = (() => {
    try {
      const allRows = XLSX.utils.sheet_to_json(
        null, { header: 1, blankrows: false }
      );
      return [];
    } catch { return []; }
  })();

  // We know the CC accounts from Budget Setup (type = "Credit Card").
  // savingsAccountNames only holds Savings accounts; we need to derive CC names separately.
  // They are already available via the workbook loaded at page start — re-derive from allTransactions
  // by looking at accounts NOT in savingsAccountNames that appear in the data.
  // Simpler: replicate the same opening-balance logic used for savings, but for CC accounts the
  // balance is: opening + expenses charged to it - payments made to it (transfers in).
  const allAccountsInData = [...new Set(data.map(row => clean(row["Account"])).filter(Boolean))];
  const ccAccounts = allAccountsInData.filter(a => !savingsAccounts.includes(a));

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
    // If a row has been fully claimed (Claim Status = "Claimed"), the reimbursed
    // portion is subtracted — it never truly cost you that amount on the CC.
    const charges = subsequent
      .filter(row => {
        const cat = clean(row["Main Category"]).toLowerCase();
        return cat !== "income" && cat !== "transfer";
      })
      .reduce((sum, row) => {
        const amount = Math.abs(getSignedAmount(row["Amount"]));
        const claimStatusKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claim status") || "Claim Status";
        const claimStatus = clean(row[claimStatusKey]).toLowerCase();
        if (claimStatus !== "claimed") return sum + amount;
        // Only subtract the actual claim amount that was reimbursed
        const claimAmountKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claim amount") || "Claim Amount";
        const claimAmount = Math.abs(getSignedAmount(row[claimAmountKey]));
        const reimbursed = claimAmount > 0 ? Math.min(claimAmount, amount) : amount;
        return sum + Math.max(0, amount - reimbursed);
      }, 0);

    // Payments: transfer rows where this CC is the destination (reduce balance owed)
    // These show up as Main Category = Transfer with Account = CC name
    const payments = subsequent
      .filter(row => clean(row["Main Category"]).toLowerCase() === "transfer")
      .reduce((sum, row) => sum + Math.abs(getSignedAmount(row["Amount"])), 0);

    ccBalances[account] = openingBalance + charges - payments;
    log(`CC [${account}]: opening=${openingBalance.toFixed(2)}, charges=${charges.toFixed(2)}, payments=${payments.toFixed(2)}, total=${ccBalances[account].toFixed(2)}, rows=${subsequent.length}`);
  });

  // Total CC outstanding (what you owe)
  const totalCcOwed = ccAccounts.reduce((sum, a) => sum + (ccBalances[a] || 0), 0);
  log(`CC accounts detected: ${ccAccounts.join(", ") || "(none)"} | Savings: ${savingsAccounts.join(", ")}`);

  // For backward-compat the HTML still uses uobOneBalance / assetsBalance IDs;
  // show total CC owed and net assets (savings minus CC debt)
  const assetsBalance = savingsTotal - totalCcOwed;

  // If there's exactly one CC account keep the label; otherwise show total
  const uobOneEl = document.getElementById("uobOneBalance");
  if (uobOneEl) uobOneEl.innerText = formatCurrency(totalCcOwed);
  document.getElementById("assetsBalance").innerText = formatCurrency(assetsBalance);
  drawSavingsTable(savingsAccounts, balances, savingsTotal);
}

function drawSavingsTable(accounts, balances, total) {
  const table = document.getElementById("savingsTable");
  table.innerHTML = "";
  accounts.forEach(account => {
    const tr = document.createElement("tr");
    const accountTd = document.createElement("td"); accountTd.innerText = account;
    const amountTd = document.createElement("td"); amountTd.innerText = formatCurrency(balances[account]); amountTd.className = "amount-cell";
    tr.appendChild(accountTd); tr.appendChild(amountTd);
    table.appendChild(tr);
  });
  const totalRow = document.createElement("tr"); totalRow.className = "total-row";
  const totalLabel = document.createElement("td"); totalLabel.innerText = "Total";
  const totalAmount = document.createElement("td"); totalAmount.innerText = formatCurrency(total); totalAmount.className = "amount-cell";
  totalRow.appendChild(totalLabel); totalRow.appendChild(totalAmount);
  table.appendChild(totalRow);
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
  allExpenses = (allExpenses || data).filter(isExpenseRow);
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
  document.querySelectorAll(".checkbox-menu input[type='checkbox']").forEach(input => { input.checked = false; });
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
function formatCurrency(value) { return value.toLocaleString("en-SG", { style:"currency", currency:"SGD", minimumFractionDigits:2, maximumFractionDigits:2 }); }

function formatKpiCurrency(value) {
  return value.toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}