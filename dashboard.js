let allTransactions = [];
let isUpdatingFilters = false;

async function loadDashboard() {
  try {
    clearOutput();

    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();

    log("Reading workbook...");
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const worksheet = workbook.Sheets[CONFIG.sheetName];

    if (!worksheet) {
      throw new Error("Sheet not found: " + CONFIG.sheetName);
    }

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      blankrows: false
    });

    prepareTransactions(rows);
    setupFilters();
    refreshAllFilters();
    updateDashboard();

    log("Dashboard loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function clean(value) {
  return String(value ?? "").trim();
}

function prepareTransactions(rows) {
  const headers = rows[0].map(header => clean(header));

  allTransactions = rows.slice(1).map(row => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = row[index];
    });

    return obj;
  });
}

function setupFilters() {
  document.getElementById("fromDate").addEventListener("change", handleFilterChange);
  document.getElementById("toDate").addEventListener("change", handleFilterChange);

  document.addEventListener("click", function(event) {
    if (!event.target.closest(".filter-box")) {
      document.querySelectorAll(".checkbox-menu").forEach(menu => {
        menu.style.display = "none";
      });
    }
  });
}

function toggleDropdown(menuId) {
  const menu = document.getElementById(menuId);
  menu.style.display = menu.style.display === "block" ? "none" : "block";
}

function getCheckedValues(menuId) {
  return Array.from(
    document.querySelectorAll(`#${menuId} input[type="checkbox"]:checked`)
  ).map(input => clean(input.value));
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

  refreshAllFilters();
  updateDashboard();
}

function rowMatchesFilters(row, filters, ignoreField = null) {
  const rowDate = parseExcelDate(row["Date"]);

  if (!rowDate) return false;

  if (
    ignoreField !== "mainCategory" &&
    filters.mainCategories.length > 0 &&
    !filters.mainCategories.includes(clean(row["Main Category"]))
  ) return false;

  if (
    ignoreField !== "subCategory" &&
    filters.subCategories.length > 0 &&
    !filters.subCategories.includes(clean(row["Sub Category"]))
  ) return false;

  if (
    ignoreField !== "transaction" &&
    filters.transactions.length > 0 &&
    !filters.transactions.includes(clean(row["Transaction"]))
  ) return false;

  if (
    ignoreField !== "date" &&
    filters.fromDate &&
    rowDate < new Date(filters.fromDate)
  ) return false;

  if (
    ignoreField !== "date" &&
    filters.toDate &&
    rowDate > new Date(filters.toDate)
  ) return false;

  return true;
}

function refreshAllFilters() {
  isUpdatingFilters = true;

  const filters = getCurrentFilters();

  rebuildCheckboxMenu(
    "mainCategoryMenu",
    "Main Category",
    allTransactions.filter(row =>
      rowMatchesFilters(row, filters, "mainCategory")
    ),
    filters.mainCategories
  );

  rebuildCheckboxMenu(
    "subCategoryMenu",
    "Sub Category",
    allTransactions.filter(row =>
      rowMatchesFilters(row, filters, "subCategory")
    ),
    filters.subCategories
  );

  rebuildCheckboxMenu(
    "transactionMenu",
    "Transaction",
    allTransactions.filter(row =>
      rowMatchesFilters(row, filters, "transaction")
    ),
    filters.transactions
  );

  isUpdatingFilters = false;
}

function rebuildCheckboxMenu(menuId, columnName, rows, selectedValues) {
  const menu = document.getElementById(menuId);

  const values = [...new Set(
    rows
      .map(row => clean(row[columnName]))
      .filter(value => value !== "")
  )].sort();

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

function updateDashboard() {
  const filters = getCurrentFilters();

  const filtered = allTransactions.filter(row =>
    rowMatchesFilters(row, filters)
  );

  document.getElementById("totalExpenses").innerText =
    filtered
      .reduce((sum, row) => sum + getAmount(row["Amount"]), 0)
      .toLocaleString("en-SG", {
        style: "currency",
        currency: "SGD",
        maximumFractionDigits: 0
      });

  document.getElementById("transactionCount").innerText = filtered.length;

  updateFinanceCards(allTransactions);

  drawMonthlyExpenseChart(filtered);
  drawSubCategoryMonthlyChart(filtered);
}

function updateFinanceCards(data) {
  const fixedRHQ = 163.98;

  const savingsAccounts = [
    "UOB Savings",
    "UOB Stash",
    "RHQ",
    "UOB CSA"
  ];

  const balances = {};

  savingsAccounts.forEach(account => {
    const income = data
      .filter(row =>
        clean(row["Account"]) === account &&
        clean(row["Main Category"]).toLowerCase() === "income"
      )
      .reduce((sum, row) => sum + getSignedAmount(row["Amount"]), 0);

    const nonIncome = data
      .filter(row =>
        clean(row["Account"]) === account &&
        clean(row["Main Category"]).toLowerCase() !== "income"
      )
      .reduce((sum, row) => sum + getSignedAmount(row["Amount"]), 0);

    balances[account] = income - nonIncome;
  });

  balances["RHQ"] = fixedRHQ;
  balances["UOB Stash"] = balances["UOB Stash"] - fixedRHQ;

  const savingsTotal = savingsAccounts.reduce(
    (sum, account) => sum + balances[account],
    0
  );

  const uobOneBalance = data
    .filter(row => clean(row["Account"]) === "UOB ONE")
    .reduce((sum, row) => sum + getSignedAmount(row["Amount"]), 0);

  const assetsBalance = savingsTotal - uobOneBalance;

  document.getElementById("uobOneBalance").innerText =
    formatCurrency(uobOneBalance);

  document.getElementById("assetsBalance").innerText =
    formatCurrency(assetsBalance);

  drawSavingsTable(savingsAccounts, balances, savingsTotal);
}

function drawSavingsTable(accounts, balances, total) {
  const table = document.getElementById("savingsTable");
  table.innerHTML = "";

  accounts.forEach(account => {
    const tr = document.createElement("tr");

    const accountTd = document.createElement("td");
    accountTd.innerText = account;

    const amountTd = document.createElement("td");
    amountTd.innerText = formatCurrency(balances[account]);
    amountTd.className = "amount-cell";

    tr.appendChild(accountTd);
    tr.appendChild(amountTd);

    table.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.className = "total-row";

  const totalLabel = document.createElement("td");
  totalLabel.innerText = "Total";

  const totalAmount = document.createElement("td");
  totalAmount.innerText = formatCurrency(total);
  totalAmount.className = "amount-cell";

  totalRow.appendChild(totalLabel);
  totalRow.appendChild(totalAmount);

  table.appendChild(totalRow);
}

function resetFilters() {
  document.querySelectorAll(".checkbox-menu input[type='checkbox']").forEach(input => {
    input.checked = false;
  });

  document.getElementById("fromDate").value = "";
  document.getElementById("toDate").value = "";

  refreshAllFilters();
  updateDashboard();
}

function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    return new Date(date.y, date.m - 1, date.d);
  }

  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function getAmount(value) {
  const number = Number(String(value).replace(/[$,]/g, ""));
  return isNaN(number) ? 0 : Math.abs(number);
}

function getSignedAmount(value) {
  if (typeof value === "number") return value;

  const cleaned = String(value)
    .replace(/\$/g, "")
    .replace(/,/g, "")
    .trim();

  const number = Number(cleaned);

  return isNaN(number) ? 0 : number;
}

function formatCurrency(value) {
  return value.toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}