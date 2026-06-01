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

    if (!transactionSheet) {
      throw new Error("Sheet not found: " + CONFIG.sheetName);
    }

    if (!budgetSheet) {
      throw new Error("Sheet not found: " + BUDGET_SHEET);
    }

    budgetTransactions = readTransactionSheet(transactionSheet);

    billsBudget = readBudgetSection(budgetSheet, "A2:B13", "Bills");
    monthlyBudget = readBudgetSection(budgetSheet, "F2:G13", "Monthly Expenses");

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
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false
  });

  const headers = rows[0].map(header => clean(header));

  return rows.slice(1).map(row => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = row[index];
    });

    return obj;
  });
}

function readBudgetSection(sheet, range, type) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: range,
    blankrows: false
  });

  return rows
    .map(row => ({
      type: type,
      category: clean(row[0]),
      allocated: toNumber(row[1])
    }))
    .filter(item => item.category !== "");
}

function addBudgetItem() {
  const type = document.getElementById("budgetType").value;
  const category = clean(document.getElementById("subCategoryInput").value);
  const allocated = Number(document.getElementById("allocatedInput").value);

  if (!category || isNaN(allocated)) {
    alert("Please enter Sub Category and Allocated.");
    return;
  }

  const list = type === "Bills" ? billsBudget : monthlyBudget;

  if (list.length >= 12) {
    alert("Maximum 12 rows allowed for this section.");
    return;
  }

  list.push({
    type: type,
    category: category,
    allocated: allocated
  });

  document.getElementById("subCategoryInput").value = "";
  document.getElementById("allocatedInput").value = "";

  renderBudget();
}

function deleteBudgetItem(type, index) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  list.splice(index, 1);
  renderBudget();
}

function updateBudgetCategory(type, index, value) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  list[index].category = clean(value);
  renderBudget();
}

function updateBudgetAllocated(type, index, value) {
  const list = type === "Bills" ? billsBudget : monthlyBudget;
  list[index].allocated = Number(value) || 0;
  renderBudget();
}

function renderBudget() {
  const computedBills = billsBudget.map(item => computeBudgetRow(item));
  const computedMonthly = monthlyBudget.map(item => computeBudgetRow(item));

  renderBudgetTable("billsTable", "Bills", computedBills);
  renderBudgetTable("monthlyTable", "Monthly Expenses", computedMonthly);

  updateBudgetCards([...computedBills, ...computedMonthly]);
}

function computeBudgetRow(item) {
  const spent = calculateSpentForCurrentMonth(item.type, item.category);

  return {
    ...item,
    spent: spent,
    balance: item.allocated - spent
  };
}

function calculateSpentForCurrentMonth(mainCategory, subCategory) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();

  return budgetTransactions
    .filter(row => {
      const rowDate = parseBudgetDate(row["Date"]);

      if (!rowDate) return false;

      const isCurrentMonth =
        rowDate.getFullYear() === currentYear &&
        rowDate.getMonth() === currentMonth;

      return (
        isCurrentMonth &&
        clean(row["Main Category"]).toLowerCase() === clean(mainCategory).toLowerCase() &&
        clean(row["Sub Category"]).toLowerCase() === clean(subCategory).toLowerCase()
      );
    })
    .reduce((sum, row) => sum + getAmount(row["Amount"]), 0);
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
    </tr>
  `;

  let totalAllocated = 0;
  let totalSpent = 0;
  let totalBalance = 0;

  rows.forEach((row, index) => {
    totalAllocated += row.allocated;
    totalSpent += row.spent;
    totalBalance += row.balance;

    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>
        <input
          value="${escapeHtml(row.category)}"
          onchange="updateBudgetCategory('${type}', ${index}, this.value)"
        >
      </td>

      <td>
        <input
          type="number"
          step="0.01"
          value="${row.allocated}"
          onchange="updateBudgetAllocated('${type}', ${index}, this.value)"
        >
      </td>

      <td>${formatCurrency(row.spent)}</td>
      <td>${formatCurrency(row.balance)}</td>

      <td>
        <button onclick="deleteBudgetItem('${type}', ${index})">Delete</button>
      </td>
    `;

    table.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.className = "total-row";

  totalRow.innerHTML = `
    <td><strong>Total</strong></td>
    <td><strong>${formatCurrency(totalAllocated)}</strong></td>
    <td><strong>${formatCurrency(totalSpent)}</strong></td>
    <td><strong>${formatCurrency(totalBalance)}</strong></td>
    <td></td>
  `;

  table.appendChild(totalRow);
}

async function saveBudgetSetupToExcel() {
  try {
    log("Saving Budget Setup to Excel...");

    await writeBudgetSetupRange("A2:B13", buildSaveValues(billsBudget));
    await writeBudgetSetupRange("F2:G13", buildSaveValues(monthlyBudget));

    alert("Budget saved to Excel.");
    log("Budget saved to Excel.");
  } catch (err) {
    log("SAVE ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function buildSaveValues(list) {
  const values = list.map(item => [
    item.category,
    item.allocated
  ]);

  while (values.length < 12) {
    values.push(["", ""]);
  }

  return values;
}

function updateBudgetCards(rows) {
  const totalAllocated = rows.reduce((sum, row) => sum + row.allocated, 0);
  const totalSpent = rows.reduce((sum, row) => sum + row.spent, 0);
  const totalBalance = rows.reduce((sum, row) => sum + row.balance, 0);

  const foodRow = rows.find(row =>
    clean(row.category).toLowerCase() === "food"
  );

  const foodBalance = foodRow ? foodRow.balance : 0;
  const daysRemaining = getDaysRemainingInMonth();

  document.getElementById("totalAllocated").innerText =
    formatCurrency(totalAllocated);

  document.getElementById("totalSpent").innerText =
    formatCurrency(totalSpent);

  document.getElementById("totalBalance").innerText =
    formatCurrency(totalBalance);

  document.getElementById("foodPerDay").innerText =
    formatCurrency(foodBalance / daysRemaining);

  document.getElementById("monthlyPerDay").innerText =
    formatCurrency(totalBalance / daysRemaining);
}

function getDaysRemainingInMonth() {
  const today = new Date();
  const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  return Math.max(lastDay.getDate() - today.getDate() + 1, 1);
}

function parseBudgetDate(value) {
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

function toNumber(value) {
  if (typeof value === "number") return value;

  const number = Number(String(value).replace(/[$,]/g, ""));
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

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}