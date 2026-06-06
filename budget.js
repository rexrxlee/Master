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
}

function updateBudgetAllocated(type, index, value) {
  (type === "Bills" ? billsBudget : monthlyBudget)[index].allocated = Number(value) || 0;
}

function renderBudget() {
  const computedBills   = billsBudget.map(item => computeBudgetRow(item));
  const computedMonthly = monthlyBudget.map(item => computeBudgetRow(item));

  renderBudgetTable("billsTable",   "Bills",            computedBills);
  renderBudgetTable("monthlyTable", "Monthly Expenses", computedMonthly);
  updateBudgetCards([...computedBills, ...computedMonthly]);
}

function computeBudgetRow(item) {
  const spent = calculateSpentForCurrentMonth(item.type, item.category);
  return { ...item, spent, balance: item.allocated - spent };
}

/**
 * Calculates spending for the current month.
 * Claimable rows only count the non-claimable portion against the budget.
 * Example: $100 expense with $40 claim amount counts as $60 spent.
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
    <td><strong>${formatCurrency(totalBalance)}</strong></td>
    <td></td>`;
  table.appendChild(totalRow);
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

function updateBudgetCards(rows) {
  const totalAllocated = rows.reduce((s, r) => s + r.allocated, 0);
  const totalSpent     = rows.reduce((s, r) => s + r.spent,     0);
  const totalBalance   = rows.reduce((s, r) => s + r.balance,   0);

  const foodRow     = rows.find(r => clean(r.category).toLowerCase() === "food");
  const foodBalance = foodRow ? foodRow.balance : 0;
  const daysLeft    = getDaysRemainingInMonth();

  document.getElementById("totalAllocated").innerText = formatCurrency(totalAllocated);
  document.getElementById("totalSpent").innerText     = formatCurrency(totalSpent);
  document.getElementById("totalBalance").innerText   = formatCurrency(totalBalance);
  document.getElementById("foodPerDay").innerText     = formatCurrency(foodBalance / daysLeft);
  document.getElementById("monthlyPerDay").innerText  = formatCurrency(totalBalance / daysLeft);
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

  const claimAmountKey = Object.keys(row).find(k => k.trim().toLowerCase() === "claim amount") || "Claim Amount";
  const claimAmountRaw = getAmount(row[claimAmountKey]);
  const claimAmount = claimAmountRaw > 0 ? claimAmountRaw : amount;
  return Math.max(0, amount - Math.min(amount, claimAmount));
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
