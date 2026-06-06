let accountsList = [];
let incomeSubCategoryList = [];

const ACCOUNTS_FULL_RANGE = "J2:K10";
const INCOME_SUBS_RANGE = "Q2:Q10";
const MAX_ACCOUNTS = 9;
const MAX_INCOME_SUBS = 9;

async function loadAccountsPage() {
  try {
    clearOutput();
    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (!budgetSheet) throw new Error("Sheet not found: Budget Setup");
    accountsList = readAccountsFromSheet(budgetSheet);
    incomeSubCategoryList = readIncomeSubsFromSheet(budgetSheet);
    renderAccountsTable();
    renderIncomeSubsTable();
    log("Accounts loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

function readAccountsFromSheet(sheet) {
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => ({ name: String(row[9] ?? "").trim(), type: String(row[10] ?? "Savings").trim() }))
    .filter(a => a.name !== "");
}

function readIncomeSubsFromSheet(sheet) {
  // Column Q = index 16
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => String(row[16] ?? "").trim())
    .filter(v => v !== "");
}

function addAccount() {
  const name = document.getElementById("accountNameInput").value.trim();
  const type = document.getElementById("accountTypeInput").value;
  if (!name) { alert("Please enter an account name."); return; }
  if (accountsList.length >= MAX_ACCOUNTS) { alert("Maximum " + MAX_ACCOUNTS + " accounts allowed."); return; }
  if (accountsList.find(a => a.name.toLowerCase() === name.toLowerCase())) { alert("Account already exists."); return; }
  accountsList.push({ name, type });
  document.getElementById("accountNameInput").value = "";
  renderAccountsTable();
}

function deleteAccount(index) { accountsList.splice(index, 1); renderAccountsTable(); }
function updateAccountName(index, value) { accountsList[index].name = value.trim(); }
function updateAccountType(index, value) { accountsList[index].type = value; }

function renderAccountsTable() {
  const table = document.getElementById("accountsTable");
  table.innerHTML = `<tr><th>Account Name</th><th>Type</th><th></th></tr>`;
  accountsList.forEach((account, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeHtml(account.name)}" onchange="updateAccountName(${index}, this.value)"></td>
      <td>
        <select onchange="updateAccountType(${index}, this.value)">
          <option value="Savings" ${account.type === "Savings" ? "selected" : ""}>Savings</option>
          <option value="Credit Card" ${account.type === "Credit Card" ? "selected" : ""}>Credit Card</option>
        </select>
      </td>
      <td><button onclick="deleteAccount(${index})">Delete</button></td>
    `;
    table.appendChild(tr);
  });
}

function addIncomeSub() {
  const name = document.getElementById("incomeSubInput").value.trim();
  if (!name) { alert("Please enter a sub-category name."); return; }
  if (incomeSubCategoryList.length >= MAX_INCOME_SUBS) { alert("Maximum " + MAX_INCOME_SUBS + " income sub-categories allowed."); return; }
  if (incomeSubCategoryList.find(s => s.toLowerCase() === name.toLowerCase())) { alert("Already exists."); return; }
  incomeSubCategoryList.push(name);
  document.getElementById("incomeSubInput").value = "";
  renderIncomeSubsTable();
}

function deleteIncomeSub(index) { incomeSubCategoryList.splice(index, 1); renderIncomeSubsTable(); }
function updateIncomeSub(index, value) { incomeSubCategoryList[index] = value.trim(); }

function renderIncomeSubsTable() {
  const table = document.getElementById("incomeSubsTable");
  table.innerHTML = `<tr><th>Income Sub Category</th><th></th></tr>`;
  incomeSubCategoryList.forEach((sub, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeHtml(sub)}" onchange="updateIncomeSub(${index}, this.value)"></td>
      <td><button onclick="deleteIncomeSub(${index})">Delete</button></td>
    `;
    table.appendChild(tr);
  });
}

async function saveAccountsToExcel() {
  try {
    log("Saving accounts to Excel...");
    const acctValues = accountsList.map(a => [a.name, a.type]);
    while (acctValues.length < MAX_ACCOUNTS) acctValues.push(["", ""]);
    await writeBudgetSetupRange(ACCOUNTS_FULL_RANGE, acctValues);
    log("Accounts saved.");

    log("Saving income sub-categories to column Q...");
    const subValues = incomeSubCategoryList.map(s => [s]);
    while (subValues.length < MAX_INCOME_SUBS) subValues.push([""]);
    await writeBudgetSetupRange(INCOME_SUBS_RANGE, subValues);

    alert("All saved to Excel.");
    log("Done.");
  } catch (err) {
    log("SAVE ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function escapeHtml(value) {
  return String(value).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
