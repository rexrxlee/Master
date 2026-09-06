// Business balances are portions of cash, not additional copies of a bank balance.
function businessKey(value) { return String(value ?? "").trim().toLowerCase(); }
function isBusinessAccount(account) { return businessKey(account?.type) === "business"; }
function readAccountBankGroups(sheet) {
  try { return JSON.parse(sheet?.AH2?.v || "{}"); } catch (_) { return {}; }
}
function expenseOwner(row) { return String(row.expenseFor ?? row["Expense For"] ?? "").trim(); }
function isBusinessTransaction(row, accounts) {
  return !!expenseOwner(row) || accounts.some(account => isBusinessAccount(account) &&
    businessKey(account.name) === businessKey(row.account ?? row["Account"]));
}
function businessTransactionRows(rows) {
  return rows.map(row => ({
    account: row.account ?? row["Account"],
    main: businessKey(row.mainCat ?? row["Main Category"]),
    sub: businessKey(row.subCat ?? row["Sub Category"]),
    amount: Number(String(row.amount ?? row["Amount"] ?? 0).replace(/[$,]/g, "")) || 0,
    owner: expenseOwner(row)
  }));
}
function businessOutstanding(accounts, rows) {
  const result = [];
  const transactions = businessTransactionRows(rows);
  accounts.filter(isBusinessAccount).forEach(business => {
    const byAccount = new Map();
    transactions.forEach(row => {
      if (businessKey(row.owner) !== businessKey(business.name)) return;
      if (!row.account || businessKey(row.account) === businessKey(business.name)) return;
      let impact = 0;
      if (row.main !== "income" && row.main !== "transfer") impact = row.amount;
      if (row.main === "transfer" && ["cc payment in", "transfer in"].includes(row.sub)) impact = -Math.abs(row.amount);
      const key = businessKey(row.account);
      const entry = byAccount.get(key) || { business: business.name, account: row.account, amount: 0 };
      entry.amount += impact;
      byAccount.set(key, entry);
    });
    byAccount.forEach(entry => {
      entry.amount = Math.round(entry.amount * 100) / 100;
      if (entry.amount > 0) result.push(entry);
    });
  });
  return result;
}
