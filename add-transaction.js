// ── State ─────────────────────────────────────────────────────────────────────
let categories = {};
let accounts = [];
let incomeSubCategories = [];
let currentMode = "transaction";
let isClaimable = false;
let recentRows = [];               // raw rows from sheet
let filteredRecentRows = [];       // after applying search/category filter
let recurringSuggestions = [];      // strict recurring matches from transaction history

const SMART_CATEGORY_RULES = [
  {
    hints: ["food"],
    keywords: ["food", "lunch", "dinner", "breakfast", "coffee", "cafe", "restaurant", "eat", "meal", "snack", "kopi", "yakun", "mcdonald", "kfc", "grabfood", "deliveroo", "grocery", "groceries", "fairprice", "ntuc", "sheng siong", "supermarket"]
  },
  {
    hints: ["transport"],
    keywords: ["transport", "grab", "gojek", "taxi", "cab", "tada", "comfort", "bus", "mrt", "train", "parking", "petrol", "fuel", "erp", "toll"]
  },
  {
    hints: ["medical", "health"],
    keywords: ["medical", "clinic", "doctor", "dentist", "dental", "pharmacy", "hospital", "medicine", "medication", "health"]
  },
  {
    hints: ["pet"],
    keywords: ["pet", "vet", "veterinary", "dog", "cat", "grooming", "paw"]
  },
  {
    hints: ["child", "kid", "baby"],
    keywords: ["child", "children", "kid", "kids", "baby", "school", "tuition", "childcare", "diaper", "toys"]
  },
  {
    hints: ["bill", "utility", "utilities", "phone", "internet", "insurance"],
    keywords: ["bill", "utility", "utilities", "phone", "mobile", "internet", "wifi", "electricity", "water", "insurance", "subscription"]
  },
  {
    hints: ["other", "others", "misc"],
    keywords: ["other", "others", "misc", "miscellaneous"]
  }
];

const SMART_MONTHS = {
  jan: 1, january: 1,
  feb: 2, february: 2,
  mar: 3, march: 3,
  apr: 4, april: 4,
  may: 5,
  jun: 6, june: 6,
  jul: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  oct: 10, october: 10,
  nov: 11, november: 11,
  dec: 12, december: 12
};

const SMART_WEEKDAYS = {
  sun: 0, sunday: 0,
  mon: 1, monday: 1,
  tue: 2, tues: 2, tuesday: 2,
  wed: 3, wednesday: 3,
  thu: 4, thur: 4, thurs: 4, thursday: 4,
  fri: 5, friday: 5,
  sat: 6, saturday: 6
};

const SMART_STOP_WORDS = new Set([
  "a", "an", "and", "at", "by", "for", "from", "in", "into", "my", "of", "on", "or", "paid", "pay", "sgd", "the", "to", "using", "via", "with"
]);

// ── Page load ─────────────────────────────────────────────────────────────────
async function loadAddTransactionPage() {
  try {
    clearOutput();
    log("Loading data...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Load categories from Budget Setup
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (budgetSheet) {
      categories = readCategories(budgetSheet);
      incomeSubCategories = readIncomeSubCategories(budgetSheet);
    }

    // Load accounts
    if (budgetSheet) {
      const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, blankrows: false });
      accounts = allRows.slice(1)
        .map(row => ({ name: String(row[9] ?? "").trim(), type: String(row[10] ?? "Savings").trim() }))
        .filter(a => a.name !== "");
    }

    // Load recent transactions
    const txSheet = workbook.Sheets[CONFIG.sheetName];
    if (txSheet) {
      const allTxRows = XLSX.utils.sheet_to_json(txSheet, { header: 1, blankrows: false });
      // headers: A=Date B=Transaction C=Amount D=MainCategory E=SubCategory F=Account G=Claimable H=ClaimStatus I=ClaimAmount J=ClaimAccount
      recentRows = allTxRows.slice(1)
        .map((row, i) => ({
          _rowIndex: i + 2, // 1-based Excel row (header is row 1)
          date:        row[0] ?? "",
          transaction: String(row[1] ?? ""),
          amount:      row[2] ?? "",
          mainCat:     String(row[3] ?? ""),
          subCat:      String(row[4] ?? ""),
          account:     String(row[5] ?? ""),
          claimable:   String(row[6] ?? "").trim(),   // "Yes" or ""
          claimStatus: String(row[7] ?? "").trim(),   // "Pending" | "Claimed" | ""
          claimAmount: row[8] ?? "",
          claimAccount:String(row[9] ?? "").trim(),
        }))
        .filter(r => r.date !== "" && r.transaction !== "Opening Balance");
    }

    populateMainCategories();
    populateAccountDropdowns();
    populateIncomeSubCategories();
    setDefaultDates();
    renderRecentTransactions();
    populateTxFilterCategories();
    renderRecurringSuggestions();
    renderClaimsTracker();

    log("Ready.");
  } catch (err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

// ── Read helpers ──────────────────────────────────────────────────────────────
function readCategories(sheet) {
  const map = {};
  // Main cats in col C (index 2), sub cats in col D (index 3) — rows 2..13 (Bills) and rows 2..13 (Monthly)
  // Actually read from Budget Setup: cols A (sub), D (main) grouping
  // We'll parse col D = main cat, col E = sub cat (if that's your setup)
  // ── fallback: derive from transaction sheet header structure ──
  // Use cols A–B for Bills, F–G for Monthly; main categories are the section headers
  const billsRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "A2:A13", blankrows: false });
  const monthlyRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "F2:F13", blankrows: false });

  map["Bills"] = billsRows.map(r => String(r[0] ?? "").trim()).filter(Boolean);
  map["Monthly Expenses"] = monthlyRows.map(r => String(r[0] ?? "").trim()).filter(Boolean);
  return map;
}

function readIncomeSubCategories(sheet) {
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => String(row[16] ?? "").trim())   // column Q = index 16
    .filter(v => v !== "");
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function setDefaultDates() {
  const today = new Date();
  const val = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");
  ["txDate","inDate","trDate","ccPayDate"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = val;
  });
}

function populateMainCategories() {
  const sel = document.getElementById("txMainCategory");
  if (!sel) return;
  sel.innerHTML = `<option value="">-- Select --</option>`;
  Object.keys(categories).forEach(mainCat => {
    const opt = document.createElement("option");
    opt.value = mainCat;
    opt.textContent = mainCat;
    sel.appendChild(opt);
  });
}

function onMainCategoryChange() {
  const mainCat = document.getElementById("txMainCategory").value;
  const subSel = document.getElementById("txSubCategory");
  subSel.innerHTML = `<option value="">-- Select --</option>`;
  (categories[mainCat] || []).forEach(sub => {
    const opt = document.createElement("option");
    opt.value = sub;
    opt.textContent = sub;
    subSel.appendChild(opt);
  });
}

function populateAccountDropdowns() {
  const savingsAccounts  = accounts.filter(a => a.type !== "Credit Card");
  const ccAccounts       = accounts.filter(a => a.type === "Credit Card");
  const allAccounts      = accounts;

  fillSelect("txAccount",        allAccounts);
  fillSelect("txClaimAccount",   allAccounts, "Select credited account");
  fillSelect("inAccount",        savingsAccounts);
  fillSelect("trFromAccount",    allAccounts);
  fillSelect("trToAccount",      allAccounts);
  fillSelect("ccPayFromAccount", savingsAccounts);
  fillSelect("ccPayToAccount",   ccAccounts);
}

function fillSelect(id, list, placeholder = "") {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = "";
  if (placeholder) {
    const placeholderOption = document.createElement("option");
    placeholderOption.value = "";
    placeholderOption.textContent = placeholder;
    sel.appendChild(placeholderOption);
  }
  list.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a.name;
    opt.textContent = a.name + (a.type === "Credit Card" ? " (CC)" : "");
    sel.appendChild(opt);
  });
}

function populateIncomeSubCategories() {
  const sel = document.getElementById("inSubCategory");
  if (!sel) return;
  sel.innerHTML = `<option value="">-- Select --</option>`;
  incomeSubCategories.forEach(sub => {
    const opt = document.createElement("option");
    opt.value = sub;
    opt.textContent = sub;
    sel.appendChild(opt);
  });
}

// ── Recurring suggestions ────────────────────────────────────────────────────
function renderRecurringSuggestions() {
  const wrapper = document.getElementById("recurringSuggestions");
  const list = document.getElementById("recurringSuggestionList");
  const badge = document.getElementById("recurringSuggestionsBadge");
  if (!wrapper || !list) return;

  recurringSuggestions = buildRecurringSuggestions(recentRows).slice(0, 6);

  if (recurringSuggestions.length === 0) {
    wrapper.style.display = "none";
    list.innerHTML = "";
    if (badge) badge.textContent = "";
    return;
  }

  wrapper.style.display = "";
  if (badge) badge.textContent = recurringSuggestions.length + " found";
  list.innerHTML = recurringSuggestions.map((suggestion, index) => `
    <button type="button" class="recurring-suggestion" onclick="applyRecurringSuggestion(${index})">
      <span class="rs-main">
        <span class="rs-desc">${escapeHtml(suggestion.transaction)}</span>
        <span class="rs-amount">${formatAmount(suggestion.amount)}</span>
      </span>
      <span class="rs-meta">${escapeHtml(suggestion.subCat)} / ${escapeHtml(suggestion.account)} / ${escapeHtml(suggestion.frequencyLabel)}</span>
      <span class="rs-meta">Last ${escapeHtml(suggestion.lastDisplayDate)}</span>
    </button>
  `).join("");
}

function buildRecurringSuggestions(rows) {
  const candidates = rows
    .map(rowToRecurringCandidate)
    .filter(Boolean);

  if (candidates.length < 3) return [];

  const latestTime = candidates.reduce((latest, item) => Math.max(latest, item.time), 0);
  const latestByBaseKey = new Map();
  const groups = new Map();

  candidates.forEach(item => {
    latestByBaseKey.set(item.baseKey, Math.max(latestByBaseKey.get(item.baseKey) ?? 0, item.time));
    if (!groups.has(item.key)) groups.set(item.key, []);
    groups.get(item.key).push(item);
  });

  return [...groups.values()]
    .map(group => groupToRecurringSuggestion(group, latestTime, latestByBaseKey))
    .filter(Boolean)
    .sort((a, b) =>
      b.lastTime - a.lastTime ||
      b.count - a.count ||
      a.transaction.localeCompare(b.transaction)
    );
}

function rowToRecurringCandidate(row) {
  const dateParts = parseInputDateParts(rawDateToInputValue(row.date));
  if (!dateParts) return null;

  const transaction = String(row.transaction ?? "").trim();
  const mainCat = String(row.mainCat ?? "").trim();
  const subCat = String(row.subCat ?? "").trim();
  const account = String(row.account ?? "").trim();
  const amount = parseMoney(row.amount);
  const descriptionKey = normalizeRecurringText(transaction);

  if (!transaction || descriptionKey.length < 3 || amount <= 0 || !mainCat || !subCat || !account) return null;
  if (!Object.prototype.hasOwnProperty.call(categories, mainCat)) return null;
  if (["income", "transfer", "saving goals", "savings goal"].includes(normalizeRecurringText(mainCat))) return null;

  const amountCents = Math.round(amount * 100);
  const claimable = row.claimable === "Yes" ? "Yes" : "";
  const claimAmountCents = claimable ? Math.round(getClaimAmount(row) * 100) : 0;
  const claimAccount = claimable ? String(row.claimAccount ?? "").trim() : "";
  const baseKeyParts = [
    descriptionKey,
    normalizeRecurringText(mainCat),
    normalizeRecurringText(subCat),
    normalizeRecurringText(account),
  ];
  const baseKey = baseKeyParts.join("|");
  const key = [
    ...baseKeyParts,
    amountCents,
    claimable,
    claimAmountCents,
    normalizeRecurringText(claimAccount),
  ].join("|");

  return {
    row,
    key,
    baseKey,
    time: dateParts.time,
    y: dateParts.y,
    m: dateParts.m,
    d: dateParts.d,
    dateKey: dateParts.inputValue,
    monthKey: dateParts.y + "-" + String(dateParts.m).padStart(2, "0"),
    transaction,
    amount,
    amountCents,
    mainCat,
    subCat,
    account,
    claimable,
    claimAmount: claimAmountCents / 100,
    claimAccount,
  };
}

function groupToRecurringSuggestion(group, latestTime, latestByBaseKey) {
  const sorted = [...group].sort((a, b) => a.time - b.time);
  const latest = sorted[sorted.length - 1];
  const baseLatestTime = latestByBaseKey.get(latest.baseKey) ?? latest.time;

  if (latest.time < baseLatestTime) return null;

  const cadence = getRecurringCadence(sorted, latestTime);
  if (!cadence) return null;

  return {
    transaction: latest.transaction,
    amount: latest.amount,
    mainCat: latest.mainCat,
    subCat: latest.subCat,
    account: latest.account,
    claimable: latest.claimable,
    claimAmount: latest.claimAmount,
    claimAccount: latest.claimAccount,
    count: sorted.length,
    frequency: cadence.frequency,
    frequencyLabel: cadence.label,
    lastTime: latest.time,
    lastDisplayDate: formatDisplayDate(latest.row.date),
  };
}

function getRecurringCadence(items, latestTime) {
  if (hasStrictDailyCadence(items, latestTime)) {
    return { frequency: "daily", label: items.length + " daily matches" };
  }
  if (hasStrictMonthlyCadence(items, latestTime)) {
    return { frequency: "monthly", label: items.length + " monthly matches" };
  }
  return null;
}

function hasStrictDailyCadence(items, latestTime) {
  const dateKeys = new Set(items.map(item => item.dateKey));
  if (items.length < 4 || dateKeys.size !== items.length) return false;
  if (daysBetweenTimes(items[items.length - 1].time, latestTime) > 5) return false;

  for (let i = 1; i < items.length; i++) {
    if (daysBetweenTimes(items[i - 1].time, items[i].time) !== 1) return false;
  }
  return true;
}

function hasStrictMonthlyCadence(items, latestTime) {
  const monthKeys = new Set(items.map(item => item.monthKey));
  if (items.length < 3 || monthKeys.size !== items.length) return false;
  if (daysBetweenTimes(items[items.length - 1].time, latestTime) > 75) return false;

  for (let i = 1; i < items.length; i++) {
    const prev = items[i - 1];
    const current = items[i];
    if (monthsBetween(prev, current) !== 1) return false;
    if (daysBetweenTimes(prev.time, current.time) < 24 || daysBetweenTimes(prev.time, current.time) > 38) return false;
    if (!sameRecurringDay(prev, current)) return false;
  }

  const days = items.map(item => item.d);
  const daySpread = Math.max(...days) - Math.min(...days);
  return daySpread <= 7 || items.every(isNearMonthEnd);
}

function applyRecurringSuggestion(index) {
  const suggestion = recurringSuggestions[index];
  if (!suggestion) return;

  const dateInput = document.getElementById("txDate");
  const selectedDate = dateInput?.value || "";

  setInputValue("txTransaction", suggestion.transaction);
  setInputValue("txAmount", suggestion.amount.toFixed(2));
  setSelectValue("txMainCategory", suggestion.mainCat);
  onMainCategoryChange();
  setSelectValue("txSubCategory", suggestion.subCat);
  setSelectValue("txAccount", suggestion.account);

  if (suggestion.claimable === "Yes") {
    if (!isClaimable) toggleClaimable();
    setInputValue("txClaimAmount", suggestion.claimAmount > 0 ? suggestion.claimAmount.toFixed(2) : "");
    setSelectValue("txClaimAccount", suggestion.claimAccount);
    syncClaimAmountCap();
  } else {
    resetClaimableUi();
  }

  if (dateInput && selectedDate) dateInput.value = selectedDate;

  document.querySelectorAll(".recurring-suggestion").forEach((button, buttonIndex) => {
    button.classList.toggle("active", buttonIndex === index);
  });
}

function setInputValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value ?? "";
}

function setSelectValue(id, value) {
  const sel = document.getElementById(id);
  if (!sel) return;
  const optionValue = String(value ?? "").trim();
  if (!optionValue) {
    sel.value = "";
    return;
  }
  if (![...sel.options].some(option => option.value === optionValue)) {
    const option = document.createElement("option");
    option.value = optionValue;
    option.textContent = optionValue;
    sel.appendChild(option);
  }
  sel.value = optionValue;
}

// ── AI prefill helper ─────────────────────────────────────────────────────────
function applyAiPrefill() {
  const input = document.getElementById("aiPrefillText");
  const raw = String(input?.value ?? "").trim();
  if (!raw) {
    setAiPrefillStatus("Enter rough expense details first.", "warn");
    return;
  }

  const suggestion = buildAiPrefillSuggestion(raw);
  setMode("transaction");

  if (suggestion.date) setInputValue("txDate", suggestion.date);
  if (suggestion.amount > 0) setInputValue("txAmount", suggestion.amount.toFixed(2));
  if (suggestion.description) setInputValue("txTransaction", suggestion.description);
  if (suggestion.mainCat) {
    setSelectValue("txMainCategory", suggestion.mainCat);
    onMainCategoryChange();
  }
  if (suggestion.subCat) setSelectValue("txSubCategory", suggestion.subCat);
  if (suggestion.account) setSelectValue("txAccount", suggestion.account);

  if (suggestion.claimable) {
    if (!isClaimable) toggleClaimable();
    setInputValue("txClaimAmount", suggestion.claimAmount > 0 ? suggestion.claimAmount.toFixed(2) : "");
    setSelectValue("txClaimAccount", suggestion.claimAccount);
  } else {
    resetClaimableUi();
  }
  syncClaimAmountCap();

  setAiPrefillStatus(buildAiPrefillStatus(suggestion), suggestion.missing.length ? "warn" : "good");
  setInputValue("aiPrefillText", "");
}

function clearAiPrefill() {
  setInputValue("aiPrefillText", "");
  setAiPrefillStatus("");
}

function setAiPrefillStatus(message, tone = "") {
  const el = document.getElementById("aiPrefillStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "ai-prefill-status" + (tone ? " " + tone : "");
}

function buildAiPrefillSuggestion(raw) {
  const amountInfo = parseSmartAmount(raw);
  const dateInfo = parseSmartDate(raw);
  const amount = amountInfo?.value || 0;
  const roughDescription = buildSmartDescription(raw, amountInfo, dateInfo);
  const historyMatch = findSmartHistoryMatch(raw, roughDescription, amount);
  const category = inferSmartCategory(raw, roughDescription, historyMatch);
  const account = inferSmartAccount(raw, historyMatch);
  const claim = inferSmartClaim(raw, amount);
  const description = inferSmartDescription(raw, amountInfo, dateInfo, historyMatch, category, account, claim);
  const date = dateInfo?.value || document.getElementById("txDate")?.value || todayInputValue();

  const missing = [];
  if (!date) missing.push("date");
  if (!(amount > 0)) missing.push("amount");
  if (!category.mainCat) missing.push("main category");
  if (!category.subCat) missing.push("sub category");
  if (!description) missing.push("description");
  if (!account) missing.push("account");

  return {
    raw,
    date,
    amount,
    description,
    mainCat: category.mainCat,
    subCat: category.subCat,
    account,
    claimable: claim.claimable,
    claimAmount: claim.amount,
    claimAccount: claim.account,
    historyMatch,
    categorySource: category.source,
    missing
  };
}

function buildAiPrefillStatus(suggestion) {
  const filled = [];
  if (suggestion.date) filled.push("date");
  if (suggestion.amount > 0) filled.push("amount");
  if (suggestion.description) filled.push("description");
  if (suggestion.mainCat && suggestion.subCat) filled.push("category");
  if (suggestion.account) filled.push("account");
  if (suggestion.claimable) filled.push("claim");

  let message = filled.length ? "Prefilled " + filled.join(", ") + "." : "Could not prefill from that text.";
  if (suggestion.historyMatch) {
    message += " Matched recent " + suggestion.historyMatch.transaction + ".";
  } else if (suggestion.categorySource === "fallback") {
    message += " Category is a fallback, so review it before saving.";
  }
  if (suggestion.missing.length) {
    message += " Still needs " + suggestion.missing.join(", ") + ".";
  }
  return message;
}

function parseSmartAmount(raw) {
  const candidates = [];
  const text = String(raw ?? "");
  const addCandidate = (match, valueIndex, score) => {
    const rawValue = match[valueIndex];
    const value = parseMoney(rawValue);
    if (!(value > 0)) return;
    candidates.push({
      value,
      rawMatch: match[0],
      index: match.index ?? 0,
      score
    });
  };

  [
    /(?:sgd|s\$|\$)\s*([0-9][0-9,]*(?:\.\d{1,2})?)/ig,
    /([0-9][0-9,]*(?:\.\d{1,2})?)\s*(?:sgd|s\$|\$)/ig,
    /(?:amount|amt|spent|paid|cost|costs|for)\s*(?:sgd|s\$|\$)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)/ig
  ].forEach(pattern => {
    for (const match of text.matchAll(pattern)) addCandidate(match, 1, 90);
  });

  for (const match of text.matchAll(/\b([0-9][0-9,]*\.\d{1,2})\b/g)) {
    addCandidate(match, 1, 70);
  }

  for (const match of text.matchAll(/\b([0-9][0-9,]*)\b/g)) {
    const value = parseMoney(match[1]);
    if (!(value > 0) || value > 100000 || isSmartDateNumber(text, match.index ?? 0, match[0].length)) continue;
    candidates.push({ value, rawMatch: match[0], index: match.index ?? 0, score: 35 });
  }

  candidates.sort((a, b) => b.score - a.score || b.rawMatch.length - a.rawMatch.length || a.index - b.index);
  return candidates[0] || null;
}

function isSmartDateNumber(text, index, length) {
  const value = Number(text.slice(index, index + length).replace(/,/g, ""));
  if (value >= 1900 && value <= 2100) return true;
  const before = text.slice(Math.max(0, index - 12), index).toLowerCase();
  const after = text.slice(index + length, index + length + 12).toLowerCase();
  if (/[\/-]\s*$/.test(before) || /^\s*[\/-]/.test(after)) return true;
  if (value >= 1 && value <= 31) {
    const near = before + " " + after;
    if (Object.keys(SMART_MONTHS).some(month => new RegExp("\\b" + month + "\\b", "i").test(near))) return true;
  }
  return false;
}

function parseSmartDate(raw) {
  const text = String(raw ?? "");
  const lower = text.toLowerCase();
  const today = new Date();

  if (/\btoday\b/.test(lower)) return { value: inputDateFromDate(today), rawMatch: "today" };
  if (/\byesterday\b/.test(lower)) return { value: inputDateFromDate(addDays(today, -1)), rawMatch: "yesterday" };
  if (/\btomorrow\b/.test(lower)) return { value: inputDateFromDate(addDays(today, 1)), rawMatch: "tomorrow" };

  let match = text.match(/\b(\d{4})-(\d{1,2})-(\d{1,2})\b/);
  if (match) return { value: inputDateFromParts(Number(match[1]), Number(match[2]), Number(match[3])), rawMatch: match[0] };

  match = text.match(/\b(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?\b/);
  if (match) {
    const year = match[3] ? normalizeSmartYear(match[3]) : today.getFullYear();
    return { value: inputDateFromParts(year, Number(match[2]), Number(match[1])), rawMatch: match[0] };
  }

  const monthNames = Object.keys(SMART_MONTHS).join("|");
  match = text.match(new RegExp("\\b(\\d{1,2})\\s+(" + monthNames + ")(?:\\s+(\\d{2,4}))?\\b", "i"));
  if (match) {
    const year = match[3] ? normalizeSmartYear(match[3]) : today.getFullYear();
    return { value: inputDateFromParts(year, SMART_MONTHS[match[2].toLowerCase()], Number(match[1])), rawMatch: match[0] };
  }

  match = text.match(new RegExp("\\b(" + monthNames + ")\\s+(\\d{1,2})(?:\\s+(\\d{2,4}))?\\b", "i"));
  if (match) {
    const year = match[3] ? normalizeSmartYear(match[3]) : today.getFullYear();
    return { value: inputDateFromParts(year, SMART_MONTHS[match[1].toLowerCase()], Number(match[2])), rawMatch: match[0] };
  }

  match = lower.match(/\blast\s+(sun(?:day)?|mon(?:day)?|tue(?:s|sday)?|wed(?:nesday)?|thu(?:r|rs|rsday|rday)?|fri(?:day)?|sat(?:urday)?)\b/);
  if (match) {
    const target = SMART_WEEKDAYS[match[1]];
    if (target !== undefined) {
      const daysBack = ((today.getDay() - target + 7) % 7) || 7;
      return { value: inputDateFromDate(addDays(today, -daysBack)), rawMatch: match[0] };
    }
  }

  return null;
}

function findSmartHistoryMatch(raw, description, amount) {
  const inputNorm = normalizeRecurringText(raw + " " + description);
  const inputTokens = new Set(tokenizeSmartText(inputNorm));
  const rows = [...recentRows]
    .filter(row => row.mainCat !== "Income" && row.mainCat !== "Transfer")
    .filter(row => row.transaction && row.mainCat && row.subCat)
    .map(row => {
      const txNorm = normalizeRecurringText(row.transaction);
      const rowTokens = tokenizeSmartText(row.transaction + " " + row.subCat + " " + row.account);
      let score = 0;

      if (txNorm && inputNorm.includes(txNorm)) score += 22;
      if (txNorm && normalizeRecurringText(description).includes(txNorm)) score += 12;
      rowTokens.forEach(token => {
        if (inputTokens.has(token)) score += txNorm.includes(token) ? 4 : 2;
      });
      if (amount > 0) {
        const rowAmount = parseMoney(row.amount);
        if (Math.abs(rowAmount - amount) < 0.01) score += 8;
        else if (Math.round(rowAmount) === Math.round(amount)) score += 3;
      }
      if (row.account && inputNorm.includes(normalizeRecurringText(row.account))) score += 6;
      if (row.subCat && inputNorm.includes(normalizeRecurringText(row.subCat))) score += 5;

      return { ...row, _smartScore: score };
    })
    .filter(row => row._smartScore >= 6)
    .sort((a, b) => b._smartScore - a._smartScore || rawDateToInputValue(b.date).localeCompare(rawDateToInputValue(a.date)));

  return rows[0] || null;
}

function inferSmartCategory(raw, description, historyMatch) {
  if (historyMatch && isKnownCategory(historyMatch.mainCat, historyMatch.subCat)) {
    return { mainCat: historyMatch.mainCat, subCat: historyMatch.subCat, source: "history" };
  }

  const inputNorm = normalizeRecurringText(raw + " " + description);
  const inputTokens = new Set(tokenizeSmartText(inputNorm));
  const pairs = getSmartCategoryPairs();
  const scored = pairs.map(pair => {
    const categoryNorm = normalizeRecurringText(pair.mainCat + " " + pair.subCat);
    const categoryTokens = tokenizeSmartText(categoryNorm);
    let score = 0;

    if (inputNorm.includes(normalizeRecurringText(pair.subCat))) score += 22;
    if (inputNorm.includes(normalizeRecurringText(pair.mainCat))) score += 5;
    categoryTokens.forEach(token => { if (inputTokens.has(token)) score += 4; });

    SMART_CATEGORY_RULES.forEach(rule => {
      const keywordHit = rule.keywords.some(keyword => inputNorm.includes(normalizeRecurringText(keyword)));
      if (!keywordHit) return;
      const hintHit = rule.hints.some(hint => categoryNorm.includes(normalizeRecurringText(hint)));
      if (hintHit) score += 18;
    });

    return { ...pair, score };
  }).sort((a, b) => b.score - a.score || a.subCat.localeCompare(b.subCat));

  if (scored[0] && scored[0].score >= 4) {
    return { mainCat: scored[0].mainCat, subCat: scored[0].subCat, source: "text" };
  }

  const fallback = pairs.find(pair => normalizeRecurringText(pair.subCat).includes("other"))
    || pairs.find(pair => pair.mainCat === "Monthly Expenses")
    || pairs[0];
  return fallback
    ? { mainCat: fallback.mainCat, subCat: fallback.subCat, source: "fallback" }
    : { mainCat: "", subCat: "", source: "" };
}

function inferSmartAccount(raw, historyMatch) {
  const inputNorm = normalizeRecurringText(raw);
  const scored = accounts.map(account => {
    const accountNorm = normalizeRecurringText(account.name);
    const accountTokens = tokenizeSmartText(account.name);
    let score = 0;
    if (accountNorm && inputNorm.includes(accountNorm)) score += 30;
    accountTokens.forEach(token => { if (inputNorm.includes(token)) score += 8; });
    if (account.type === "Credit Card" && /\b(cc|credit card|visa|mastercard|amex)\b/.test(inputNorm)) score += 3;
    return { account, score };
  }).sort((a, b) => b.score - a.score || a.account.name.localeCompare(b.account.name));

  if (scored[0]?.score >= 8) return scored[0].account.name;
  if (historyMatch?.account) return historyMatch.account;
  return document.getElementById("txAccount")?.value || "";
}

function inferSmartClaim(raw, amount) {
  const text = String(raw ?? "");
  const inputNorm = normalizeRecurringText(text);
  const claimable = /\b(claim|claimable|reimburse|reimbursed|reimbursement)\b/.test(inputNorm);
  if (!claimable) return { claimable: false, amount: 0, account: "" };

  let claimAmount = 0;
  const match = text.match(/(?:claim|claimable|reimburse|reimbursed|reimbursement)\s*(?:amount)?\s*(?:sgd|s\$|\$)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)/i);
  if (match) claimAmount = parseMoney(match[1]);
  else if (/\bhalf\b/.test(inputNorm) && amount > 0) claimAmount = amount / 2;
  else claimAmount = amount;

  const accountMatch = text.match(/(?:claim|reimburse(?:d|ment)?|paid)\s+(?:to|into)\s+(.+)$/i);
  const claimAccount = accountMatch ? inferSmartAccount(accountMatch[1], null) : "";
  return { claimable: true, amount: claimAmount, account: claimAccount };
}

function inferSmartDescription(raw, amountInfo, dateInfo, historyMatch, category, account, claim) {
  let description = buildSmartDescription(raw, amountInfo, dateInfo);
  [account, category?.mainCat, category?.subCat, claim?.account].filter(Boolean).forEach(value => {
    description = removeSmartPhrase(description, value);
  });
  if (account) {
    tokenizeSmartText(account).forEach(token => {
      description = removeSmartWord(description, token);
    });
  }
  description = description
    .replace(/\b(claim|claimable|reimburse|reimbursed|reimbursement|amount|amt|spent|paid|pay|using|with|via|from|on|account|card|sgd|s\$|full|half)\b/ig, " ")
    .replace(/\s+/g, " ")
    .replace(/^[,.;:-]+|[,.;:-]+$/g, "")
    .replace(/^\s*(to|at|for|by|with|via|on|from)\b\s*/i, "")
    .replace(/\s*\b(to|at|for|by|with|via|on|from)\b\s*$/i, "")
    .trim();

  if (description.length >= 3) return description;
  if (historyMatch?.transaction) return historyMatch.transaction;
  return String(raw ?? "").trim().slice(0, 80);
}

function buildSmartDescription(raw, amountInfo, dateInfo) {
  let description = String(raw ?? "");
  [amountInfo?.rawMatch, dateInfo?.rawMatch].filter(Boolean).forEach(value => {
    description = removeSmartPhrase(description, value);
  });
  description = description
    .replace(/\b(today|yesterday|tomorrow)\b/ig, " ")
    .replace(/\blast\s+(sun(?:day)?|mon(?:day)?|tue(?:s|sday)?|wed(?:nesday)?|thu(?:r|rs|rsday|rday)?|fri(?:day)?|sat(?:urday)?)\b/ig, " ")
    .replace(/\b\d{1,2}[\/-]\d{1,2}(?:[\/-]\d{2,4})?\b/g, " ")
    .replace(/\b\d{4}-\d{1,2}-\d{1,2}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return description;
}

function removeSmartPhrase(value, phrase) {
  const escaped = String(phrase ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  if (!escaped) return value;
  return String(value ?? "").replace(new RegExp(escaped, "ig"), " ");
}

function removeSmartWord(value, word) {
  const escaped = String(word ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  if (!escaped) return value;
  return String(value ?? "").replace(new RegExp("\\b" + escaped + "\\b", "ig"), " ");
}

function tokenizeSmartText(value) {
  return normalizeRecurringText(value)
    .split(" ")
    .filter(token => token.length > 1 && !SMART_STOP_WORDS.has(token));
}

function getSmartCategoryPairs() {
  return Object.entries(categories).flatMap(([mainCat, subs]) =>
    (subs || []).map(subCat => ({ mainCat, subCat }))
  );
}

function isKnownCategory(mainCat, subCat) {
  return (categories[mainCat] || []).includes(subCat);
}

function todayInputValue() {
  return inputDateFromDate(new Date());
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function inputDateFromDate(date) {
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}

function inputDateFromParts(year, month, day) {
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return "";
  return inputDateFromDate(date);
}

function normalizeSmartYear(yearText) {
  const year = Number(yearText);
  return year < 100 ? 2000 + year : year;
}

function normalizeRecurringText(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function parseInputDateParts(inputValue) {
  const match = String(inputValue ?? "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const y = Number(match[1]);
  const m = Number(match[2]);
  const d = Number(match[3]);
  const time = Date.UTC(y, m - 1, d);
  if (Number.isNaN(time)) return null;
  return { y, m, d, time, inputValue: match[0] };
}

function monthsBetween(a, b) {
  return (b.y - a.y) * 12 + (b.m - a.m);
}

function daysBetweenTimes(startTime, endTime) {
  return Math.round((endTime - startTime) / 86400000);
}

function sameRecurringDay(a, b) {
  return Math.abs(a.d - b.d) <= 7 || (isNearMonthEnd(a) && isNearMonthEnd(b));
}

function isNearMonthEnd(item) {
  return item.d >= daysInMonth(item.y, item.m) - 2;
}

function daysInMonth(year, month) {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

// ── Mode switching ─────────────────────────────────────────────────────────────
function setMode(mode) {
  currentMode = mode;
  ["transaction","income","transfer","ccpay"].forEach(m => {
    const formId = m === "transaction" ? "transactionForm"
                 : m === "income"      ? "incomeForm"
                 : m === "transfer"    ? "transferForm"
                 :                       "ccPayForm";
    const btn = document.getElementById("btn" + m.charAt(0).toUpperCase() + m.slice(1));
    const form = document.getElementById(formId);
    if (form) form.style.display = m === mode ? "block" : "none";
    if (btn)  btn.classList.toggle("mode-active", m === mode);
  });
}

// ── Claimable toggle (expense form) ──────────────────────────────────────────
function toggleClaimable() {
  isClaimable = !isClaimable;
  const btn = document.getElementById("claimableToggle");
  if (btn) {
    btn.textContent  = isClaimable ? "🔖 Claimable: ON" : "🔖 Mark as Claimable";
    btn.style.background = isClaimable ? "#e67e22" : "";
    btn.style.color      = isClaimable ? "white"   : "";
    btn.style.borderColor = isClaimable ? "#e67e22" : "";
  }
  const hint = document.getElementById("claimableHint");
  if (hint) hint.style.display = isClaimable ? "block" : "none";
  const details = document.getElementById("claimDetails");
  if (details) details.classList.toggle("open", isClaimable);
  if (isClaimable) syncClaimAmountCap();
}

function setClaimAmountPreset(mode) {
  const amount = parseFloat(document.getElementById("txAmount")?.value);
  const claimInput = document.getElementById("txClaimAmount");
  if (!claimInput || isNaN(amount) || amount <= 0) return;
  claimInput.value = (mode === "half" ? amount / 2 : amount).toFixed(2);
}

function syncClaimAmountCap() {
  const amount = parseFloat(document.getElementById("txAmount")?.value);
  const claimInput = document.getElementById("txClaimAmount");
  if (!claimInput) return;
  if (!isNaN(amount) && amount > 0) claimInput.max = String(amount);
  const claimAmount = parseFloat(claimInput.value);
  if (!isNaN(amount) && !isNaN(claimAmount) && claimAmount > amount) {
    claimInput.value = amount.toFixed(2);
  }
}

function resetClaimableUi() {
  isClaimable = false;
  const btn = document.getElementById("claimableToggle");
  if (btn) { btn.textContent = "🔖 Mark as Claimable"; btn.style.background = ""; btn.style.color = ""; btn.style.borderColor = ""; }
  const hint = document.getElementById("claimableHint");
  if (hint) hint.style.display = "none";
  const details = document.getElementById("claimDetails");
  if (details) details.classList.remove("open");
  const claimInput = document.getElementById("txClaimAmount");
  if (claimInput) claimInput.value = "";
  const claimAccount = document.getElementById("txClaimAccount");
  if (claimAccount) claimAccount.value = "";
}

// ── Date formatting helper ────────────────────────────────────────────────────
function formatDateForExcel(dateInputValue) {
  // Write as YYYY-MM-DD. Slash-based formats (DD/MM/YYYY) are ambiguous —
  // Excel / Graph API interprets slashes as MM/DD/YYYY (American), causing
  // e.g. 10 June to be stored as 6 October. ISO format is unambiguous.
  return dateInputValue; // input[type="date"] already yields YYYY-MM-DD
}

// ── Save transaction ──────────────────────────────────────────────────────────
async function saveTransaction() {
  const dateVal  = document.getElementById("txDate").value;
  const mainCat  = document.getElementById("txMainCategory").value;
  const subCat   = document.getElementById("txSubCategory").value;
  const txDesc   = document.getElementById("txTransaction").value.trim();
  const amount   = parseFloat(document.getElementById("txAmount").value);
  const account  = document.getElementById("txAccount").value;

  if (!dateVal || !mainCat || !subCat || !txDesc || isNaN(amount) || !account) {
    alert("Please fill in all fields."); return;
  }

  const dateStr     = formatDateForExcel(dateVal);
  const claimFlag   = isClaimable ? "Yes" : "";
  const claimStatus = isClaimable ? "Pending" : "";
  let claimAmount = "";
  let claimAccount = "";

  if (isClaimable) {
    const rawClaimAmount = parseFloat(document.getElementById("txClaimAmount")?.value);
    claimAccount = document.getElementById("txClaimAccount")?.value || "";
    claimAmount = isNaN(rawClaimAmount) ? amount : rawClaimAmount;
    if (claimAmount <= 0) { alert("Claim amount must be more than $0."); return; }
    if (claimAmount > amount) { alert("Claim amount cannot be more than the expense amount."); return; }
    claimAmount = Number(claimAmount.toFixed(2));
  }

  // Columns: A=Date, B=Transaction, C=Amount, D=MainCategory, E=SubCategory, F=Account, G=Claimable, H=ClaimStatus, I=ClaimAmount, J=ClaimAccount
  const row = [dateStr, txDesc, amount, mainCat, subCat, account, claimFlag, claimStatus, claimAmount, claimAccount];

  try {
    log("Saving transaction...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await ensureClaimHeaders();
    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:J${nextRow}`, [row]);

    resetClaimableUi();

    document.getElementById("txTransaction").value = "";
    document.getElementById("txAmount").value = "";
    document.getElementById("transactionSuccess").style.display = "block";
    setTimeout(() => document.getElementById("transactionSuccess").style.display = "none", 4000);

    log("Saved. Reloading recent transactions...");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save income ───────────────────────────────────────────────────────────────
async function saveIncome() {
  const dateVal = document.getElementById("inDate").value;
  const txDesc  = document.getElementById("inTransaction").value.trim();
  const subCat  = document.getElementById("inSubCategory").value;
  const amount  = parseFloat(document.getElementById("inAmount").value);
  const account = document.getElementById("inAccount").value;

  if (!dateVal || !txDesc || isNaN(amount) || !account) {
    alert("Please fill in all fields."); return;
  }

  const dateStr = formatDateForExcel(dateVal);
  // Income rows: MainCategory = "Income", Claimable = ""
  const row = [dateStr, txDesc, amount, "Income", subCat || "", account, "", ""];

  try {
    log("Saving income...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow}`, [row]);

    document.getElementById("inTransaction").value = "";
    document.getElementById("inAmount").value = "";
    document.getElementById("incomeSuccess").style.display = "block";
    setTimeout(() => document.getElementById("incomeSuccess").style.display = "none", 4000);

    log("Income saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save transfer ─────────────────────────────────────────────────────────────
async function saveTransfer() {
  const dateVal     = document.getElementById("trDate").value;
  const fromAccount = document.getElementById("trFromAccount").value;
  const toAccount   = document.getElementById("trToAccount").value;
  const amount      = parseFloat(document.getElementById("trAmount").value);
  const note        = document.getElementById("trNote").value.trim() || "Transfer";

  if (!dateVal || !fromAccount || !toAccount || isNaN(amount)) {
    alert("Please fill in all fields."); return;
  }
  if (fromAccount === toAccount) { alert("From and To accounts must be different."); return; }

  const dateStr = formatDateForExcel(dateVal);
  const rows = [
    [dateStr, note, amount, "Transfer", "Transfer Out", fromAccount, "", ""],
    [dateStr, note, amount, "Transfer", "Transfer In",  toAccount,   "", ""],
  ];

  try {
    log("Saving transfer...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow + 1}`, rows);

    document.getElementById("trAmount").value = "";
    document.getElementById("trNote").value = "";
    document.getElementById("transferSuccess").style.display = "block";
    setTimeout(() => document.getElementById("transferSuccess").style.display = "none", 4000);

    log("Transfer saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save CC payment ───────────────────────────────────────────────────────────
async function saveCcPayment() {
  const dateVal     = document.getElementById("ccPayDate").value;
  const fromAccount = document.getElementById("ccPayFromAccount").value;
  const toAccount   = document.getElementById("ccPayToAccount").value;
  const amount      = parseFloat(document.getElementById("ccPayAmount").value);

  if (!dateVal || !fromAccount || !toAccount || isNaN(amount)) {
    alert("Please fill in all fields."); return;
  }

  const dateStr = formatDateForExcel(dateVal);
  const rows = [
    [dateStr, "CC Payment",  amount, "Transfer", "CC Payment Out", fromAccount, "", ""],
    [dateStr, "CC Payment",  amount, "Transfer", "CC Payment In",  toAccount,   "", ""],
  ];

  try {
    log("Saving CC payment...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow + 1}`, rows);

    document.getElementById("ccPayAmount").value = "";
    document.getElementById("ccPaySuccess").style.display = "block";
    setTimeout(() => document.getElementById("ccPaySuccess").style.display = "none", 4000);

    log("CC payment saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Recent transactions table ─────────────────────────────────────────────────
// Convert any date value (serial number or DD/MM/YYYY string) to YYYY-MM-DD for <input type="date">
function rawDateToInputValue(value) {
  if (!value && value !== 0) return "";
  // Excel serial number (numeric or stringified numeric)
  if (typeof value === "number" || (typeof value === "string" && /^\d{5}$/.test(value.trim()))) {
    try {
      const d = XLSX.SSF.parse_date_code(Number(value));
      return d.y + "-" + String(d.m).padStart(2,"0") + "-" + String(d.d).padStart(2,"0");
    } catch { return ""; }
  }
  const s = String(value).trim();
  // DD/MM/YYYY
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return ddmm[3] + "-" + ddmm[2].padStart(2,"0") + "-" + ddmm[1].padStart(2,"0");
  // YYYY-MM-DD already
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  return "";
}

function renderRecentTransactions() {
  applyTxFilter();
}

function populateTxFilterCategories() {
  const sel = document.getElementById("txFilterMainCat");
  if (!sel) return;
  const current = sel.value;
  sel.innerHTML = `<option value="">All Categories</option>`;
  const allCats = [...new Set(recentRows.map(r => r.mainCat).filter(Boolean))].sort();
  allCats.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    if (cat === current) opt.selected = true;
    sel.appendChild(opt);
  });
}

function applyTxFilter() {
  const text    = (document.getElementById("txFilterText")?.value ?? "").toLowerCase().trim();
  const mainCat = (document.getElementById("txFilterMainCat")?.value ?? "");

  const sorted = [...recentRows]
    .filter(r => rawDateToInputValue(r.date))
    .sort((a, b) => rawDateToInputValue(b.date).localeCompare(rawDateToInputValue(a.date)))
    .slice(0, 50);

  filteredRecentRows = sorted.filter(r => {
    const matchesCat  = !mainCat || r.mainCat === mainCat;
    const matchesText = !text ||
      r.transaction.toLowerCase().includes(text) ||
      r.subCat.toLowerCase().includes(text) ||
      r.account.toLowerCase().includes(text) ||
      r.claimAccount.toLowerCase().includes(text);
    return matchesCat && matchesText;
  });

  renderRecentTransactionsTable();
}

function clearTxFilter() {
  const txt = document.getElementById("txFilterText");
  const sel = document.getElementById("txFilterMainCat");
  if (txt) txt.value = "";
  if (sel) sel.value = "";
  applyTxFilter();
}

function renderRecentTransactionsTable() {
  const container = document.getElementById("txTableBody");
  if (!container) return;

  const badge = document.getElementById("recentTxBadge");
  if (badge) badge.textContent = filteredRecentRows.length + " shown";

  if (filteredRecentRows.length === 0) {
    container.innerHTML = '<div class="tx-empty">No transactions match the filter.</div>';
    return;
  }

  const allMainCats = [...new Set(Object.keys(categories))].sort();
  const extraCats   = ["Income", "Transfer", "Saving Goals"].filter(c => !allMainCats.includes(c));
  const allCats     = [...allMainCats, ...extraCats];
  const allAccounts = accounts.map(a => a.name);

  let html = `
    <div class="tx-row-header">
      <span>Date</span>
      <span>Description</span>
      <span>Amount</span>
      <span>Category</span>
      <span>Account</span>
      <span></span>
    </div>`;

  filteredRecentRows.forEach(row => {
    const idx     = row._rowIndex;
    const dateVal = rawDateToInputValue(row.date);
    const dispDate = formatDisplayDate(row.date);
    const amt     = parseFloat(String(row.amount).replace(/[$,]/g, "")) || 0;
    const amtStr  = amt.toLocaleString("en-SG", { style: "currency", currency: "SGD", minimumFractionDigits: 2 });
    const catLabel = [row.mainCat, row.subCat].filter(Boolean).join(" › ");
    const claimAmt = getClaimAmount(row);
    const isPartialClaim = row.claimable === "Yes" && claimAmt > 0 && claimAmt < amt;
    const claimBadge = row.claimable === "Yes"
      ? `<span style="font-size:10px;background:#fff3e0;color:#e67e22;border:1px solid #f0c080;border-radius:4px;padding:1px 5px;margin-left:4px;" title="Claim amount: ${formatAmount(claimAmt)}">${isPartialClaim ? "🔖 Partial" : "🔖"}</span>`
      : "";

    const catOptions   = allCats.map(c =>
      `<option value="${c}" ${c === row.mainCat ? "selected" : ""}>${c}</option>`).join("");
    const acctOptions  = allAccounts.map(a =>
      `<option value="${a}" ${a === row.account ? "selected" : ""}>${a}</option>`).join("");
    const claimAcctOptions = `<option value="">-- Not set --</option>` + allAccounts.map(a =>
      `<option value="${a}" ${a === row.claimAccount ? "selected" : ""}>${a}</option>`).join("");

    html += `
      <div class="tx-row" id="tx-row-${idx}" onclick="toggleEditPanel(${idx})">
        <span class="tx-row-date">${dispDate}</span>
        <span class="tx-row-desc">${escapeHtml(row.transaction)}${claimBadge}</span>
        <span class="tx-row-amount">${amtStr}</span>
        <span class="tx-row-cat">${escapeHtml(catLabel)}</span>
        <span class="tx-row-account">${escapeHtml(row.account)}</span>
        <span class="tx-row-actions"><span class="tx-row-chevron">▾</span></span>
      </div>
      <div class="tx-edit-panel" id="tx-edit-${idx}">
        <div class="tx-edit-grid">
          <div class="ef"><label>Date</label><input type="date" id="edit-date-${idx}" value="${dateVal}"></div>
          <div class="ef"><label>Description</label><input type="text" id="edit-desc-${idx}" value="${escapeHtml(row.transaction)}"></div>
          <div class="ef"><label>Amount</label><input type="number" step="0.01" id="edit-amt-${idx}" value="${amt || ""}"></div>
          <div class="ef"><label>Account</label><select id="edit-acc-${idx}">${acctOptions}</select></div>
          <div class="ef"><label>Main Category</label><select id="edit-main-${idx}">${catOptions}</select></div>
          <div class="ef"><label>Sub Category</label><input type="text" id="edit-sub-${idx}" value="${escapeHtml(row.subCat)}"></div>
          <div class="ef"><label>Claimable</label>
            <select id="edit-claim-${idx}">
              <option value="" ${row.claimable !== "Yes" ? "selected" : ""}>—</option>
              <option value="Yes" ${row.claimable === "Yes" ? "selected" : ""}>Yes</option>
            </select>
          </div>
          <div class="ef"><label>Claim Status</label>
            <select id="edit-claimst-${idx}">
              <option value="" ${!row.claimStatus ? "selected" : ""}>—</option>
              <option value="Pending" ${row.claimStatus === "Pending" ? "selected" : ""}>Pending</option>
              <option value="Claimed" ${row.claimStatus === "Claimed" ? "selected" : ""}>Claimed</option>
            </select>
          </div>
          <div class="ef"><label>Claim Amount</label><input type="number" step="0.01" min="0" id="edit-claimamt-${idx}" value="${row.claimable === "Yes" ? getClaimAmount(row) : ""}" placeholder="Full amount if blank"></div>
          <div class="ef"><label>Claim Paid To</label><select id="edit-claimacc-${idx}">${claimAcctOptions}</select></div>
        </div>
        <div class="tx-edit-actions">
          <button class="btn-edit-save" onclick="saveRow(${idx}); event.stopPropagation();">💾 Save</button>
          <button class="btn-edit-delete" onclick="deleteRow(${idx}); event.stopPropagation();">🗑 Delete</button>
          <button class="btn-edit-cancel" onclick="toggleEditPanel(${idx}); event.stopPropagation();">Cancel</button>
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

let openEditPanel = null;

function toggleEditPanel(idx) {
  const panel = document.getElementById("tx-edit-" + idx);
  const row   = document.getElementById("tx-row-"  + idx);
  if (!panel) return;

  const isOpen = panel.classList.contains("open");

  // Close previously open panel
  if (openEditPanel !== null && openEditPanel !== idx) {
    const prev = document.getElementById("tx-edit-" + openEditPanel);
    const prevRow = document.getElementById("tx-row-" + openEditPanel);
    if (prev) prev.classList.remove("open");
    if (prevRow) prevRow.classList.remove("expanded");
  }

  panel.classList.toggle("open", !isOpen);
  row.classList.toggle("expanded", !isOpen);
  openEditPanel = isOpen ? null : idx;
}

// ── Save an edited row back to Excel ─────────────────────────────────────────
async function saveRow(excelRowNumber) {
  const dateInput = document.getElementById("edit-date-" + excelRowNumber).value;
  const desc      = document.getElementById("edit-desc-" + excelRowNumber).value.trim();
  const amt       = parseFloat(document.getElementById("edit-amt-" + excelRowNumber).value);
  const mainCat   = document.getElementById("edit-main-" + excelRowNumber).value;
  const subCat    = document.getElementById("edit-sub-" + excelRowNumber).value.trim();
  const account   = document.getElementById("edit-acc-" + excelRowNumber).value;
  const claimable = document.getElementById("edit-claim-" + excelRowNumber).value;
  const claimSt   = document.getElementById("edit-claimst-" + excelRowNumber).value;
  const claimAccount = claimable === "Yes"
    ? document.getElementById("edit-claimacc-" + excelRowNumber).value
    : "";
  let claimAmount = "";

  if (!dateInput) { alert("Please enter a date."); return; }
  if (!desc)      { alert("Please enter a description."); return; }
  if (isNaN(amt)) { alert("Please enter a valid amount."); return; }
  if (claimable === "Yes") {
    const rawClaimAmount = parseFloat(document.getElementById("edit-claimamt-" + excelRowNumber).value);
    claimAmount = isNaN(rawClaimAmount) ? amt : rawClaimAmount;
    if (claimAmount <= 0) { alert("Claim amount must be more than $0."); return; }
    if (claimAmount > amt) { alert("Claim amount cannot be more than the expense amount."); return; }
    if (claimSt === "Claimed" && !claimAccount) { alert("Please select the account where the claim was paid."); return; }
    claimAmount = Number(claimAmount.toFixed(2));
  }

  try {
    log("Saving row " + excelRowNumber + "...");
    await ensureClaimHeaders();
    await writeExcelRange(
      CONFIG.sheetName,
      `A${excelRowNumber}:J${excelRowNumber}`,
      [[dateInput, desc, amt, mainCat, subCat, account, claimable, claimSt, claimAmount, claimAccount]]
    );
    log("Row " + excelRowNumber + " saved.");
    // Flash the row green briefly
    const tr = document.getElementById("tx-row-" + excelRowNumber);
    if (tr) { tr.style.background = "#f0fdf4"; setTimeout(() => { tr.style.background = ""; }, 1500); }
    // Reload so recentRows stay in sync
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to save: " + err.message);
  }
}

// ── Claims tracker section ────────────────────────────────────────────────────
function renderClaimsTracker() {
  const container = document.getElementById("claimsContainer");
  if (!container) return;

  const claimRows = recentRows.filter(r => r.claimable === "Yes");

  const badge = document.getElementById("claimsBadge");
  if (badge) {
    const pending = claimRows.filter(r => r.claimStatus !== "Claimed").length;
    badge.textContent = pending > 0 ? pending + " pending" : "";
    badge.style.display = pending > 0 ? "" : "none";
  }

  if (claimRows.length === 0) {
    container.innerHTML = `<p style="color:#aaa;font-size:14px;">No claimable transactions yet.</p>`;
    return;
  }

  const pending = claimRows.filter(r => r.claimStatus !== "Claimed");
  const claimed = claimRows.filter(r => r.claimStatus === "Claimed");

  const pendingTotal = pending.reduce((s, r) => s + getClaimAmount(r), 0);

  let html = `
    <div style="background:#fff8f0;border:1px solid #f0c080;border-radius:10px;padding:16px;margin-bottom:16px;max-width:680px;">
      <strong style="color:#e67e22;">⏳ Pending Claims</strong>
      <span style="float:right;font-weight:bold;color:#e67e22;">${formatAmount(pendingTotal)} to recover</span>
    </div>`;

  if (pending.length > 0) {
    html += `<table class="data-table" style="max-width:680px;margin-bottom:24px;"><tr>
      <th>Date</th><th>Description</th><th>Claim</th><th>Expense</th><th>Category</th><th>Spent From</th><th>Claim Paid To</th><th></th></tr>`;
    pending.forEach(row => {
      const selectedOptions = accountOptionsHtml(row.claimAccount);
      html += `<tr>
        <td>${formatDisplayDate(row.date)}</td>
        <td>${escapeHtml(row.transaction)}</td>
        <td style="color:#e67e22;font-weight:bold;">${formatAmount(getClaimAmount(row))}</td>
        <td>${formatAmount(row.amount)}</td>
        <td>${escapeHtml(row.subCat)}</td>
        <td>${escapeHtml(row.account)}</td>
        <td><select id="claim-account-${row._rowIndex}" style="min-width:150px;">${selectedOptions}</select></td>
        <td><button onclick="markClaimed(${row._rowIndex})"
          style="background:#27ae60;color:white;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-size:12px;">
          ✓ Mark Claimed</button></td>
      </tr>`;
    });
    html += `</table>`;
  }

  if (claimed.length > 0) {
    html += `<details style="max-width:680px;"><summary style="cursor:pointer;color:#888;font-size:13px;margin-bottom:8px;">
      ✅ ${claimed.length} claimed transaction(s)</summary>
      <table class="data-table"><tr><th>Date</th><th>Description</th><th>Claim</th><th>Expense</th><th>Spent From</th><th>Paid To</th></tr>`;
    claimed.forEach(row => {
      html += `<tr style="opacity:0.6;">
        <td>${formatDisplayDate(row.date)}</td>
        <td>${escapeHtml(row.transaction)}</td>
        <td>${formatAmount(getClaimAmount(row))}</td>
        <td>${formatAmount(row.amount)}</td>
        <td>${escapeHtml(row.account)}</td>
        <td>${escapeHtml(row.claimAccount || "-")}</td>
      </tr>`;
    });
    html += `</table></details>`;
  }

  container.innerHTML = html;
}

function accountOptionsHtml(selectedAccount = "") {
  return `<option value="">Select account</option>` + accounts.map(account => {
    const label = account.name + (account.type === "Credit Card" ? " (CC)" : "");
    return `<option value="${escapeHtml(account.name)}" ${account.name === selectedAccount ? "selected" : ""}>${escapeHtml(label)}</option>`;
  }).join("");
}

// ── Mark a row as Claimed (writes status and claim account) ──────────────────
async function markClaimed(excelRowNumber) {
  const row = recentRows.find(item => item._rowIndex === excelRowNumber);
  const claimAccount = document.getElementById("claim-account-" + excelRowNumber)?.value || row?.claimAccount || "";
  if (!claimAccount) {
    alert("Please select the account where the claim was paid.");
    return;
  }

  try {
    log("Marking row " + excelRowNumber + " as Claimed...");
    await ensureClaimHeaders();
    await writeExcelRange(CONFIG.sheetName, `H${excelRowNumber}:H${excelRowNumber}`, [["Claimed"]]);
    await writeExcelRange(CONFIG.sheetName, `J${excelRowNumber}:J${excelRowNumber}`, [[claimAccount]]);
    log("Marked as Claimed.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Delete row ────────────────────────────────────────────────────────────────
async function deleteRow(excelRowNumber) {
  if (!confirm("Delete this transaction? This cannot be undone.")) return;
  try {
    log("Deleting row " + excelRowNumber + "...");
    // Overwrite with blank row
    await writeExcelRange(CONFIG.sheetName, `A${excelRowNumber}:J${excelRowNumber}`, [["","","","","","","","","",""]]);
    log("Row cleared.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function formatDisplayDate(value) {
  const iso = rawDateToInputValue(value);
  if (!iso) return String(value ?? "");
  const [y, m, d] = iso.split("-");
  return d + "/" + m + "/" + y;
}

function formatAmount(value) {
  const n = parseFloat(String(value).replace(/[$,]/g, ""));
  if (isNaN(n)) return "—";
  return n.toLocaleString("en-SG", { style: "currency", currency: "SGD", minimumFractionDigits: 2 });
}

function parseMoney(value) {
  const n = parseFloat(String(value ?? "").replace(/[$,]/g, ""));
  return isNaN(n) ? 0 : Math.abs(n);
}

function getClaimAmount(row) {
  if (!row || row.claimable !== "Yes") return 0;
  const expenseAmount = parseMoney(row.amount);
  const storedClaimAmount = parseMoney(row.claimAmount);
  const claimAmount = storedClaimAmount > 0 ? storedClaimAmount : expenseAmount;
  return Math.min(expenseAmount, claimAmount);
}

async function ensureClaimHeaders() {
  await writeExcelRange(CONFIG.sheetName, "I1:J1", [["Claim Amount", "Claim Account"]]);
}

function escapeHtml(v) {
  return String(v ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
