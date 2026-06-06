let expenseChart = null;
let subCategoryChart = null;
let incomeSubCategoryChart = null;

Chart.register(ChartDataLabels);

function drawMonthlyExpenseChart(expenseData, incomeData, filters = {}) {
  const monthlyExpenses = {};
  const monthlyIncome = {};

  expenseData.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const amount = getAmount(row["Amount"]);
    if (!date) return;
    const key = monthKey(date);
    monthlyExpenses[key] = (monthlyExpenses[key] || 0) + amount;
  });

  (incomeData || []).forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const amount = getAmount(row["Amount"]);
    if (!date) return;
    const key = monthKey(date);
    monthlyIncome[key] = (monthlyIncome[key] || 0) + amount;
  });

  const hasExpenseFilters = Boolean(
    (filters.mainCategories && filters.mainCategories.length) ||
    (filters.subCategories && filters.subCategories.length) ||
    (filters.transactions && filters.transactions.length)
  );
  const expenseMonths = Object.keys(monthlyExpenses).sort();
  const incomeMonths = Object.keys(monthlyIncome).sort();
  const rawMonths = hasExpenseFilters
    ? expenseMonths
    : [...new Set([...expenseMonths, ...incomeMonths])].sort();
  const allMonths = fillMonthRange(rawMonths);

  // If no data at all, show a placeholder
  if (rawMonths.length === 0) {
    const ctx = document.getElementById("expenseChart");
    if (expenseChart) { expenseChart.destroy(); expenseChart = null; }
    if (ctx) {
      const ctxParent = ctx.parentElement;
      ctx.style.display = "none";
      if (!ctxParent.querySelector(".no-data-msg")) {
        const msg = document.createElement("div");
        msg.className = "no-data-msg";
        msg.style.cssText = "text-align:center;color:#aaa;padding:60px 0;font-size:16px;";
        msg.textContent = "No data to display for the selected filters.";
        ctxParent.appendChild(msg);
      }
    }
    return;
  }

  const ctx = document.getElementById("expenseChart");
  if (ctx) {
    ctx.style.display = "";
    const oldMsg = ctx.parentElement.querySelector(".no-data-msg");
    if (oldMsg) oldMsg.remove();
  }

  const expenseValues = allMonths.map(m => monthlyExpenses[m] || 0);
  const incomeValues  = allMonths.map(m => monthlyIncome[m]  || 0);
  const netValues     = allMonths.map(m => (monthlyIncome[m] || 0) - (monthlyExpenses[m] || 0));

  if (expenseChart) expenseChart.destroy();

  expenseChart = new Chart(ctx, {
    data: {
      labels: allMonths,
      datasets: [
        {
          type: "bar",
          label: "Income",
          data: incomeValues,
          backgroundColor: "rgba(39,174,96,0.6)",
          borderColor: "rgba(39,174,96,1)",
          borderWidth: 1,
          yAxisID: "y",
          datalabels: {
            display: false,
            anchor: "end",
            align: "top",
            color: "#27ae60",
            font: { size: 9, weight: "bold" },
            formatter: v => v > 0 ? formatChartCurrency(v) : ""
          }
        },
        {
          type: "bar",
          label: "Expenses",
          data: expenseValues,
          backgroundColor: "rgba(231,76,60,0.6)",
          borderColor: "rgba(231,76,60,1)",
          borderWidth: 1,
          yAxisID: "y",
          datalabels: {
            display: false,
            anchor: "end",
            align: "top",
            color: "#c0392b",
            font: { size: 9, weight: "bold" },
            formatter: v => v > 0 ? formatChartCurrency(v) : ""
          }
        },
        {
          type: "line",
          label: "Net Cashflow",
          data: netValues,
          borderColor: "#2c3e50",
          backgroundColor: "transparent",
          borderWidth: 2,
          tension: 0.3,
          pointRadius: 5,
          yAxisID: "y",
          datalabels: {
            display: false,
            anchor: "end",
            align: "top",
            offset: 4,
            clip: false,
            color: ctx => netValues[ctx.dataIndex] >= 0 ? "#27ae60" : "#c0392b",
            formatter: value => (value >= 0 ? "+" : "") + formatChartCurrency(value),
            font: { size: 10, weight: "bold" }
          }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 45, right: 25, bottom: 10, left: 10 } },
      plugins: {
        legend: { display: true, position: "top" },
        tooltip: {
          callbacks: {
            label: ctx => {
              const val = ctx.parsed.y;
              const prefix = ctx.dataset.label === "Net Cashflow" && val >= 0 ? "+" : "";
              return ctx.dataset.label + ": " + prefix + formatExactCurrency(val);
            }
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: "20%",
          ticks: { callback: v => formatAxisCurrency(v) }
        }
      }
    }
  });
}

function drawSubCategoryMonthlyChart(data, filters = {}) {
  const selectedSubs = filters.subCategories || [];

  // Detect if all visible data/date range is within a single month → use daily grouping
  const dates = data.map(row => parseExcelDate(row["Date"])).filter(Boolean);
  const rangeStart = filters.fromDate ? parseInputDate(filters.fromDate) : (dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null);
  const rangeEnd = filters.toDate ? parseInputDate(filters.toDate) : (dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null);
  const months = new Set((rangeStart && rangeEnd ? [rangeStart, rangeEnd] : dates).map(d => monthKey(d)));
  const useDailyView = rangeStart && rangeEnd
    ? monthKey(rangeStart) === monthKey(rangeEnd)
    : months.size <= 1;

  const bucketKey = useDailyView
    ? d => d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0")
    : d => d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");

  const bucketTotals = {};
  const subCategoriesSet = new Set();

  data.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const amount = getAmount(row["Amount"]);
    const subCategory = clean(row["Sub Category"]);
    if (!date || !subCategory) return;
    const key = bucketKey(date);
    if (!bucketTotals[key]) bucketTotals[key] = {};
    bucketTotals[key][subCategory] = (bucketTotals[key][subCategory] || 0) + amount;
    subCategoriesSet.add(subCategory);
  });

  const rawLabels = Object.keys(bucketTotals).sort();
  const labels = buildContinuousLabels(rawLabels, rangeStart, rangeEnd, useDailyView);
  const subCategories = [...new Set([...subCategoriesSet, ...selectedSubs])].sort();

  if (labels.length === 0 || subCategories.length === 0) {
    const ctx = document.getElementById("subCategoryChart");
    if (subCategoryChart) { subCategoryChart.destroy(); subCategoryChart = null; }
    if (ctx) {
      const ctxParent = ctx.parentElement;
      ctx.style.display = "none";
      if (!ctxParent.querySelector(".no-data-msg")) {
        const msg = document.createElement("div");
        msg.className = "no-data-msg";
        msg.style.cssText = "text-align:center;color:#aaa;padding:80px 0;font-size:16px;";
        msg.textContent = "No expense data to display for the selected filters.";
        ctxParent.appendChild(msg);
      }
    }
    return;
  }

  const ctx = document.getElementById("subCategoryChart");
  if (ctx) {
    ctx.style.display = "";
    const oldMsg = ctx.parentElement.querySelector(".no-data-msg");
    if (oldMsg) oldMsg.remove();
  }

  const datasets = subCategories.map(subCategory => ({
    label: subCategory,
    data: labels.map(m => (bucketTotals[m] && bucketTotals[m][subCategory]) || 0),
    borderWidth: 2,
    tension: 0.3,
    pointRadius: 4,
    datalabels: {
      display: true,
      anchor: "end",
      align: "top",
      offset: 4,
      clamp: true,
      font: { size: 9 },
      formatter: v => formatChartCurrency(v)
    }
  }));

  if (subCategoryChart) subCategoryChart.destroy();

  subCategoryChart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 30, right: 10 } },
      plugins: {
        legend: { display: true, position: "bottom" },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.dataset.label}: ${formatExactCurrency(ctx.parsed.y || 0)}`
          }
        },
        title: {
          display: useDailyView,
          text: "Daily view — " + (rangeStart ? monthKey(rangeStart) : ""),
          font: { size: 13 },
          color: "#888"
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: "20%",
          ticks: { callback: v => formatAxisCurrency(v) }
        }
      }
    }
  });
}

function drawIncomeSubCategoryChart(data, filters = {}) {
  const ctx = document.getElementById("incomeSubCategoryChart");
  if (!ctx) return;

  const monthlyTotals = {};
  const subCategoriesSet = new Set();
  const dates = (data || []).map(row => parseExcelDate(row["Date"])).filter(Boolean);
  const rangeStart = filters.fromDate ? parseInputDate(filters.fromDate) : (dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null);
  const rangeEnd = filters.toDate ? parseInputDate(filters.toDate) : (dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null);

  (data || []).forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const subCategory = clean(row["Sub Category"]) || "Uncategorised";
    if (!date || !subCategory) return;
    const key = monthKey(date);
    if (!monthlyTotals[key]) monthlyTotals[key] = {};
    monthlyTotals[key][subCategory] = (monthlyTotals[key][subCategory] || 0) + getAmount(row["Amount"]);
    subCategoriesSet.add(subCategory);
  });

  const rawLabels = Object.keys(monthlyTotals).sort();
  const labels = buildContinuousLabels(rawLabels, rangeStart, rangeEnd, false);
  const subCategories = [...subCategoriesSet].sort((a, b) => {
    const totalA = labels.reduce((sum, month) => sum + ((monthlyTotals[month] && monthlyTotals[month][a]) || 0), 0);
    const totalB = labels.reduce((sum, month) => sum + ((monthlyTotals[month] && monthlyTotals[month][b]) || 0), 0);
    return totalB - totalA;
  });

  if (labels.length === 0 || subCategories.length === 0) {
    if (incomeSubCategoryChart) { incomeSubCategoryChart.destroy(); incomeSubCategoryChart = null; }
    const ctxParent = ctx.parentElement;
    ctx.style.display = "none";
    if (!ctxParent.querySelector(".no-data-msg")) {
      const msg = document.createElement("div");
      msg.className = "no-data-msg";
      msg.style.cssText = "text-align:center;color:#aaa;padding:70px 0;font-size:16px;";
      msg.textContent = "No income data to display for the selected date range.";
      ctxParent.appendChild(msg);
    }
    return;
  }

  ctx.style.display = "";
  const oldMsg = ctx.parentElement.querySelector(".no-data-msg");
  if (oldMsg) oldMsg.remove();

  const datasets = subCategories.map(subCategory => ({
    label: subCategory,
    data: labels.map(month => (monthlyTotals[month] && monthlyTotals[month][subCategory]) || 0),
    borderWidth: 2,
    tension: 0.3,
    pointRadius: 4,
    datalabels: {
      display: true,
      anchor: "end",
      align: "top",
      offset: 4,
      clamp: true,
      font: { size: 9 },
      formatter: value => value > 0 ? formatChartCurrency(value) : ""
    }
  }));

  if (incomeSubCategoryChart) incomeSubCategoryChart.destroy();

  incomeSubCategoryChart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 30, right: 16, bottom: 8, left: 8 } },
      plugins: {
        legend: { display: true, position: "bottom" },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.dataset.label}: ${formatExactCurrency(ctx.parsed.y || 0)}`
          }
        },
        datalabels: { display: true }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: "20%",
          ticks: { callback: v => formatAxisCurrency(v) }
        }
      }
    }
  });
}

function monthKey(date) {
  return date.getFullYear() + "-" + String(date.getMonth()+1).padStart(2,"0");
}

function parseInputDate(value) {
  if (!value) return null;
  const [y, m, d] = String(value).split("-").map(Number);
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  return new Date(y, m - 1, d);
}

function fillMonthRange(monthKeys) {
  if (!monthKeys.length) return [];
  const [startY, startM] = monthKeys[0].split("-").map(Number);
  const [endY, endM] = monthKeys[monthKeys.length - 1].split("-").map(Number);
  const labels = [];
  const cursor = new Date(startY, startM - 1, 1);
  const end = new Date(endY, endM - 1, 1);
  while (cursor <= end) {
    labels.push(monthKey(cursor));
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return labels;
}

function buildContinuousLabels(rawLabels, rangeStart, rangeEnd, useDailyView) {
  if (rangeStart && rangeEnd) {
    const labels = [];
    const cursor = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), useDailyView ? rangeStart.getDate() : 1);
    const end = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), useDailyView ? rangeEnd.getDate() : 1);
    while (cursor <= end) {
      labels.push(useDailyView ? dateKey(cursor) : monthKey(cursor));
      useDailyView ? cursor.setDate(cursor.getDate() + 1) : cursor.setMonth(cursor.getMonth() + 1);
    }
    return labels;
  }
  if (!rawLabels.length) return [];
  if (!useDailyView) return fillMonthRange(rawLabels);
  return fillDayRange(rawLabels);
}

function fillDayRange(dayKeys) {
  if (!dayKeys.length) return [];
  const start = parseInputDate(dayKeys[0]);
  const end = parseInputDate(dayKeys[dayKeys.length - 1]);
  if (!start || !end) return dayKeys;
  const labels = [];
  const cursor = new Date(start);
  while (cursor <= end) {
    labels.push(dateKey(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return labels;
}

function dateKey(date) {
  return monthKey(date) + "-" + String(date.getDate()).padStart(2,"0");
}

function formatExactCurrency(value) {
  return Number(value || 0).toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatChartCurrency(value) {
  return Number(value || 0).toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    maximumFractionDigits: 0
  });
}

function formatAxisCurrency(value) {
  const n = Number(value || 0);
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(n);
  if (abs >= 1000) return sign + "$" + (abs / 1000).toLocaleString("en-SG", { maximumFractionDigits: abs >= 10000 ? 0 : 1 }) + "k";
  return sign + "$" + abs.toLocaleString("en-SG", { maximumFractionDigits: 0 });
}
