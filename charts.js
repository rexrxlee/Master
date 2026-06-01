let expenseChart = null;
let subCategoryChart = null;

Chart.register(ChartDataLabels);

function drawMonthlyExpenseChart(data) {
  const monthlyTotals = {};

  data.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const amount = getAmount(row["Amount"]);

    if (!date) return;

    const monthKey =
      date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");

    monthlyTotals[monthKey] = (monthlyTotals[monthKey] || 0) + amount;
  });

  const labels = Object.keys(monthlyTotals).sort();
  const values = labels.map(month => monthlyTotals[month]);

  const ctx = document.getElementById("expenseChart");

  if (expenseChart) {
    expenseChart.destroy();
  }

  expenseChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: "Monthly Expenses",
        data: values,
        borderWidth: 3,
        tension: 0.3,
        pointRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,

      layout: {
        padding: {
          top: 35,
          right: 25,
          bottom: 10,
          left: 10
        }
      },

      plugins: {
        legend: {
          display: true
        },

        datalabels: {
          display: true,
          anchor: "end",
          align: "top",
          offset: 6,
          clip: false,
          formatter: value =>
            value.toLocaleString("en-SG", {
              style: "currency",
              currency: "SGD",
              maximumFractionDigits: 0
            }),
          font: {
            size: 11,
            weight: "bold"
          }
        }
      },

      scales: {
        y: {
          beginAtZero: true,
          grace: "20%"
        }
      }
    }
  });
}

function drawSubCategoryMonthlyChart(data) {
  const monthlySubCategoryTotals = {};
  const subCategoriesSet = new Set();

  data.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    const amount = getAmount(row["Amount"]);
    const subCategory = clean(row["Sub Category"]);

    if (!date || !subCategory) return;

    const monthKey =
      date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");

    if (!monthlySubCategoryTotals[monthKey]) {
      monthlySubCategoryTotals[monthKey] = {};
    }

    monthlySubCategoryTotals[monthKey][subCategory] =
      (monthlySubCategoryTotals[monthKey][subCategory] || 0) + amount;

    subCategoriesSet.add(subCategory);
  });

  const labels = Object.keys(monthlySubCategoryTotals).sort();
  const subCategories = [...subCategoriesSet].sort();

  const datasets = subCategories.map(subCategory => ({
    label: subCategory,
    data: labels.map(month =>
      monthlySubCategoryTotals[month][subCategory] || 0
    ),
    borderWidth: 2,
    tension: 0.3,
    pointRadius: 4
  }));

  const ctx = document.getElementById("subCategoryChart");

  if (subCategoryChart) {
    subCategoryChart.destroy();
  }

  subCategoryChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          display: true,
          position: "bottom"
        },

        datalabels: {
          display: false
        }
      },

      scales: {
        y: {
          beginAtZero: true,
          grace: "20%"
        }
      }
    }
  });
}