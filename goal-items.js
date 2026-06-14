// ═══════════════════════════════════════════════════════════════════
//  GOAL ITEMS (Planner Tab) — Fintrack  (goal-items.js)
//
//  Storage: Excel sheet "Goal Items"
//  Columns: A=ItemID, B=GoalName, C=Description, D=EstAmount,
//           E=RequiredBy, F=Priority, G=Link, H=Comments,
//           I=Status, J=PurchasedAmount, K=PurchasedDate
//
//  Status values: "Planned" | "Purchased" | "Cancelled"
//
//  On "Mark Purchased":
//    1. Updates row I/J/K in "Goal Items" sheet
//    2. Appends a transaction row to CONFIG.sheetName
//       (same format as add-transaction.js saves)
// ═══════════════════════════════════════════════════════════════════

const GOAL_ITEMS_SHEET   = "Goal Items";
const GOAL_ITEMS_HEADERS = ["ItemID","GoalName","Description","EstAmount","RequiredBy","Priority","Link","Comments","Status","PurchasedAmount","PurchasedDate"];

// Module state — populated by loadGoalItemsTab()
let _goalItems      = [];   // parsed rows from Goal Items sheet
let _giGoalNames    = [];   // goal names pulled from goalsData (parent module)
let _giAccounts     = [];   // accounts from allAccounts (parent module)
let _giExpanded     = {};   // { goalName: bool } collapse state
let _giSheetExists  = false;
let _giAutoSaveTimer = null;
let _giAutoSaveInFlight = false;

// ─── Entry Point (called from tab switch) ─────────────────────────

async function loadGoalItemsTab() {
  const container = document.getElementById("plannerTabContent");
  if (!container) return;
  container.innerHTML = `<div class="gi-loading">⏳ Loading goal items…</div>`;

  try {
    const arrayBuffer = await downloadExcelFile();
    const workbook    = XLSX.read(arrayBuffer, { type: "array" });

    // Grab goal names from already-loaded goalsData (parent module state)
    _giGoalNames = (typeof goalsData !== "undefined" ? goalsData : []).map(g => g.name);
    _giAccounts  = (typeof allAccounts !== "undefined" ? allAccounts : []);

    // Read Goal Items sheet (may not exist yet)
    const sheet = workbook.Sheets[GOAL_ITEMS_SHEET];
    if (sheet) {
      _giSheetExists = true;
      _goalItems = _readGoalItemsSheet(sheet);
    } else {
      _giSheetExists = false;
      _goalItems = [];
    }

    // Default expand first goal
    if (Object.keys(_giExpanded).length === 0 && _giGoalNames.length > 0) {
      _giExpanded[_giGoalNames[0]] = true;
    }

    _renderPlannerTab(container);
  } catch (err) {
    container.innerHTML = `<div class="gi-error">❌ Failed to load: ${escapeHtmlGI(err.message)}</div>`;
    console.error(err);
  }
}

// ─── Sheet Reader ──────────────────────────────────────────────────

function _readGoalItemsSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  if (rows.length < 2) return [];
  // rows[0] = headers, skip
  return rows.slice(1).map(row => ({
    itemId:          String(row[0] ?? "").trim(),
    goalName:        String(row[1] ?? "").trim(),
    description:     String(row[2] ?? "").trim(),
    estAmount:       Number(row[3] ?? 0) || 0,
    requiredBy:      _giDateToInput(row[4]),
    priority:        String(row[5] ?? "Medium").trim() || "Medium",
    link:            String(row[6] ?? "").trim(),
    comments:        String(row[7] ?? "").trim(),
    status:          String(row[8] ?? "Planned").trim() || "Planned",
    purchasedAmount: Number(row[9] ?? 0) || 0,
    purchasedDate:   _giDateToInput(row[10]),
  })).filter(r => r.itemId && r.goalName);
}

// ─── Render ────────────────────────────────────────────────────────

function _renderPlannerTab(container) {
  if (_giGoalNames.length === 0) {
    container.innerHTML = `
      <div class="gi-empty-state">
        <div class="gi-empty-icon">◎</div>
        <div class="gi-empty-title">No goals yet</div>
        <div class="gi-empty-sub">Add goals in the Overview tab first, then come back to plan your purchases.</div>
      </div>`;
    return;
  }

  let html = `
    <div class="gi-header-bar">
      <div class="gi-header-summary" id="giHeaderSummary"></div>
      <span class="gi-autosave-status" id="goalItemsAutosaveStatus">Autosaves to Excel</span>
    </div>
    <div class="gi-goals-list" id="giGoalsList"></div>
  `;
  container.innerHTML = html;

  _renderGoalsSummaryBar();
  _renderAllGoalSections();
}

function _renderGoalsSummaryBar() {
  const el = document.getElementById("giHeaderSummary");
  if (!el) return;

  const totalItems     = _goalItems.filter(i => i.status !== "Cancelled").length;
  const totalPlanned   = _goalItems.filter(i => i.status === "Planned").reduce((s,i) => s + i.estAmount, 0);
  const totalPurchased = _goalItems.filter(i => i.status === "Purchased").reduce((s,i) => s + (i.purchasedAmount || i.estAmount), 0);
  const overdue        = _goalItems.filter(i => i.status === "Planned" && i.requiredBy && i.requiredBy < _todayStr()).length;

  el.innerHTML = `
    <span class="gi-sum-pill neutral">${totalItems} items</span>
    <span class="gi-sum-pill warn">📋 ${_fmtCurrency(totalPlanned)} planned</span>
    <span class="gi-sum-pill ok">✓ ${_fmtCurrency(totalPurchased)} purchased</span>
    ${overdue > 0 ? `<span class="gi-sum-pill red">⚠ ${overdue} overdue</span>` : ""}
  `;
}

function _renderAllGoalSections() {
  const list = document.getElementById("giGoalsList");
  if (!list) return;
  list.innerHTML = "";

  _giGoalNames.forEach(goalName => {
    const goalData  = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === goalName);
    const items     = _goalItems.filter(i => i.goalName === goalName);
    const section   = document.createElement("div");
    section.className = "gi-goal-section";
    section.id = "giSection_" + _slugify(goalName);
    section.innerHTML = _buildGoalSectionHTML(goalName, goalData, items);
    list.appendChild(section);
    _renderInsightsForGoal(goalName, items);
  });
}

function _buildGoalSectionHTML(goalName, goalData, items) {
  const isExpanded  = !!_giExpanded[goalName];
  const slug        = _slugify(goalName);
  const goalTarget  = goalData ? goalData.target * (1 + (goalData.goalBuffer || 0) / 100) : 0;
  const goalColor   = goalData ? goalData.color : "#6366f1";

  // Budget maths
  const active      = items.filter(i => i.status !== "Cancelled");
  const planned     = active.filter(i => i.status === "Planned");
  const purchased   = active.filter(i => i.status === "Purchased");
  const subTotal    = active.reduce((s,i) => s + i.estAmount, 0);
  const purchTotal  = purchased.reduce((s,i) => s + (i.purchasedAmount || i.estAmount), 0);
  const gap         = goalTarget - subTotal;        // positive = room left, negative = over
  const overTarget  = goalData && subTotal > goalTarget;
  const bufferedAmt = goalData ? goalData.target * (1 + (goalData.goalBuffer||0)/100) : 0;
  const overBuffered = goalData && subTotal > bufferedAmt && goalData.goalBuffer > 0;

  // Budget status badge
  let budgetBadge = "";
  if (!goalData || goalTarget === 0) {
    budgetBadge = `<span class="gi-budget-badge neutral">No target set</span>`;
  } else if (overTarget) {
    const over = subTotal - goalTarget;
    budgetBadge = `<span class="gi-budget-badge over">⚠ ${_fmtCurrency(over)} over budget
      <button class="gi-bump-btn" onclick="bumpGoalTarget('${escapeHtmlGI(goalName)}')">Bump target ↑</button>
    </span>`;
  } else if (gap > 0) {
    budgetBadge = `<span class="gi-budget-badge under">${_fmtCurrency(gap)} unallocated</span>`;
  } else {
    budgetBadge = `<span class="gi-budget-badge exact">✓ Exactly budgeted</span>`;
  }

  // Progress within sub-items
  const progressPct = goalTarget > 0 ? Math.min(100, (purchTotal / goalTarget) * 100) : 0;

  // Overdue count
  const overdueCount = planned.filter(i => i.requiredBy && i.requiredBy < _todayStr()).length;

  // Sort items: overdue first, then by requiredBy, then by priority weight
  const priOrder = { Critical:0, High:1, Medium:2, Low:3 };
  const sortedItems = [...items].sort((a, b) => {
    const aOverdue = a.status === "Planned" && a.requiredBy && a.requiredBy < _todayStr() ? 0 : 1;
    const bOverdue = b.status === "Planned" && b.requiredBy && b.requiredBy < _todayStr() ? 0 : 1;
    if (aOverdue !== bOverdue) return aOverdue - bOverdue;
    if (a.requiredBy && b.requiredBy) return a.requiredBy.localeCompare(b.requiredBy);
    return (priOrder[a.priority]??2) - (priOrder[b.priority]??2);
  });

  const headerExtra = [
    active.length > 0 ? `<span class="gi-hdr-count">${active.length} item${active.length!==1?"s":""}</span>` : "",
    overdueCount > 0 ? `<span class="gi-hdr-overdue">⚠ ${overdueCount} overdue</span>` : "",
  ].filter(Boolean).join("");

  return `
    <div class="gi-section-header" onclick="toggleGISection('${escapeHtmlGI(goalName)}')" style="border-left:3px solid ${goalColor};">
      <div class="gi-section-title-row">
        <span class="gi-chevron ${isExpanded ? "open" : ""}">▶</span>
        <span class="gi-section-name">${escapeHtmlGI(goalName)}</span>
        ${headerExtra}
      </div>
      <div class="gi-section-meta">
        <span class="gi-meta-item">
          <span class="gi-meta-label">Sub-items total</span>
          <span class="gi-meta-val ${overTarget ? "red" : ""}">${_fmtCurrency(subTotal)}</span>
        </span>
        ${goalTarget > 0 ? `
        <span class="gi-meta-item">
          <span class="gi-meta-label">Goal target</span>
          <span class="gi-meta-val">${_fmtCurrency(goalTarget)}</span>
        </span>` : ""}
        <span class="gi-meta-item">${budgetBadge}</span>
      </div>
      ${goalTarget > 0 ? `
      <div class="gi-section-progress">
        <div class="gi-prog-bar">
          <div class="gi-prog-fill" style="width:${progressPct.toFixed(1)}%;background:${goalColor};"></div>
        </div>
        <span class="gi-prog-label">${progressPct.toFixed(0)}% purchased</span>
      </div>` : ""}
    </div>

    <div class="gi-section-body ${isExpanded ? "open" : ""}" id="giBody_${slug}">
      ${_buildItemsTable(goalName, goalData, sortedItems)}
      ${items.filter(i => i.status !== "Cancelled").length > 0 ? _buildInsightsPanel(goalName, goalData, items) : ""}
      <div class="gi-add-row" id="giAddRow_${slug}">
        ${_buildAddItemForm(goalName)}
      </div>
    </div>
  `;
}

function _buildItemsTable(goalName, goalData, items) {
  if (items.length === 0) {
    return `<div class="gi-no-items">No items yet — add your first planned purchase below.</div>`;
  }

  const rows = items.map(item => _buildItemRow(item)).join("");
  return `
    <div class="gi-table-wrap">
      <table class="gi-table">
        <thead>
          <tr>
            <th class="gi-th-check"></th>
            <th class="gi-th-desc">Description</th>
            <th class="gi-th-amt">Est. Amount</th>
            <th class="gi-th-date">Required By</th>
            <th class="gi-th-pri">Priority</th>
            <th class="gi-th-link">Link</th>
            <th class="gi-th-comments">Comments</th>
            <th class="gi-th-status">Status</th>
            <th class="gi-th-actions"></th>
          </tr>
        </thead>
        <tbody id="giTbody_${_slugify(goalName)}">
          ${rows}
        </tbody>
      </table>
    </div>
  `;
}

function _buildItemRow(item) {
  const isComplete  = item.status === "Purchased";
  const isCancelled = item.status === "Cancelled";
  const isOverdue   = item.status === "Planned" && item.requiredBy && item.requiredBy < _todayStr();

  const rowClass = [
    "gi-item-row",
    isComplete  ? "gi-row-purchased"  : "",
    isCancelled ? "gi-row-cancelled"  : "",
    isOverdue   ? "gi-row-overdue"    : "",
  ].filter(Boolean).join(" ");

  const priColors = { Critical:"#dc2626", High:"#ea580c", Medium:"#ca8a04", Low:"#16a34a" };
  const priColor  = priColors[item.priority] || "#ca8a04";

  const displayDate = item.requiredBy ? _inputDateToDisplay(item.requiredBy) : "—";
  const linkHtml    = item.link
    ? `<a href="${escapeHtmlGI(item.link)}" target="_blank" rel="noopener" class="gi-link-btn" title="${escapeHtmlGI(item.link)}">🔗</a>`
    : `<span class="gi-link-empty">—</span>`;

  const statusBadge = isComplete
    ? `<span class="gi-status-badge purchased">✓ Purchased<br><small>${_fmtCurrency(item.purchasedAmount || item.estAmount)} on ${item.purchasedDate ? _inputDateToDisplay(item.purchasedDate) : "?"}</small></span>`
    : isCancelled
    ? `<span class="gi-status-badge cancelled">✕ Cancelled</span>`
    : isOverdue
    ? `<span class="gi-status-badge overdue">⚠ Overdue</span>`
    : `<span class="gi-status-badge planned">Planned</span>`;

  const actionBtns = isComplete || isCancelled ? `
    <button class="gi-action-btn gi-undo-btn" onclick="undoGoalItem('${escapeHtmlGI(item.itemId)}')" title="Undo">↩</button>
    <button class="gi-action-btn gi-del-btn" onclick="deleteGoalItem('${escapeHtmlGI(item.itemId)}')" title="Delete">🗑</button>
  ` : `
    <button class="gi-action-btn gi-edit-btn" onclick="openEditGoalItem('${escapeHtmlGI(item.itemId)}')" title="Edit">✏️</button>
    <button class="gi-action-btn gi-buy-btn" onclick="openPurchaseModal('${escapeHtmlGI(item.itemId)}')" title="Mark Purchased">✓</button>
    <button class="gi-action-btn gi-cancel-btn" onclick="cancelGoalItem('${escapeHtmlGI(item.itemId)}')" title="Cancel item">✕</button>
    <button class="gi-action-btn gi-del-btn" onclick="deleteGoalItem('${escapeHtmlGI(item.itemId)}')" title="Delete">🗑</button>
  `;

  return `
    <tr class="${rowClass}" id="giRow_${item.itemId}">
      <td class="gi-td-check">
        ${!isComplete && !isCancelled ? `<input type="checkbox" class="gi-check" onchange="onGICheckboxChange('${escapeHtmlGI(item.itemId)}', this.checked)">` : ""}
      </td>
      <td class="gi-td-desc">
        <span class="gi-desc-text ${isComplete||isCancelled?"gi-strikethrough":""}">${escapeHtmlGI(item.description)}</span>
      </td>
      <td class="gi-td-amt">
        <span class="gi-mono ${isComplete?"gi-muted":""}">${_fmtCurrency(item.estAmount)}</span>
        ${isComplete && item.purchasedAmount && item.purchasedAmount !== item.estAmount
          ? `<span class="gi-actual-diff ${item.purchasedAmount > item.estAmount ? "red" : "green"}">
               actual: ${_fmtCurrency(item.purchasedAmount)}
             </span>`
          : ""}
      </td>
      <td class="gi-td-date">
        <span class="${isOverdue ? "gi-overdue-text" : ""}">${displayDate}</span>
      </td>
      <td class="gi-td-pri">
        <span class="gi-pri-dot" style="background:${priColor};"></span>
        <span style="color:${priColor};font-size:12px;font-weight:600;">${item.priority}</span>
      </td>
      <td class="gi-td-link">${linkHtml}</td>
      <td class="gi-td-comments">
        <span class="gi-comment-text" title="${escapeHtmlGI(item.comments)}">${item.comments ? escapeHtmlGI(item.comments.length > 40 ? item.comments.slice(0,40)+"…" : item.comments) : "—"}</span>
      </td>
      <td class="gi-td-status">${statusBadge}</td>
      <td class="gi-td-actions">${actionBtns}</td>
    </tr>
    <!-- Inline edit row (hidden by default) -->
    <tr class="gi-edit-row" id="giEditRow_${item.itemId}" style="display:none;">
      <td colspan="9">${_buildEditRowHTML(item)}</td>
    </tr>
  `;
}

function _buildEditRowHTML(item) {
  return `
    <div class="gi-edit-panel">
      <div class="gi-edit-grid">
        <div class="gi-edit-group gi-edit-full">
          <label class="gi-edit-label">Description</label>
          <input class="gi-edit-input" id="giEdit_desc_${item.itemId}" value="${escapeHtmlGI(item.description)}" placeholder="What are you buying?">
        </div>
        <div class="gi-edit-group">
          <label class="gi-edit-label">Estimated Amount (SGD)</label>
          <input class="gi-edit-input" type="number" step="0.01" id="giEdit_amt_${item.itemId}" value="${item.estAmount}">
        </div>
        <div class="gi-edit-group">
          <label class="gi-edit-label">Required By</label>
          <input class="gi-edit-input" type="date" id="giEdit_date_${item.itemId}" value="${item.requiredBy || ""}">
        </div>
        <div class="gi-edit-group">
          <label class="gi-edit-label">Priority</label>
          <select class="gi-edit-input" id="giEdit_pri_${item.itemId}">
            ${["Critical","High","Medium","Low"].map(p => `<option value="${p}" ${item.priority===p?"selected":""}>${p}</option>`).join("")}
          </select>
        </div>
        <div class="gi-edit-group gi-edit-full">
          <label class="gi-edit-label">Link (URL for reference)</label>
          <input class="gi-edit-input" type="url" id="giEdit_link_${item.itemId}" value="${escapeHtmlGI(item.link)}" placeholder="https://...">
        </div>
        <div class="gi-edit-group gi-edit-full">
          <label class="gi-edit-label">Comments</label>
          <input class="gi-edit-input" id="giEdit_comments_${item.itemId}" value="${escapeHtmlGI(item.comments)}" placeholder="Notes, specs, alternatives…">
        </div>
      </div>
      <div class="gi-edit-actions">
        <button class="gi-save-item-btn" onclick="saveEditGoalItem('${escapeHtmlGI(item.itemId)}')">✓ Save Changes</button>
        <button class="gi-cancel-edit-btn" onclick="closeEditGoalItem('${escapeHtmlGI(item.itemId)}')">Cancel</button>
      </div>
    </div>
  `;
}

function _buildAddItemForm(goalName) {
  const slug = _slugify(goalName);
  return `
    <details class="gi-add-details" id="giAddDetails_${slug}">
      <summary class="gi-add-summary">＋ Add item to "${escapeHtmlGI(goalName)}"</summary>
      <div class="gi-add-form">
        <div class="gi-add-grid">
          <div class="gi-add-group gi-add-full">
            <label class="gi-edit-label">Description *</label>
            <input class="gi-edit-input" id="giAdd_desc_${slug}" placeholder="e.g. Flight SIN→NRT, Sony WH-1000XM5">
          </div>
          <div class="gi-add-group">
            <label class="gi-edit-label">Estimated Amount (SGD) *</label>
            <input class="gi-edit-input" type="number" step="0.01" id="giAdd_amt_${slug}" placeholder="0.00">
          </div>
          <div class="gi-add-group">
            <label class="gi-edit-label">Required By</label>
            <input class="gi-edit-input" type="date" id="giAdd_date_${slug}">
          </div>
          <div class="gi-add-group">
            <label class="gi-edit-label">Priority</label>
            <select class="gi-edit-input" id="giAdd_pri_${slug}">
              <option value="High">High</option>
              <option value="Medium" selected>Medium</option>
              <option value="Low">Low</option>
              <option value="Critical">Critical</option>
            </select>
          </div>
          <div class="gi-add-group gi-add-full">
            <label class="gi-edit-label">Link (paste product/booking URL)</label>
            <input class="gi-edit-input" type="url" id="giAdd_link_${slug}" placeholder="https://shopee.sg/...">
          </div>
          <div class="gi-add-group gi-add-full">
            <label class="gi-edit-label">Comments</label>
            <input class="gi-edit-input" id="giAdd_comments_${slug}" placeholder="Notes, specs, alternatives…">
          </div>
        </div>
        <div class="gi-add-actions">
          <button class="gi-add-confirm-btn" onclick="addGoalItem('${escapeHtmlGI(goalName)}')">Add Item</button>
          <button class="gi-add-cancel-btn" onclick="document.getElementById('giAddDetails_${slug}').removeAttribute('open')">Cancel</button>
        </div>
      </div>
    </details>
  `;
}

// ─── Toggle collapse ───────────────────────────────────────────────

function toggleGISection(goalName) {
  _giExpanded[goalName] = !_giExpanded[goalName];
  const slug = _slugify(goalName);
  const body    = document.getElementById("giBody_"    + slug);
  const section = document.getElementById("giSection_" + slug);
  if (!body || !section) return;
  body.classList.toggle("open", !!_giExpanded[goalName]);
  const chevron = section.querySelector(".gi-chevron");
  if (chevron) chevron.classList.toggle("open", !!_giExpanded[goalName]);
}

// ─── Add Item ──────────────────────────────────────────────────────

function addGoalItem(goalName) {
  const slug = _slugify(goalName);
  const desc     = document.getElementById("giAdd_desc_"     + slug)?.value.trim();
  const amt      = parseFloat(document.getElementById("giAdd_amt_"  + slug)?.value);
  const date     = document.getElementById("giAdd_date_"     + slug)?.value;
  const pri      = document.getElementById("giAdd_pri_"      + slug)?.value || "Medium";
  const link     = document.getElementById("giAdd_link_"     + slug)?.value.trim() || "";
  const comments = document.getElementById("giAdd_comments_" + slug)?.value.trim() || "";

  if (!desc)      { alert("Please enter a description."); return; }
  if (isNaN(amt) || amt <= 0) { alert("Please enter a valid amount."); return; }

  const newItem = {
    itemId:          _generateId(),
    goalName,
    description:     desc,
    estAmount:       amt,
    requiredBy:      date || "",
    priority:        pri,
    link,
    comments,
    status:          "Planned",
    purchasedAmount: 0,
    purchasedDate:   "",
  };

  _goalItems.push(newItem);

  // Re-render the section
  _reRenderGoalSection(goalName);
  _renderGoalsSummaryBar();

  // Close the add form
  const details = document.getElementById("giAddDetails_" + slug);
  if (details) details.removeAttribute("open");
  scheduleGoalItemsAutoSave();
}

// ─── Edit Item ─────────────────────────────────────────────────────

function openEditGoalItem(itemId) {
  // Close any other open edit rows
  document.querySelectorAll(".gi-edit-row").forEach(r => { r.style.display = "none"; });
  const editRow = document.getElementById("giEditRow_" + itemId);
  if (editRow) editRow.style.display = "";
}

function closeEditGoalItem(itemId) {
  const editRow = document.getElementById("giEditRow_" + itemId);
  if (editRow) editRow.style.display = "none";
}

function saveEditGoalItem(itemId) {
  const item = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;

  const desc     = document.getElementById("giEdit_desc_"     + itemId)?.value.trim();
  const amt      = parseFloat(document.getElementById("giEdit_amt_"  + itemId)?.value);
  const date     = document.getElementById("giEdit_date_"     + itemId)?.value;
  const pri      = document.getElementById("giEdit_pri_"      + itemId)?.value;
  const link     = document.getElementById("giEdit_link_"     + itemId)?.value.trim();
  const comments = document.getElementById("giEdit_comments_" + itemId)?.value.trim();

  if (!desc)            { alert("Description required."); return; }
  if (isNaN(amt)||amt<=0){ alert("Valid amount required."); return; }

  item.description = desc;
  item.estAmount   = amt;
  item.requiredBy  = date || "";
  item.priority    = pri || "Medium";
  item.link        = link || "";
  item.comments    = comments || "";

  _reRenderGoalSection(item.goalName);
  _renderGoalsSummaryBar();
  scheduleGoalItemsAutoSave();
}

// ─── Checkbox quick-action ─────────────────────────────────────────

function onGICheckboxChange(itemId, checked) {
  if (checked) {
    // Uncheck it — we open the purchase modal instead
    const cb = document.querySelector(`#giRow_${itemId} .gi-check`);
    if (cb) cb.checked = false;
    openPurchaseModal(itemId);
  }
}

// ─── Cancel / Undo ────────────────────────────────────────────────

function cancelGoalItem(itemId) {
  const item = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;
  if (!confirm(`Cancel "${item.description}"? It will remain visible but struck through.`)) return;
  item.status = "Cancelled";
  _reRenderGoalSection(item.goalName);
  _renderGoalsSummaryBar();
  scheduleGoalItemsAutoSave();
}

function undoGoalItem(itemId) {
  const item = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;
  item.status          = "Planned";
  item.purchasedAmount = 0;
  item.purchasedDate   = "";
  _reRenderGoalSection(item.goalName);
  _renderGoalsSummaryBar();
  scheduleGoalItemsAutoSave();
}

function deleteGoalItem(itemId) {
  const item = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;
  if (!confirm(`Permanently delete "${item.description}"?`)) return;
  _goalItems = _goalItems.filter(i => i.itemId !== itemId);
  _reRenderGoalSection(item.goalName);
  _renderGoalsSummaryBar();
  scheduleGoalItemsAutoSave();
}

// ─── Purchase Modal ────────────────────────────────────────────────

function openPurchaseModal(itemId) {
  const item = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;

  const goalData  = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === item.goalName);
  const todayStr  = _todayStr();

  // Build account options — savings first, then CC
  const savings = _giAccounts.filter(a => a.type !== "Credit Card");
  const cc      = _giAccounts.filter(a => a.type === "Credit Card");
  const acctOpts = [...savings, ...cc].map(a =>
    `<option value="${escapeHtmlGI(a.name)}">${escapeHtmlGI(a.name)}${a.type==="Credit Card"?" (CC)":""}</option>`
  ).join("");

  // Build sub-category hint (from goals budget items) — default to goal name
  const defaultSubCat = item.description.length > 30 ? item.description.slice(0,30) + "…" : item.description;

  const modal = document.createElement("div");
  modal.className = "gi-modal-overlay";
  modal.id = "giPurchaseModal";
  modal.innerHTML = `
    <div class="gi-modal">
      <div class="gi-modal-header">
        <div class="gi-modal-title">✓ Mark as Purchased</div>
        <button class="gi-modal-close" onclick="closePurchaseModal()">✕</button>
      </div>
      <div class="gi-modal-body">
        <div class="gi-modal-item-name">${escapeHtmlGI(item.description)}</div>
        <div class="gi-modal-item-goal">Goal: <strong>${escapeHtmlGI(item.goalName)}</strong></div>

        <div class="gi-modal-fields">
          <div class="gi-modal-group">
            <label class="gi-edit-label">Actual Amount Paid (SGD) *</label>
            <input class="gi-edit-input" type="number" step="0.01" id="giModal_amt" value="${item.estAmount}">
            <div class="gi-modal-hint">Pre-filled with estimate (${_fmtCurrency(item.estAmount)})</div>
          </div>
          <div class="gi-modal-group">
            <label class="gi-edit-label">Purchase Date *</label>
            <input class="gi-edit-input" type="date" id="giModal_date" value="${todayStr}">
          </div>
          <div class="gi-modal-group">
            <label class="gi-edit-label">Deduct from Account *</label>
            <select class="gi-edit-input" id="giModal_acct">${acctOpts}</select>
          </div>
          <div class="gi-modal-group gi-modal-full">
            <label class="gi-edit-label">Transaction Description (written to Excel)</label>
            <input class="gi-edit-input" id="giModal_txdesc" value="${escapeHtmlGI(defaultSubCat)}">
          </div>
        </div>

        <div class="gi-modal-info">
          <span class="gi-modal-info-icon">ℹ</span>
          This will write a <strong>Savings Goal</strong> transaction row to your Excel sheet,
          reducing the balance of the selected account. The item will be marked as purchased here.
        </div>

        <div class="gi-modal-diff" id="giModalDiff" style="display:none;"></div>
      </div>
      <div class="gi-modal-footer">
        <button class="gi-confirm-purchase-btn" onclick="confirmPurchase('${escapeHtmlGI(itemId)}')">✓ Confirm Purchase</button>
        <button class="gi-modal-cancel-btn" onclick="closePurchaseModal()">Cancel</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  // Show diff hint as user types actual amount
  document.getElementById("giModal_amt").addEventListener("input", function() {
    const actual = parseFloat(this.value) || 0;
    const diff   = actual - item.estAmount;
    const el     = document.getElementById("giModalDiff");
    if (Math.abs(diff) > 0.01) {
      el.style.display = "";
      el.className = "gi-modal-diff " + (diff > 0 ? "over" : "under");
      el.textContent = diff > 0
        ? `+${_fmtCurrency(diff)} over estimate`
        : `${_fmtCurrency(Math.abs(diff))} under estimate`;
    } else {
      el.style.display = "none";
    }
  });
}

function closePurchaseModal() {
  const modal = document.getElementById("giPurchaseModal");
  if (modal) modal.remove();
}

async function confirmPurchase(itemId) {
  const item   = _goalItems.find(i => i.itemId === itemId);
  if (!item) return;

  const amt    = parseFloat(document.getElementById("giModal_amt")?.value);
  const date   = document.getElementById("giModal_date")?.value;
  const acct   = document.getElementById("giModal_acct")?.value;
  const txdesc = document.getElementById("giModal_txdesc")?.value.trim() || item.description;

  if (isNaN(amt) || amt <= 0) { alert("Please enter a valid amount."); return; }
  if (!date)  { alert("Please select a purchase date."); return; }
  if (!acct)  { alert("Please select an account."); return; }

  // Confirm button → disable to prevent double-click
  const btn = document.querySelector(".gi-confirm-purchase-btn");
  if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }

  try {
    // 1. Write transaction row to transaction sheet
    await _writeTransactionRow(date, txdesc, amt, "Savings Goal", "Goal: " + item.goalName, acct);

    // 2. Update item state
    item.status          = "Purchased";
    item.purchasedAmount = amt;
    item.purchasedDate   = date;

    // 3. Save Goal Items sheet
    await _writeGoalItemsSheet();

    // 4. Close modal & re-render
    closePurchaseModal();
    _reRenderGoalSection(item.goalName);
    _renderGoalsSummaryBar();

    // 5. Reload goals overview data so account balances update
    if (typeof loadGoalsPage === "function") await loadGoalsPage();

    alert(`✓ Purchased! ${_fmtCurrency(amt)} deducted from ${acct}.`);
  } catch (err) {
    if (btn) { btn.disabled = false; btn.textContent = "✓ Confirm Purchase"; }
    alert("Failed to save: " + err.message);
    console.error(err);
  }
}

// ─── Bump goal target ──────────────────────────────────────────────

async function bumpGoalTarget(goalName) {
  const goal = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === goalName);
  if (!goal) return;

  const items    = _goalItems.filter(i => i.goalName === goalName && i.status !== "Cancelled");
  const subTotal = items.reduce((s,i) => s + i.estAmount, 0);

  if (!confirm(`Bump target for "${goalName}" from ${_fmtCurrency(goal.target)} to ${_fmtCurrency(subTotal)}?`)) return;

  goal.target = subTotal;
  // Update the input in the allocator row if visible
  const allocInput = document.querySelector(`#allocRow_${(typeof goalsData !== "undefined" ? goalsData : []).indexOf(goal)} .alloc-input`);
  if (allocInput) allocInput.value = subTotal;

  _reRenderGoalSection(goalName);
  _renderGoalsSummaryBar();
  if (typeof refreshGoalCardsGrid === "function") refreshGoalCardsGrid();
  if (typeof refreshGoalInsightPanels === "function") refreshGoalInsightPanels();
  if (typeof persistGoalsToExcel === "function") {
    const saved = await persistGoalsToExcel({ silent: true, includeBoosts: false });
    if (!saved) {
      alert("Target updated on this page, but Excel autosave failed. Check the autosave status and try again.");
      return;
    }
  }
  alert(`Target updated to ${_fmtCurrency(subTotal)}.`);
}

// ─── Autosave / Save All to Excel ─────────────────────────────────

function setGoalItemsAutoSaveStatus(message, tone = "") {
  const el = document.getElementById("goalItemsAutosaveStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "gi-autosave-status" + (tone ? " " + tone : "");
}

function scheduleGoalItemsAutoSave(delay = 700) {
  clearTimeout(_giAutoSaveTimer);
  setGoalItemsAutoSaveStatus("Saving soon...");
  _giAutoSaveTimer = setTimeout(() => {
    persistGoalItemsToExcel({ silent: true });
  }, delay);
}

async function persistGoalItemsToExcel(options = {}) {
  if (_giAutoSaveInFlight) {
    scheduleGoalItemsAutoSave(1000);
    return false;
  }

  const silent = options.silent === true;
  _giAutoSaveInFlight = true;

  try {
    setGoalItemsAutoSaveStatus("Saving...");
    await _writeGoalItemsSheet();
    setGoalItemsAutoSaveStatus("Saved to Excel", "ok");
    return true;
  } catch (err) {
    setGoalItemsAutoSaveStatus("Save failed", "error");
    if (!silent) alert("Save failed: " + err.message);
    console.error(err);
    return false;
  } finally {
    _giAutoSaveInFlight = false;
  }
}

async function saveAllGoalItems() {
  const btn = document.querySelector(".gi-save-btn");
  if (btn) { btn.disabled = true; btn.textContent = "⏳ Saving…"; }
  try {
    const saved = await persistGoalItemsToExcel({ silent: false });
    if (!saved) {
      if (btn) { btn.disabled = false; btn.textContent = "💾 Save to Excel"; }
      return;
    }
    if (btn) { btn.disabled = false; btn.textContent = "💾 Save to Excel"; }
    // Flash success
    const bar = document.getElementById("giHeaderSummary");
    if (bar) {
      const flash = document.createElement("span");
      flash.className = "gi-sum-pill ok";
      flash.textContent = "✓ Saved!";
      bar.prepend(flash);
      setTimeout(() => flash.remove(), 2500);
    }
  } catch (err) {
    if (btn) { btn.disabled = false; btn.textContent = "💾 Save to Excel"; }
    alert("Save failed: " + err.message);
    console.error(err);
  }
}

// ─── Excel write helpers ───────────────────────────────────────────

async function _writeGoalItemsSheet() {
  const token       = await getToken();
  const encodedPath = getEncodedExcelPath();

  // Build rows: header + data
  const headerRow = [GOAL_ITEMS_HEADERS];
  const dataRows  = _goalItems.map(i => [
    i.itemId,
    i.goalName,
    i.description,
    i.estAmount,
    i.requiredBy  || "",
    i.priority,
    i.link        || "",
    i.comments    || "",
    i.status,
    i.purchasedAmount || 0,
    i.purchasedDate   || "",
  ]);

  const allRows   = [...headerRow, ...dataRows];
  const rowCount  = allRows.length;
  const endCol    = "K"; // 11 columns (A–K)
  const writeRange = `A1:${endCol}${rowCount}`;

  // If the sheet doesn't exist yet, we need to create it first via Graph API
  if (!_giSheetExists) {
    await _ensureGoalItemsSheet(token, encodedPath);
    _giSheetExists = true;
  }

  // Write the data
  await writeExcelRange(GOAL_ITEMS_SHEET, writeRange, allRows);

  // Clear any stale rows below (in case items were deleted)
  // We'll write up to row 100 to clear
  if (rowCount < 100) {
    const clearRows = 100 - rowCount;
    const clearRange = `A${rowCount + 1}:K${rowCount + clearRows}`;
    const emptyRows  = Array(clearRows).fill(Array(11).fill(""));
    try { await writeExcelRange(GOAL_ITEMS_SHEET, clearRange, emptyRows); } catch(_) {}
  }
}

async function _ensureGoalItemsSheet(token, encodedPath) {
  // Create the sheet using Graph API
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/workbook/worksheets`;
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ name: GOAL_ITEMS_SHEET }),
  });
  if (!resp.ok && resp.status !== 409) { // 409 = already exists, that's fine
    const err = await resp.json().catch(() => ({}));
    throw new Error("Could not create Goal Items sheet: " + (err.error?.message || resp.status));
  }
}

async function _writeTransactionRow(date, description, amount, mainCategory, subCategory, account) {
  // Mirrors the pattern in saveDeductGoal() in goals.js and add-transaction.js
  const token = await getToken();
  const encodedPath = getEncodedExcelPath();
  const usedRangeUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/workbook/worksheets('${CONFIG.sheetName}')/usedRange(valuesOnly=true)`;
  const data    = await graphGetJson(usedRangeUrl, token);
  const nextRow = data.rowCount + 1;

  await writeExcelRange(CONFIG.sheetName, `A${nextRow}:F${nextRow}`, [[
    date, description, amount, mainCategory, subCategory, account
  ]]);
}

// ─── Re-render single section ──────────────────────────────────────

function _reRenderGoalSection(goalName) {
  const slug    = _slugify(goalName);
  const section = document.getElementById("giSection_" + slug);
  if (!section) return;
  const goalData = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === goalName);
  const items    = _goalItems.filter(i => i.goalName === goalName);
  const items2   = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === goalName);

  // Sort same as _buildGoalSectionHTML
  const priOrder = { Critical:0, High:1, Medium:2, Low:3 };
  const sorted   = [...items].sort((a, b) => {
    const aOverdue = a.status === "Planned" && a.requiredBy && a.requiredBy < _todayStr() ? 0 : 1;
    const bOverdue = b.status === "Planned" && b.requiredBy && b.requiredBy < _todayStr() ? 0 : 1;
    if (aOverdue !== bOverdue) return aOverdue - bOverdue;
    if (a.requiredBy && b.requiredBy) return a.requiredBy.localeCompare(b.requiredBy);
    return (priOrder[a.priority]??2) - (priOrder[b.priority]??2);
  });

  section.innerHTML = _buildGoalSectionHTML(goalName, goalData, sorted);
  _renderInsightsForGoal(goalName, items);
}

// ─── Utilities ────────────────────────────────────────────────────

function _generateId() {
  return "gi_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2,7);
}

function _slugify(str) {
  return String(str).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function _todayStr() {
  const d = new Date();
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

function _giDateToInput(value) {
  if (!value && value !== 0) return "";
  if (typeof value === "number" && value > 1000) {
    try { const d = XLSX.SSF.parse_date_code(value); return d.y + "-" + String(d.m).padStart(2,"0") + "-" + String(d.d).padStart(2,"0"); } catch { return ""; }
  }
  const s = String(value).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return s;
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return ddmm[3] + "-" + ddmm[2].padStart(2,"0") + "-" + ddmm[1].padStart(2,"0");
  return "";
}

function _inputDateToDisplay(iso) {
  if (!iso) return "";
  const [y,m,d] = iso.split("-");
  return d + "/" + m + "/" + y;
}

function _fmtCurrency(v) {
  return Number(v||0).toLocaleString("en-SG", { style:"currency", currency:"SGD", minimumFractionDigits:2, maximumFractionDigits:2 });
}

function _fmtShortCurrency(v) {
  const n = Number(v || 0);
  const abs = Math.abs(n);
  if (abs >= 1000000) return "$" + (n / 1000000).toFixed(abs >= 10000000 ? 0 : 1) + "m";
  if (abs >= 1000) return "$" + (n / 1000).toFixed(abs >= 10000 ? 0 : 1) + "k";
  return "$" + Math.round(n).toLocaleString("en-SG");
}

function escapeHtmlGI(v) {
  return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}

// ─── Insights Panel ────────────────────────────────────────────────

function _buildInsightsPanel(goalName, goalData, items) {
  const slug = _slugify(goalName);
  const activeItems = items.filter(i => i.status !== "Cancelled");

  if (activeItems.length === 0) {
    return `
      <div class="gi-insights-panel" id="giInsights_${slug}">
        <div class="gi-insights-header" onclick="toggleGIInsights('${slug}')">
          <span class="gi-insights-title">Target Breakdown</span>
          <span class="gi-insights-chevron" id="giInsightsChevron_${slug}">▾</span>
        </div>
        <div class="gi-insights-body" id="giInsightsBody_${slug}">
          <div class="gi-insight-no-dates">Add sub-items to see the goal target breakdown.</div>
        </div>
      </div>`;
  }

  return `
    <div class="gi-insights-panel" id="giInsights_${slug}">
      <div class="gi-insights-header" onclick="toggleGIInsights('${slug}')">
        <span class="gi-insights-title">Target Breakdown</span>
        <span class="gi-insights-chevron" id="giInsightsChevron_${slug}">▾</span>
      </div>
      <div class="gi-insights-body" id="giInsightsBody_${slug}">
        <div class="gi-insights-table-wrap" id="giInsightsTable_${slug}"></div>
        <div class="gi-insights-chart-wrap">
          <div class="gi-insights-chart-scroll">
            <div class="gi-insights-chart-box" id="giChartBox_${slug}">
              <canvas id="giChart_${slug}"></canvas>
            </div>
          </div>
        </div>
        <div class="gi-insights-note" id="giInsightsNote_${slug}"></div>
      </div>
    </div>`;
}

function _renderInsightsChart(goalName, items) {
  const slug = _slugify(goalName);
  const canvas = document.getElementById("giChart_" + slug);
  const chartBox = document.getElementById("giChartBox_" + slug);
  const tableWrap = document.getElementById("giInsightsTable_" + slug);
  const noteEl = document.getElementById("giInsightsNote_" + slug);
  if (!canvas) return;
  if (typeof Chart === "undefined") {
    if (tableWrap) tableWrap.innerHTML = `<div class="gi-insight-no-dates">Chart.js is unavailable.</div>`;
    return;
  }

  const avgSavings = (typeof historicalStats !== "undefined" ? historicalStats.avgMonthlySavings : 0) || 0;
  const today = new Date();
  const goalData = (typeof goalsData !== "undefined" ? goalsData : []).find(g => g.name === goalName);
  const goalTarget = goalData ? goalData.target * (1 + (goalData.goalBuffer || 0) / 100) : 0;

  const goalSavingsBalance = (() => {
    if (typeof savingsBalances === "undefined" || typeof goalSavingsAccts === "undefined") return 0;
    return goalSavingsAccts.reduce((s, a) => s + (savingsBalances[a] || 0), 0);
  })();

  const activeItems = items.filter(i => i.status !== "Cancelled");
  const plannedItems = activeItems
    .filter(i => i.status === "Planned" && i.requiredBy)
    .sort((a, b) => a.requiredBy.localeCompare(b.requiredBy));
  const purchasedAmt = activeItems
    .filter(i => i.status === "Purchased")
    .reduce((s, i) => s + (i.purchasedAmount || i.estAmount), 0);
  const subTotal = activeItems.reduce((s, i) => s + i.estAmount, 0);
  const plannedTotal = activeItems
    .filter(i => i.status === "Planned")
    .reduce((s, i) => s + i.estAmount, 0);

  if (activeItems.length === 0) return;

  // Per-item feasibility is still based on deadline order; the visual chart is sorted by target size.
  let cumulativeCost = purchasedAmt;
  const itemInsights = plannedItems.map(item => {
    cumulativeCost += item.estAmount;
    const deadline = new Date(item.requiredBy + "T00:00:00");
    const monthsLeft = Math.max(0,
      (deadline.getFullYear() - today.getFullYear()) * 12 +
      (deadline.getMonth() - today.getMonth())
    );
    const projectedAtDeadline = goalSavingsBalance + avgSavings * monthsLeft;
    const feasible = projectedAtDeadline >= cumulativeCost;
    const shortfall = feasible ? 0 : cumulativeCost - projectedAtDeadline;
    return { ...item, monthsLeft, feasible, shortfall, cumulativeCost, projectedAtDeadline };
  });
  const insightById = new Map(itemInsights.map(i => [i.itemId, i]));
  const priorityOrder = { Critical:0, High:1, Medium:2, Low:3 };
  const priorityBorders = { Critical:"#991b1b", High:"#ea580c", Medium:"#ca8a04", Low:"#16a34a" };
  const statusMetaFor = (item) => {
    if (item.status === "Purchased") return { label:"Purchased", className:"purchased", color:"#16a34a" };
    if (item.requiredBy && item.feasible === false) return { label:"At risk", className:"risk", color:"#dc2626" };
    if (item.requiredBy && item.feasible === true) return { label:"On track", className:"track", color:"#2563eb" };
    return { label:"No date", className:"nodate", color:"#64748b" };
  };

  const breakdownItems = activeItems.map(item => {
    const insight = insightById.get(item.itemId) || {};
    const requiredTime = item.requiredBy
      ? new Date(item.requiredBy + "T00:00:00").getTime()
      : Number.POSITIVE_INFINITY;
    return {
      ...item,
      ...insight,
      chartAmount: Math.max(0, item.estAmount || 0),
      requiredTime
    };
  }).sort((a, b) => {
    if (b.chartAmount !== a.chartAmount) return b.chartAmount - a.chartAmount;
    if (a.requiredTime !== b.requiredTime) return a.requiredTime - b.requiredTime;
    return (priorityOrder[a.priority] ?? 2) - (priorityOrder[b.priority] ?? 2);
  });

  // Destroy existing chart if any
  const existingChart = Chart.getChart(canvas);
  if (existingChart) existingChart.destroy();

  const onTrackCount = itemInsights.filter(i => i.feasible).length;
  const atRiskCount  = itemInsights.filter(i => !i.feasible).length;
  const noDateCount = activeItems.filter(i => i.status === "Planned" && !i.requiredBy).length;
  const purchasedCount = activeItems.filter(i => i.status === "Purchased").length;
  const atRiskRows  = itemInsights.filter(i => !i.feasible);
  const totalShortfall = atRiskRows.reduce((s, i) => s + i.shortfall, 0);
  const targetGap = goalTarget > 0 ? goalTarget - subTotal : 0;
  const denominator = Math.max(goalTarget, subTotal, 1);
  const largestItem = breakdownItems[0];

  const segmentHtml = breakdownItems
    .filter(i => i.chartAmount > 0)
    .map(item => {
      const meta = statusMetaFor(item);
      const pct = (item.chartAmount / denominator) * 100;
      return `<span class="gi-target-segment ${meta.className}" style="flex-basis:${pct.toFixed(4)}%;" title="${escapeHtmlGI(item.description)}: ${_fmtCurrency(item.chartAmount)}"></span>`;
    }).join("");
  const unallocatedHtml = goalTarget > subTotal
    ? `<span class="gi-target-segment unallocated" style="flex-basis:${((goalTarget - subTotal) / denominator * 100).toFixed(4)}%;" title="Unallocated: ${_fmtCurrency(goalTarget - subTotal)}"></span>`
    : "";
  const targetMarkerHtml = goalTarget > 0 && subTotal > goalTarget
    ? `<span class="gi-target-marker" style="left:${(goalTarget / denominator * 100).toFixed(3)}%;" title="Goal target: ${_fmtCurrency(goalTarget)}"></span>`
    : "";

  const targetStatus = goalTarget <= 0
    ? `<span class="gi-ins-pill neutral">No goal target set</span>`
    : targetGap >= 0
    ? `<span class="gi-ins-pill ok">${_fmtCurrency(targetGap)} unallocated</span>`
    : `<span class="gi-ins-pill risk">${_fmtCurrency(Math.abs(targetGap))} over target</span>`;
  const stripEndLabel = goalTarget > 0 && subTotal <= goalTarget
    ? "Target " + _fmtShortCurrency(goalTarget)
    : "Sub-items " + _fmtShortCurrency(subTotal);

  const overviewHtml = `
    <div class="gi-insights-summary-bar">
      <span class="gi-ins-pill neutral">${activeItems.length} sub-item${activeItems.length !== 1 ? "s" : ""}</span>
      <span class="gi-ins-pill neutral">Sub-items ${_fmtCurrency(subTotal)}</span>
      ${goalTarget > 0 ? `<span class="gi-ins-pill neutral">Goal target ${_fmtCurrency(goalTarget)}</span>` : ""}
      <span class="gi-ins-pill ok">${_fmtCurrency(purchasedAmt)} purchased</span>
      ${plannedTotal > 0 ? `<span class="gi-ins-pill track">${_fmtCurrency(plannedTotal)} planned</span>` : ""}
      ${targetStatus}
    </div>
    <div class="gi-target-strip-wrap">
      <div class="gi-target-strip-head">
        <span>Target allocation</span>
        <strong>${_fmtCurrency(subTotal)}${goalTarget > 0 ? " / " + _fmtCurrency(goalTarget) : ""}</strong>
      </div>
      <div class="gi-target-strip">
        ${segmentHtml || `<span class="gi-target-segment nodate" style="flex-basis:100%;"></span>`}
        ${unallocatedHtml}
        ${targetMarkerHtml}
      </div>
      <div class="gi-target-strip-axis">
        <span>$0</span>
        <span>${stripEndLabel}</span>
      </div>
    </div>
    <div class="gi-target-legend">
      <span><i class="gi-target-legend-dot purchased"></i>Purchased ${purchasedCount}</span>
      <span><i class="gi-target-legend-dot track"></i>On track ${onTrackCount}</span>
      <span><i class="gi-target-legend-dot risk"></i>At risk ${atRiskCount}</span>
      <span><i class="gi-target-legend-dot nodate"></i>No date ${noDateCount}</span>
    </div>
    ${largestItem ? `<div class="gi-insights-mini-note">Largest item: <strong>${escapeHtmlGI(largestItem.description)}</strong> at ${_fmtCurrency(largestItem.chartAmount)}${subTotal > 0 ? " (" + Math.round(largestItem.chartAmount / subTotal * 100) + "% of sub-items)" : ""}.</div>` : ""}
    ${atRiskRows.length > 0 ? `
    <details class="gi-ins-details">
      <summary class="gi-ins-details-summary">At-risk items (${atRiskCount})</summary>
      <table class="gi-ins-table">
        <tr><th>Item</th><th>Due</th><th>Cost</th><th>Projected</th><th>Shortfall</th></tr>
        ${atRiskRows.map(i => `
          <tr>
            <td>${escapeHtmlGI(i.description.length > 35 ? i.description.slice(0,35)+"..." : i.description)}</td>
            <td>${_inputDateToDisplay(i.requiredBy)}</td>
            <td>${_fmtCurrency(i.estAmount)}</td>
            <td>${_fmtCurrency(i.projectedAtDeadline)}</td>
            <td class="gi-ins-risk-cell">${_fmtCurrency(i.shortfall)}</td>
          </tr>`).join("")}
      </table>
    </details>` : ""}`;

  if (tableWrap) tableWrap.innerHTML = overviewHtml;
  if (chartBox) {
    const chartHeight = Math.min(2600, Math.max(270, breakdownItems.length * 30 + 90));
    chartBox.style.height = chartHeight + "px";
  }

  new Chart(canvas, {
    type: "bar",
    data: {
      labels: breakdownItems.map(i => i.description),
      datasets: [
        {
          label: "Estimated target",
          data: breakdownItems.map(i => i.chartAmount),
          backgroundColor: breakdownItems.map(i => statusMetaFor(i).color),
          borderColor: breakdownItems.map(i => priorityBorders[i.priority] || "#64748b"),
          borderWidth: 1,
          borderRadius: 5,
          borderSkipped: false,
          maxBarThickness: 18,
          datalabels: {
            anchor: "end",
            align: "right",
            clamp: true,
            offset: 6,
            color: "#374151",
            font: { size: 10, family: "DM Mono" },
            formatter: value => value > 0 ? _fmtShortCurrency(value) : ""
          }
        }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 250 },
      layout: { padding: { top: 4, right: 76, bottom: 0, left: 0 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: ctx => {
              const item = breakdownItems[ctx[0]?.dataIndex];
              return item ? item.description : "";
            },
            label: ctx => {
              const item = breakdownItems[ctx.dataIndex];
              if (!item) return "";
              const meta = statusMetaFor(item);
              const lines = [
                "Target: " + _fmtCurrency(item.chartAmount),
                "Status: " + meta.label,
                "Priority: " + (item.priority || "Medium")
              ];
              if (item.requiredBy) lines.push("Required by: " + _inputDateToDisplay(item.requiredBy));
              if (item.status === "Purchased") lines.push("Purchased: " + _fmtCurrency(item.purchasedAmount || item.estAmount));
              if (item.cumulativeCost != null) lines.push("Cumulative due cost: " + _fmtCurrency(item.cumulativeCost));
              if (item.projectedAtDeadline != null) lines.push("Projected savings: " + _fmtCurrency(item.projectedAtDeadline));
              if (item.shortfall > 0) lines.push("Shortfall: " + _fmtCurrency(item.shortfall));
              return lines;
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          grace: "8%",
          ticks: {
            callback: v => _fmtShortCurrency(v),
            maxTicksLimit: 8,
            font: { size: 11 }
          },
          grid: { color: "rgba(148,163,184,0.22)" }
        },
        y: {
          ticks: {
            autoSkip: false,
            callback: function(value) {
              const label = this.getLabelForValue ? this.getLabelForValue(value) : (breakdownItems[value]?.description || "");
              return label.length > 30 ? label.slice(0, 29) + "..." : label;
            },
            font: { size: 11 }
          },
          grid: { display: false }
        }
      }
    }
  });

  if (noteEl) noteEl.innerHTML =
    "Goal savings pool " + _fmtCurrency(goalSavingsBalance) +
    (avgSavings > 0 ? " + " + _fmtCurrency(avgSavings) + "/mo avg" : " (no savings history yet)") +
    " &nbsp;·&nbsp; Feasibility checks dated planned items against cumulative target cost" +
    (totalShortfall > 0 ? " &nbsp;·&nbsp; Total shortfall " + _fmtCurrency(totalShortfall) : "");
}

// Called after section renders to draw charts (canvas must be in DOM first)
function _renderInsightsForGoal(goalName, items) {
  // Use rAF to ensure canvas is painted
  requestAnimationFrame(() => _renderInsightsChart(goalName, items));
}

function toggleGIInsights(slug) {
  const body = document.getElementById("giInsightsBody_" + slug);
  const chev = document.getElementById("giInsightsChevron_" + slug);
  if (!body) return;
  const isOpen = body.style.display !== "none";
  body.style.display = isOpen ? "none" : "";
  if (chev) chev.textContent = isOpen ? "▸" : "▾";
}
