let taxState, taxRows = [], taxRevision = 0, taxDirty = false, taxSaving = false;
const taxEl = id => document.getElementById(id);
const taxMoney = value => new Intl.NumberFormat('en-SG',{style:'currency',currency:'SGD'}).format(value);
const taxEscape = value => String(value ?? '').replace(/[&<>"']/g,c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

async function loadTax(force = false, discard = false) {
  if (taxSaving) return;
  if (taxDirty && !discard) { taxEl('taxReloadConfirm').hidden = false; taxEl('taxReloadConfirm').scrollIntoView({block:'center'}); return; }
  taxEl('taxReloadConfirm').hidden = true;
  taxEl('taxMessage').textContent = 'Loading tax settings...';
  taxEl('taxForm').hidden = true;
  try {
    let workbook = XLSX.read(await downloadExcelFile(force),{type:'array'});
    let sheet = workbook.Sheets['Budget Setup'];
    if (!sheet) throw new Error('Budget Setup sheet was not found.');
    taxState = FinTaxStore.parse(sheet.AK2?.v);
    if (taxState.pendingGoal) {
      await FinTaxStore.recover();
      workbook = XLSX.read(await downloadExcelFile(true),{type:'array'});
      sheet = workbook.Sheets['Budget Setup'];
      taxState = FinTaxStore.parse(sheet.AK2?.v);
    }
    taxRows = XLSX.utils.sheet_to_json(sheet,{header:1,range:'S2:AC20',blankrows:true});
    taxRevision = taxState.revision;
    taxDirty = false;
    renderTax();
    taxEl('taxMessage').textContent = 'Based on the IRAS YA 2026 resident calculator. Later years are estimates using these rules until IRAS publishes that year’s rules. Not a tax assessment.';
  } catch (error) { taxEl('taxMessage').textContent = 'Unable to load: ' + error.message; }
}

function renderTax() {
  const year = Number(taxState.activeYear), record = taxState.years[year];
  taxEl('taxForm').hidden = false;
  const years = [...new Set([year,2026,2027,new Date().getFullYear()+1])].sort();
  taxEl('taxYear').innerHTML = years.filter(y => y >= 2026).map(y => `<option value="${y}" ${y===year?'selected':''}>YA ${y} (income ${y-1})</option>`).join('');
  taxEl('taxYear').disabled = !!taxState.linkedName;
  taxEl('taxBasis').textContent = `Income earned 1 Jan to 31 Dec ${year-1}. ${record.forecast ? 'Forecast: review copied amounts and future tax rules.' : 'YA 2026 calculator rules.'}`;
  taxEl('taxGoal').innerHTML = '<option value="">Create Personal Tax goal</option>' + taxRows.filter(r => r[0] && (/tax|iras/i.test(r[0]) || r[0] === taxState.linkedName)).map(r => `<option value="${taxEscape(r[0])}">${taxEscape(r[0])}</option>`).join('');
  const linkedExists = taxRows.some(row => String(row[0] || '').trim() === taxState.linkedName);
  if (taxState.linkedName && !linkedExists) taxEl('taxGoal').insertAdjacentHTML('beforeend',`<option value="${taxEscape(taxState.linkedName)}">Missing goal: ${taxEscape(taxState.linkedName)} - choose a replacement</option>`);
  taxEl('taxGoal').value = taxState.linkedName || '';
  taxEl('taxGoal').disabled = !!taxState.linkedName && linkedExists;
  const sections = [...new Set(FinTax.fields.map(f => f.section))];
  taxEl('taxFields').innerHTML = sections.map(section => `<fieldset><legend>${section}</legend><div class="tax-grid">${FinTax.fields.filter(f => f.section === section).map(f => `<label for="tax_${f.id}">${taxEscape(f.label)}<input id="tax_${f.id}" type="number" step="0.01" ${f.id==='trade'?'':'min="0"'} ${f.id==='employment'?'required':''} value="${taxEscape(record.inputs[f.id])}" aria-describedby="help_${f.id}"><small id="help_${f.id}">${taxEscape(f.prompt)}</small></label>`).join('')}</div></fieldset>`).join('');
  taxEl('taxBillDate').value = record.billDate;
  taxEl('taxDueOverride').value = record.dueOverride;
  taxEl('taxConfirmed').checked = !!record.confirmed;
  taxEl('taxPaid').checked = false;
  taxEl('taxSaveStatus').textContent = '';
  const paid = Object.values(taxState.years).filter(r => r.paid);
  taxEl('taxHistory').innerHTML = paid.length ? '<h2>Paid tax years</h2><table><thead><tr><th>Year</th><th>Estimate</th><th>Paid</th></tr></thead><tbody>' + paid.map(r => `<tr><td>YA ${Number(r.year)}</td><td>${taxMoney(r.finalEstimate)}</td><td>${taxEscape(r.paidDate)}</td></tr>`).join('') + '</tbody></table>' : '';
  previewTax();
}

function collectTax() {
  const draft = JSON.parse(JSON.stringify(taxState));
  const record = draft.years[draft.activeYear];
  for (const f of FinTax.fields) record.inputs[f.id] = taxEl('tax_'+f.id).value;
  record.billDate = taxEl('taxBillDate').value;
  record.dueOverride = taxEl('taxDueOverride').value;
  record.confirmed = taxEl('taxConfirmed').checked;
  draft.linkedName = taxEl('taxGoal').value;
  return draft;
}

function previewTax() {
  try {
    const draft = collectTax(), record = draft.years[draft.activeYear];
    const deadline = FinTax.deadline(record);
    taxEl('taxDeadline').textContent = deadline ? `Payment deadline: ${deadline}` : 'Awaiting tax bill: no official payment deadline yet.';
    const result = FinTax.calculate(record.inputs);
    taxEl('taxResult').innerHTML = `<dl><dt>Assessable income</dt><dd>${taxMoney(result.assessable)}</dd><dt>Personal reliefs (capped at $80,000)</dt><dd>${taxMoney(result.reliefs)}</dd><dt>Chargeable income</dt><dd>${taxMoney(result.chargeable)}</dd><dt>Tax before rebate</dt><dd>${taxMoney(result.grossTax)}</dd><dt>Parenthood Tax Rebate used</dt><dd>${taxMoney(result.rebateUsed)}</dd><dt class="tax-total">Personal Tax goal</dt><dd class="tax-total">${taxMoney(result.netTax)}</dd></dl>`;
  } catch(error) { taxEl('taxResult').textContent = error.message; }
}

taxEl('taxForm').addEventListener('input',() => {
  taxDirty = true;
  taxEl('taxSaveStatus').textContent = 'Unsaved changes';
  previewTax();
});
taxEl('taxYear').addEventListener('change',() => {
  const draft = collectTax();
  const year = Number(taxEl('taxYear').value);
  draft.activeYear = year;
  draft.years[year] ||= FinTax.newRecord(year);
  taxState = draft;
  renderTax();
  taxDirty = true;
});
taxEl('taxForm').addEventListener('submit',async event => {
  event.preventDefault();
  if (taxSaving) return;
  taxSaving = true;
  taxEl('taxSave').disabled = true;
  // Freeze the draft while writing so an in-flight edit cannot be lost.
  const controls = [...taxEl('taxForm').querySelectorAll('input,select')];
  const oldDisabled = controls.map(el => el.disabled);
  let saved = false;
  try {
    const draft = collectTax();
    const paid = taxEl('taxPaid').checked;
    controls.forEach(el => { el.disabled = true; });
    taxEl('taxSaveStatus').textContent = 'Saving calculator and linked goal...';
    taxState = await FinTaxStore.save(draft,taxRevision,paid);
    taxRevision = taxState.revision;
    taxDirty = false;
    saved = true;
    const rowIndex = taxRows.findIndex(row => row[0] === draft.linkedName);
    if (rowIndex >= 0) taxRows[rowIndex][0] = taxState.linkedName;
    else taxRows.push([taxState.linkedName]);
    renderTax();
    taxEl('taxSaveStatus').textContent = paid ? `Saved as paid. YA ${taxState.activeYear} forecast is now linked to Goals. Review its figures when available.` : 'Saved to Excel and linked to Goals.';
  } catch(error) {
    taxEl('taxSaveStatus').textContent = 'Save could not finish: ' + error.message + ' Reload to check or recover the saved record before retrying.';
  } finally {
    controls.forEach((el,i) => { el.disabled = saved && (el.id === 'taxYear' || el.id === 'taxGoal') ? true : oldDisabled[i]; });
    taxEl('taxSave').disabled = false;
    taxSaving = false;
  }
});
window.addEventListener('beforeunload',event => { if (taxDirty || taxSaving) { event.preventDefault(); event.returnValue = ''; } });
