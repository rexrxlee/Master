(function(root) {
  function parse(value) {
    if (!value) { const year=+FinPropertyTax.today().slice(0,4)+1;return {version:1,revision:0,activeYear:year,linkedName:'',years:{[year]:FinPropertyTax.newRecord(year)}}; }
    const state = JSON.parse(String(value));
    if (state.version !== 1 || !state.years?.[state.activeYear]) throw new Error('The saved tax record cannot be read. No data was changed.');
    return state;
  }
  function prepare(state, rows, paid, today) {
    let next = JSON.parse(JSON.stringify(state));
    const record = next.years[next.activeYear];
    FinPropertyTax.calculate(record);
    if (record.dueOverride) FinPropertyTax.dueDate(record.dueOverride);
    const due = FinPropertyTax.deadline(record);
    if (record.billDate && due && due < record.billDate) throw new Error('Payment deadline must not be before the bill date.');
    let index = next.linkedName ? rows.findIndex(row => String(row[0] || '').trim() === next.linkedName) : -1;
    if (next.linkedName && index < 0) throw new Error('The linked goal was renamed or removed. Reload and select a goal.');
    if (index < 0) index = rows.findIndex(row => !String(row[0] || '').trim());
    if (index < 0) throw new Error('All 19 goal slots are in use. Free a slot before linking Property Tax.');
    let goal = Array.from({length:11}, (_,i) => rows[index]?.[i] ?? '');
    if (paid) {
      record.goalSnapshot = goal;
      next = FinPropertyTax.completeYear(next,today);
      next.linkedName = FinPropertyTax.goalName(next.activeYear);
      goal = [next.linkedName,0,0,0,today,'','Critical','','#2563eb',0,1];
    } else if (!next.linkedName) {
      next.linkedName = FinPropertyTax.goalName(next.activeYear);
      goal = [next.linkedName,0,0,0,today,'','Critical','','#2563eb',0,1];
    }
    if (rows.some((row,i) => i !== index && String(row[0] || '').trim() === next.linkedName)) throw new Error('A goal with this tax year name already exists. Select it instead.');
    const active = next.years[next.activeYear];
    goal[0] = next.linkedName;
    goal[1] = FinPropertyTax.calculate(active).netTax;
    goal[5] = FinPropertyTax.deadline(active);
    goal[7] = 'Managed by Property Tax. '+(FinPropertyTax.calculate(active).provisional?'Provisional savings target copied from previous bill; update AV or actual bill.':'Review estimate against your IRAS bill.');
    goal[9] = 0;
    next.revision = Number(state.revision || 0) + 1;
    return {state:next,index,goal};
  }
  async function save(state, expectedRevision, paid) {
    const live = await readBudgetSetupRange('S2:AM20');
    const rows = Array.from({length:19},(_,i) => live.values?.[i]?.slice(0,11) || []);
    const current = parse(live.values?.[0]?.[20]);
    if (current.pendingGoal) throw new Error('A previous tax save needs recovery. Reload the Tax page first.');
    const personal=JSON.parse(live.values?.[0]?.[18]||'{}'), insurance=JSON.parse(live.values?.[0]?.[19]||'{}');
    if(personal.pendingGoal||insurance.pending)throw Error('Recover the other managed goal save first.');
    if(state.linkedName&&(state.linkedName===personal.linkedName||state.linkedName===insurance.linkedName))throw Error('That goal is already managed by Tax or Insurance.');
    if (Number(current.revision || 0) !== Number(expectedRevision || 0)) throw new Error('Tax settings changed in another page. Reload before saving.');
    const today = new Date().toLocaleDateString('en-CA',{timeZone:'Asia/Singapore'});
    const result = prepare(state,rows,paid,today);
    // Metadata is authoritative; a retry repairs a partially completed goal update.
    result.state.pendingGoal = {index:result.index,values:result.goal,previous:rows[result.index]};
    const completed = JSON.parse(JSON.stringify(result.state));
    delete completed.pendingGoal;
    await writeExcelRanges([
      {sheetName:'Budget Setup',rangeAddress:'AM2',values:[[JSON.stringify(result.state)]]},
      {sheetName:'Budget Setup',rangeAddress:`S${result.index+2}:AC${result.index+2}`,values:[result.goal]},
      {sheetName:'Budget Setup',rangeAddress:'AM2',values:[[JSON.stringify(completed)]]}
    ]);
    return completed;
  }
  async function recover() {
    const live = await readBudgetSetupRange('S2:AM20');
    const state = parse(live.values?.[0]?.[20]);
    const pending = state.pendingGoal;
    if (!pending) return state;
    const equal = (a,b) => Array.from({length:11},(_,i) => String(a?.[i] ?? '')).join('\u001f') === Array.from({length:11},(_,i) => String(b?.[i] ?? '')).join('\u001f');
    const row = live.values?.[pending.index];
    if (!equal(row,pending.previous) && !equal(row,pending.values)) throw new Error('The goal changed during an interrupted tax save. No goal was overwritten. Resolve the conflicting goal before retrying.');
    delete state.pendingGoal;
    await writeExcelRanges([
      {sheetName:'Budget Setup',rangeAddress:`S${pending.index+2}:AC${pending.index+2}`,values:[pending.values]},
      {sheetName:'Budget Setup',rangeAddress:'AM2',values:[[JSON.stringify(state)]]}
    ]);
    return state;
  }
  root.FinPropertyTaxStore = {parse,prepare,save,recover};
})(typeof window === 'undefined' ? globalThis : window);
