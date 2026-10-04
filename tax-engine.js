/* IRAS resident calculator YA 2026: H12:H55, verified 2026-10-02. */
(function(root) {
  const source = 'https://www.iras.gov.sg/quick-links/calculators';
  const reliefSource = 'https://www.iras.gov.sg/taxes/individual-income-tax/basics-of-individual-income-tax/tax-reliefs-rebates-and-deductions';
  const paymentSource = 'https://www.iras.gov.sg/taxes/individual-income-tax/basics-of-individual-income-tax/receive-tax-bill-pay-tax-check-refunds/late-payment-or-non-payment-of-individual-income-tax';
  const fields = [
    ['employment','Employment income','Annual gross pay, bonuses, allowances and taxable benefits for the income year. Use your IR8A; do not subtract employee CPF.','Income'],
    ['employmentExpenses','Allowable employment expenses','Unreimbursed expenses that qualify for an IRAS employment deduction. Exclude personal expenses and amounts paid back by your employer.','Income'],
    ['trade','Net trade / business income','Tax-adjusted sole-proprietor or partnership profit for the income year. Enter an allowable loss as a negative amount; not business sales or cash balance.','Income'],
    ['dividends','Taxable dividends','Only dividends taxable in Singapore. Exclude exempt Singapore one-tier dividends. Enter 0 if none.','Other income'],
    ['interest','Taxable interest','Only taxable interest. Exclude exempt interest from approved Singapore bank deposits. Enter 0 if none.','Other income'],
    ['rent','Net rental income','Your share of rent after allowable rental expenses, before personal reliefs. Enter 0 if none.','Other income'],
    ['royalties','Royalties, charges and estate / trust income','Taxable amount received from these sources during the income year. Enter 0 if none.','Other income'],
    ['otherIncome','Other taxable gains or income','Taxable income not entered above. Do not include the same income twice, or exempt capital gains.','Other income'],
    ['donations','Approved donation deduction','Enter the DEDUCTION shown by IRAS, normally 2.5 times eligible donations: a $100 qualifying donation gives $250 here. Do not multiply twice.','Deductions'],
    ['earned','Earned Income Relief','Enter your eligible relief: up to $1,000 below age 55, $6,000 at 55-59, or $8,000 from age 60, using age at income-year end. Disability limits are $4,000 / $10,000 / $12,000. Never exceed taxable earned income.','Personal reliefs'],
    ['spouse','Spouse Relief','Your eligible claim: $2,000, or $5,500 for disability. Check the qualifying conditions and spouse income limit with IRAS; otherwise enter 0.','Personal reliefs'],
    ['child','Qualifying Child Relief','Your agreed share of eligible relief: up to $4,000 per child, or $7,500 for disability. Exclude the share claimed by your spouse.','Personal reliefs'],
    ['workingMother','Working Mother\'s Child Relief','Eligible working mothers only. Enter your eligible amount after child-order, birth/adoption-date and combined child-relief limits; otherwise 0. Refer to IRAS for the applicable percentage or fixed amount.','Personal reliefs'],
    ['parent','Parent Relief','Your eligible share for up to two dependants. Usual per-dependant limits: $9,000 living with you / $5,500 apart; disability $14,000 / $10,000. Check support, income and sharing conditions.','Personal reliefs'],
    ['caregiver','Grandparent Caregiver Relief','Enter $3,000 only if you qualify as a working mother under IRAS conditions; otherwise 0.','Personal reliefs'],
    ['sibling','Sibling Relief (Disability)','Your eligible share for supporting a sibling with disability, up to $5,500 per dependant. Otherwise enter 0.','Personal reliefs'],
    ['cpf','CPF / provident fund relief','Your allowable employee or self-employed CPF relief for the income year, subject to CPF limits. Use IRAS / CPF records. Exclude employer CPF and cash top-ups entered below.','Personal reliefs'],
    ['life','Life Insurance Relief','Enter eligible relief, not your total premiums. Usually unavailable when allowable CPF contributions reach $5,000. Check policy, premium and sum-assured limits with IRAS.','Personal reliefs'],
    ['cpfTopup','CPF cash top-up relief','Eligible cash top-ups only: up to $8,000 for self plus $8,000 for eligible loved ones, shared with eligible MediSave top-ups. Exclude transfers between CPF accounts and ineligible matched top-ups.','Personal reliefs'],
    ['srs','SRS relief','Eligible contributions in the income year: up to $15,300 for citizens / PRs or $35,700 for foreigners. Use the applicable limit and exclude refunded contributions.','Personal reliefs'],
    ['nsman','NSman relief','Your eligible IRAS relief: self usually $1,500 / $3,000, or $3,500 / $5,000 for key appointment holders; wife / parent relief $750 where eligible. Use the amount granted to you.','Personal reliefs'],
    ['ptr','Available Parenthood Tax Rebate','Your unused PTR balance allocated to you, from myTax Portal. This is a tax rebate, not child relief. Do not re-enter rebate already used in an earlier year.','Tax rebate']
  ].map(([id,label,prompt,section]) => ({id,label,prompt,section}));
  const reliefIds = fields.filter(field => field.section === 'Personal reliefs').map(field => field.id);
  const brackets = [[20000,0],[30000,.02],[40000,.035],[80000,.07],[120000,.115],[160000,.15],[200000,.18],[240000,.19],[280000,.195],[320000,.20],[500000,.22],[1000000,.23],[Infinity,.24]];
  const money = value => Math.round((value + Number.EPSILON) * 100) / 100;
  function calculate(inputs) {
    const values = {};
    for (const field of fields) {
      if (field.id === 'employment' && (inputs[field.id] === '' || inputs[field.id] == null)) throw new Error('Enter annual employment income, including 0 if you had none.');
      const value = Number(inputs[field.id] || 0);
      if (!Number.isFinite(value) || (value < 0 && field.id !== 'trade')) throw new Error('Check ' + field.label + ': enter a valid annual amount.');
      values[field.id] = value;
    }
    const netEmployment = values.employment - values.employmentExpenses;
    const totalIncome = netEmployment + values.trade + values.dividends + values.interest + values.rent + values.royalties + values.otherIncome;
    const assessable = Math.max(0, totalIncome - values.donations);
    const enteredReliefs = reliefIds.reduce((sum,id) => sum + values[id], 0);
    const reliefs = Math.min(80000, enteredReliefs);
    const chargeable = Math.max(0, assessable - reliefs);
    let lower = 0, grossTax = 0;
    for (const [upper,rate] of brackets) {
      grossTax += Math.max(0, Math.min(chargeable,upper) - lower) * rate;
      lower = upper;
      if (chargeable <= upper) break;
    }
    grossTax = money(grossTax);
    const rebateUsed = Math.min(grossTax, values.ptr);
    return {netEmployment,totalIncome,assessable,enteredReliefs,reliefs,chargeable,grossTax,rebateUsed,rebateRemaining:money(values.ptr-rebateUsed),netTax:money(Math.max(0,grossTax-rebateUsed))};
  }
  function dueDate(billDate) {
    if (!billDate) return '';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(billDate)) throw new Error('Enter a valid tax bill date.');
    const [y,m,d] = billDate.split('-').map(Number);
    const date = new Date(Date.UTC(y,m-1,d));
    if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m-1 || date.getUTCDate() !== d) throw new Error('Enter a valid tax bill date.');
    const lastDay = new Date(Date.UTC(y,m+1,0)).getUTCDate();
    return new Date(Date.UTC(y,m,Math.min(d,lastDay))).toISOString().slice(0,10);
  }
  function newRecord(year, previous) {
    const inputs = Object.fromEntries(fields.map(field => [field.id,field.id === 'employment' ? '' : 0]));
    if (previous) Object.assign(inputs,previous.inputs,{donations:0,cpfTopup:0,ptr:calculate(previous.inputs).rebateRemaining});
    return {year,inputs,billDate:'',dueOverride:'',paid:false,confirmed:false,forecast:!!previous || year>2026};
  }
  function goalName(year) { return 'Personal Tax YA ' + year; }
  function deadline(record) { return record.dueOverride || dueDate(record.billDate); }
  function completeYear(state, paidDate) {
    const copy = JSON.parse(JSON.stringify(state));
    const record = copy.years[copy.activeYear];
    if (!record || record.paid) throw new Error('This tax year is already paid or unavailable.');
    if (!record.confirmed) throw new Error('Confirm the annual income and reliefs before marking this year paid.');
    record.paid = true; record.paidDate = paidDate; record.finalEstimate = calculate(record.inputs).netTax;
    const nextYear = Number(copy.activeYear) + 1;
    copy.years[nextYear] ||= newRecord(nextYear,record);
    copy.activeYear = nextYear;
    return copy;
  }
  root.FinTax = {fields,source,reliefSource,paymentSource,calculate,dueDate,deadline,newRecord,goalName,completeYear,money};
})(typeof window === 'undefined' ? globalThis : window);
