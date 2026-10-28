let ptState,ptRows=[],ptBusy=false,ptDirty=false;
const ptEl=id=>document.getElementById(id);
async function loadPropertyTax(){
  if(ptBusy)return;
  if(ptDirty){ptEl('ptMessage').textContent='Unsaved changes: save first, or refresh the page to discard.';return;}
  ptBusy=true;ptEl('ptForm').hidden=true;
  try{
    let w=XLSX.read(await downloadExcelFile(true),{type:'array'}),s=w.Sheets['Budget Setup'];
    if(!s)throw Error('Budget Setup is missing.');
    ptState=FinPropertyTaxStore.parse(s.AM2?.v);
    if(ptState.pendingGoal){await FinPropertyTaxStore.recover();w=XLSX.read(await downloadExcelFile(true),{type:'array'});s=w.Sheets['Budget Setup'];ptState=FinPropertyTaxStore.parse(s.AM2?.v);}
    ptRows=XLSX.utils.sheet_to_json(s,{header:1,range:'S2:AC20',blankrows:true});
    renderPropertyTax();ptEl('ptMessage').textContent='Ready. New setup defaults to next year; select the current year for an unpaid bill.';
  }catch(e){ptEl('ptMessage').textContent=e.message;}finally{ptBusy=false;}
}
function renderPropertyTax(){
  const r=ptState.years[ptState.activeYear];
  for(const [id,v] of Object.entries({ptYear:r.year,ptAV:r.av,ptActual:r.actual,ptDue:r.dueOverride}))ptEl(id).value=v??'';
  ptEl('ptOwner').checked=!!r.ownerOccupied;ptEl('ptFirst').checked=!!r.firstBill;ptEl('ptPaid').checked=false;
  ptEl('ptYear').disabled=!!ptState.linkedName;
  ptEl('ptGoal').innerHTML='<option value="">Create Property Tax goal</option>'+ptRows.filter(r=>r[0]&&(/property.*tax/i.test(r[0])||r[0]===ptState.linkedName)).map(r=>`<option>${taxEscape(r[0])}</option>`).join('');
  if(ptState.linkedName&&!ptRows.some(r=>r[0]===ptState.linkedName))ptEl('ptGoal').insertAdjacentHTML('beforeend',`<option>${taxEscape(ptState.linkedName)}</option>`);
  ptEl('ptGoal').value=ptState.linkedName||'';ptEl('ptGoal').disabled=!!ptState.linkedName;
  ptEl('ptForm').hidden=false;
  ptEl('ptHistory').innerHTML=Object.values(ptState.years).filter(r=>r.paid).map(r=>`<p>${Number(r.year)}: ${taxMoney(r.finalEstimate)} marked paid ${taxEscape(r.paidDate)}</p>`).join('');
  previewPropertyTax();
}
function collectPropertyTax(){
  const d=structuredClone(ptState),year=+ptEl('ptYear').value;
  if(d.linkedName&&year!==+d.activeYear)throw Error('Use annual rollover to change a linked year.');
  d.activeYear=year;d.linkedName=ptEl('ptGoal').value;
  d.years[year]={...FinPropertyTax.newRecord(year),...d.years[year],year,av:ptEl('ptAV').value,actual:ptEl('ptActual').value,dueOverride:ptEl('ptDue').value,ownerOccupied:ptEl('ptOwner').checked,firstBill:ptEl('ptFirst').checked};
  return d;
}
function previewPropertyTax(){try{
  const d=collectPropertyTax(),r=d.years[d.activeYear],v=FinPropertyTax.calculate(r),due=FinPropertyTax.deadline(r);
  ptEl('ptResult').textContent=(v.provisional?'Provisional target from previous bill (review required): ':v.actual?'Actual net bill: ':`Tax before rebate: ${taxMoney(v.grossTax)}. Rebate: ${taxMoney(v.rebate)}. Estimated annual tax: `)+taxMoney(v.netTax)+`. Annual amount ÷ 12: ${taxMoney(v.netTax/12)} (not a GIRO schedule). Goal deadline: ${due||'awaiting first bill'}. Goals allocates savings using your available funds and other goals.`;
}catch(e){ptEl('ptResult').textContent=e.message;}}
ptEl('ptForm').addEventListener('input',()=>{ptDirty=true;previewPropertyTax();});
ptEl('ptForm').addEventListener('submit',async e=>{
  e.preventDefault();if(ptBusy)return;ptBusy=true;
  const controls=[...ptEl('ptForm').querySelectorAll('input,select,button')],disabled=controls.map(c=>c.disabled);
  try{
    const d=collectPropertyTax(),paid=ptEl('ptPaid').checked;
    FinPropertyTax.calculate(d.years[d.activeYear]);
    controls.forEach(c=>c.disabled=true);
    ptState=await FinPropertyTaxStore.save(d,ptState.revision,paid);ptDirty=false;
    ptEl('ptMessage').textContent=paid?`Paid year saved. ${ptState.activeYear} forecast opened; review Annual Value and future rates.`:'Property Tax saved and linked to Goals.';
    const row=ptRows.find(r=>r[0]===d.linkedName);if(row)row[0]=ptState.linkedName;else ptRows.push([ptState.linkedName]);
  }catch(err){ptEl('ptMessage').textContent=err.message+' If a save was interrupted, refresh Tax to recover before retrying.';}
  finally{controls.forEach((c,i)=>c.disabled=disabled[i]);ptBusy=false;if(!ptDirty)renderPropertyTax();}
});
window.addEventListener('beforeunload',e=>{if(ptDirty||ptBusy){e.preventDefault();e.returnValue='';}});
