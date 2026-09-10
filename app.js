const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const dialog = $('#operator-dialog');
const toast = $('#toast');
const state = { operatorId:'', conditions:[], editing:null, saving:false, saved:false };
const conditionFields = ['productCode','jumboNo','slitNo','width','speed','ramp','tension','pressure','torque','rollLength','rollDiameter','crossSection'];

function notify(message){ toast.textContent=message; toast.classList.add('show'); setTimeout(()=>toast.classList.remove('show'),3400); }
function escapeHtml(value){return String(value).replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));}
function showView(id){ $$('.view').forEach(view=>view.classList.toggle('active',view.id===id)); window.scrollTo({top:0,behavior:'smooth'}); }
function documentNo(){ const d=$('[name="recordDate"]').value,s=$('[name="shift"]').value,m=$('[name="machine"]').value; return d&&s&&m?`${d.replaceAll('-','')}${s==='Day'?'D':'N'}-${m}`:''; }
function refreshDocumentNo(){ $('#doc-number').textContent=documentNo()||'เลือกวันที่และกะ'; }
function goStep(step){
  if((state.saving||state.saved)&&step!=='success')return;
  $$('.step-panel').forEach(p=>p.classList.toggle('active',p.dataset.panel===step));
  const order=['setup','inspection','condition','confirm','success'], target=order.indexOf(step);
  $$('.step').forEach((button,index)=>{button.classList.toggle('active',index===target);button.classList.toggle('done',index<target);});
  window.scrollTo({top:70,behavior:'smooth'});
}
function enableThrough(step){ const order=['setup','inspection','condition','confirm','success'],end=order.indexOf(step); $$('.step').forEach((b,i)=>b.disabled=i>end); }

$('#home-button').addEventListener('click',()=>{if(!state.saving)showView('home-view');});
$('[data-open="operator-dialog"]').addEventListener('click',()=>dialog.showModal());
$('#operator-login').addEventListener('click',()=>{
  const employeeId=dialog.querySelector('[name="employeeId"]').value.trim(),pin=dialog.querySelector('[name="pin"]').value.trim();
  if(!employeeId||!pin)return notify('กรุณากรอก Employee ID และ PIN');
  state.operatorId=employeeId; dialog.querySelector('[name="pin"]').value='';
  $('#operator-name').textContent=`ID ${employeeId}`; $('#operator-exit').hidden=false; dialog.close(); showView('operator-view');
});
$('#operator-exit').addEventListener('click',()=>{if(!state.saving){$('#operator-exit').hidden=true;showView('home-view');}});
$$('.step').forEach(button=>button.addEventListener('click',()=>!button.disabled&&goStep(button.dataset.step)));
$$('.back-step').forEach(button=>button.addEventListener('click',()=>goStep(button.dataset.back)));

$$('[name="recordDate"],[name="shift"],[name="machine"]').forEach(control=>control.addEventListener('change',refreshDocumentNo));
$('[name="area"]').addEventListener('change',event=>{const values=event.target.value==='TH3'?['3A','3B']:['4A','4B']; $('[name="machine"]').innerHTML=values.map(v=>`<option>${v}</option>`).join('');refreshDocumentNo();});
$('[name="team"]').addEventListener('change',event=>{$('#operator-team').textContent=`ทีม ${event.target.value}`;});
$$('[name="recordType"]').forEach(radio=>radio.addEventListener('change',()=>{$$('.type-option').forEach(label=>label.classList.toggle('selected',label.querySelector('input').checked));}));
$('#start-inspection').addEventListener('click',()=>{
  for(const input of $$('[data-panel="setup"] [required]'))if(!input.checkValidity()){input.reportValidity();return;}
  if($('[name="recordType"]:checked').value!=='Slitter')return notify('Auto Pack Record จะเชื่อมในขั้นตอนถัดไป');
  enableThrough('inspection');goStep('inspection');
});
$('#complete-inspection').addEventListener('click',()=>{
  const checks=$$('#machine-checks input,.dolly-grid input[type="checkbox"]'),raw=$('#dolly-cuts').value,cuts=raw===''?NaN:Number(raw);
  if(!checks.every(c=>c.checked)||!Number.isInteger(cuts)||cuts<0||cuts>=400)return notify('ต้องตรวจครบทุกข้อ และจำนวนครั้งตัดต้องน้อยกว่า 400');
  enableThrough('condition');goStep('condition');
});

const knifeGrid=$('#knife-grid');
for(let number=1;number<=50;number++){const v=String(number).padStart(2,'0');knifeGrid.insertAdjacentHTML('beforeend',`<label><input type="checkbox" value="${v}"><span>${v}</span></label>`);}
$('#select-all-knives').addEventListener('click',()=>$$('#knife-grid input').forEach(i=>i.checked=true));
$('#clear-knives').addEventListener('click',()=>$$('#knife-grid input').forEach(i=>i.checked=false));
$('#copy-condition').addEventListener('change',event=>{if(event.target.value==='')return;const source=state.conditions[Number(event.target.value)];if(!source)return;$$('#knife-grid input').forEach(i=>i.checked=source.knives.includes(i.value));notify(`คัดลอกใบมีดจาก Condition ${Number(event.target.value)+1} แล้ว`);});

function addReplacement(value={}){
  $('#replacement-list .empty-row')?.remove();const row=document.createElement('div');row.className='replacement-row';
  row.innerHTML='<label>เบอร์ใบมีด<input class="replacement-knife" type="number" min="1" max="50" required></label><label>อายุการใช้งาน (ชม.)<input class="replacement-life" type="number" min="0" required></label><label>สาเหตุ<input class="replacement-reason" required></label><button type="button">ลบ</button>';
  row.querySelector('.replacement-knife').value=value.knifeNo||'';row.querySelector('.replacement-life').value=value.lifeHours??'';row.querySelector('.replacement-reason').value=value.reason||'';
  row.querySelector('button').addEventListener('click',()=>{row.remove();if(!$('.replacement-row'))$('#replacement-list').innerHTML='<p class="empty-row">ไม่มีการเปลี่ยนใบมีดใน Condition นี้</p>';});$('#replacement-list').append(row);
}
$('#add-replacement').addEventListener('click',()=>addReplacement());
function clearEditor(){
  conditionFields.forEach(name=>{const field=$(`[name="${name}"]`);field.value='';});$$('#knife-grid input').forEach(i=>i.checked=false);$('#replacement-list').innerHTML='<p class="empty-row">ไม่มีการเปลี่ยนใบมีดใน Condition นี้</p>';state.editing=null;$('#condition-label').textContent=`CONDITION ${state.conditions.length+1}`;$('#cancel-condition').hidden=true;$('#condition-editor').hidden=false;
}
function readCondition(){
  for(const field of $$('#condition-editor input,#condition-editor select'))if(!field.checkValidity()){field.reportValidity();throw new Error('กรุณากรอกข้อมูล Condition ให้ครบและถูกต้อง');}
  const knives=$$('#knife-grid input:checked').map(i=>i.value);if(!knives.length)throw new Error('กรุณาเลือกใบมีดอย่างน้อย 1 ใบ');
  const data={};conditionFields.forEach(name=>data[name]=$(`[name="${name}"]`).value.trim());data.productCode=data.productCode.toUpperCase();data.knives=knives;
  data.replacements=$$('.replacement-row').map(row=>{const knifeField=row.querySelector('.replacement-knife'),lifeField=row.querySelector('.replacement-life'),reasonField=row.querySelector('.replacement-reason'),knifeNo=Number(knifeField.value),life=Number(lifeField.value),reason=reasonField.value.trim();if(!knifeField.checkValidity()||!lifeField.checkValidity()||!reasonField.checkValidity()||!reason)throw new Error('กรุณากรอกข้อมูลการเปลี่ยนใบมีดให้ครบ');return{knifeNo,lifeHours:life,reason};});return data;
}
function renderConditions(){
  $('#condition-cards').innerHTML=state.conditions.map((c,i)=>`<article class="condition-card"><span class="card-index">CONDITION ${i+1}</span><h3>${escapeHtml(c.productCode)}</h3><p>Jumbo ${escapeHtml(c.jumboNo)} · Slit ${escapeHtml(c.slitNo)}</p><p>${c.knives.length} ใบมีด</p><div class="card-tools"><button type="button" data-edit="${i}">แก้ไข</button><button class="delete" type="button" data-delete="${i}">ลบ</button></div></article>`).join('');
  $$('[data-edit]').forEach(b=>b.addEventListener('click',()=>editCondition(Number(b.dataset.edit))));
  $$('[data-delete]').forEach(b=>b.addEventListener('click',()=>{state.conditions.splice(Number(b.dataset.delete),1);renderConditions();clearEditor();}));
  $('#copy-condition').innerHTML='<option value="">คัดลอกจาก Condition</option>'+state.conditions.map((_,i)=>`<option value="${i}">Condition ${i+1}</option>`).join('');
  $('#add-condition').hidden=!state.conditions.length||state.conditions.length>=6;$('#review-record').disabled=!state.conditions.length;
}
function editCondition(index){const c=state.conditions[index];state.editing=index;conditionFields.forEach(name=>$(`[name="${name}"]`).value=c[name]??'');$$('#knife-grid input').forEach(i=>i.checked=c.knives.includes(i.value));$('#replacement-list').innerHTML='<p class="empty-row">ไม่มีการเปลี่ยนใบมีดใน Condition นี้</p>';c.replacements.forEach(addReplacement);$('#condition-label').textContent=`แก้ไข CONDITION ${index+1}`;$('#cancel-condition').hidden=false;$('#condition-editor').hidden=false;}
$('#cancel-condition').addEventListener('click',clearEditor);
$('#save-condition').addEventListener('click',()=>{try{const condition=readCondition();if(state.editing===null)state.conditions.push(condition);else state.conditions[state.editing]=condition;renderConditions();$('#condition-editor').hidden=true;notify('บันทึก Condition แล้ว');}catch(error){notify(error.message);}});
$('#add-condition').addEventListener('click',()=>{if(state.conditions.length<6)clearEditor();});

function buildConfirmation(){
  const checks=$$('#machine-checks input,.dolly-grid input[type="checkbox"]');
  $('#confirm-summary').innerHTML=[['Document No.',documentNo()],['Area / Machine',`${$('[name="area"]').value} / ${$('[name="machine"]').value}`],['Operator',state.operatorId],['ทีม / กะ',`${$('[name="team"]').value} / ${$('[name="shift"]').value}`],['ผ่าน',checks.filter(c=>c.checked).length+1],['ไม่ผ่าน','0'],['Condition',state.conditions.length],['สถานะ','พร้อมบันทึก']].map(([k,v])=>`<div class="summary-item"><span>${escapeHtml(k)}</span><b>${escapeHtml(v)}</b></div>`).join('');
  $('#confirm-conditions').innerHTML=state.conditions.map((c,i)=>`<div class="confirm-row"><b>#${i+1}</b><span>${escapeHtml(c.productCode)}</span><span>Jumbo ${escapeHtml(c.jumboNo)}</span><span>${c.knives.length} ใบมีด</span></div>`).join('');
}
$('#review-record').addEventListener('click',()=>{buildConfirmation();enableThrough('confirm');goStep('confirm');});

function collectGroups(){
  const setup=Object.fromEntries(new FormData($('#record-form'))),docNo=documentNo(),now=new Date().toISOString(),revision=1,common={docNo,revision,area:setup.area,operatorId:state.operatorId},inspectionKey=`${docNo}-R${revision}`;
  const checks=$$('#machine-checks input,.dolly-grid input[type="checkbox"]'),machineCount=$$('#machine-checks input').length;
  const results=checks.map((input,index)=>{const category=input.closest('.dolly-grid')?'Dolly':'Machine',no=category==='Dolly'?index-machineCount+1:index+1,itemCode=`${category==='Dolly'?'D':'M'}${String(no).padStart(2,'0')}`;return{...common,resultKey:`${inspectionKey}-${itemCode}`,inspectionKey,category,itemCode,checklistName:input.parentElement.textContent.trim(),inputType:'PassFail',resultText:'Pass',resultStatus:'Pass',checkedAt:now,remark:setup.inspectionRemark.trim()};});
  results.push({...common,resultKey:`${inspectionKey}-D04`,inspectionKey,category:'Dolly',itemCode:'D04',checklistName:'จำนวนครั้งตัด Paper Core Cutter',inputType:'Number',resultNumber:Number($('#dolly-cuts').value),resultStatus:'Pass',checkedAt:now,remark:setup.inspectionRemark.trim()});
  const production=[],conditionValues=[],knifeRows=[],replacementRows=[];
  state.conditions.forEach((c,index)=>{
    const round=index+1,num=key=>c[key]===''?null:Number(c[key]);
    production.push({docNo,revision,isLatest:false,recordDate:setup.recordDate,shift:setup.shift,team:setup.team,operatorId:state.operatorId,productCode:c.productCode,jumboNo:c.jumboNo,conditionRound:round,slitNo:c.slitNo,width:num('width'),speed:num('speed'),ramp:c.ramp,tension:c.tension,pressure:c.pressure,torque:c.torque,rollLength:num('rollLength'),rollDiameter:num('rollDiameter')});
    const params=[['WIDTH','Width on Jumbo','Number',num('width'),null,'mm'],['SPEED','Speed','Number',num('speed'),null,'m/min'],['RAMP','Ramp Up / Down','Text',null,c.ramp,''],['TENSION','Winder Tension','Number',num('tension'),null,''],['PRESSURE','Rider Roll Pressure','Number',num('pressure'),null,''],['TORQUE','Torque','Number',num('torque'),null,''],['ROLL_LENGTH','Roll Length','Number',num('rollLength'),null,'m'],['ROLL_DIAMETER','Roll Diameter','Number',num('rollDiameter'),null,'mm'],['CROSS_SECTION','ตรวจสอบหน้าตัดม้วน','Choice',null,c.crossSection,'']];
    params.filter(([, , ,n,t])=>n!==null||t).forEach(([code,name,dataType,valueNumber,valueText,unit],p)=>{const recordKey=`${docNo}-C${round}-${String(p+1).padStart(2,'0')}`;conditionValues.push({...common,title:recordKey,recordKey,productCode:c.productCode,jumboNo:c.jumboNo,conditionRound:String(round),parameterCode:code,parameterName:name,dataType,valueNumber,valueText,unit,recordedAt:now});});
    c.knives.forEach(value=>{const knifeNo=Number(value),recordKey=`${docNo}-C${round}-K${value}`;knifeRows.push({...common,title:recordKey,recordKey,productCode:c.productCode,jumboNo:c.jumboNo,widthPattern:c.width,knifeNo,recordedAt:now});});
    c.replacements.forEach((r,i)=>{const recordKey=`${docNo}-C${round}-R${i+1}`;replacementRows.push({...common,title:recordKey,recordKey,productCode:c.productCode,jumboNo:c.jumboNo,knifeNo:r.knifeNo,lifeHours:r.lifeHours,lifeUnit:'Hours',reason:r.reason,replacedAt:now});});
  });
  return[{list:'ProductionRecords',rows:production},{list:'MachineInspections',rows:[{...common,inspectionKey,inspectedAt:now,shift:setup.shift,team:setup.team,overallResult:'Pass',remark:setup.inspectionRemark.trim()}]},{list:'InspectionResults',rows:results},{list:'ConditionValues',rows:conditionValues},{list:'KnifeSelectionRecords',rows:knifeRows},{list:'KnifeReplacements',rows:replacementRows}];
}
$('#record-form').addEventListener('submit',async event=>{event.preventDefault();if(state.saving||state.saved)return;const controls=$$('#record-form input,#record-form select,#record-form textarea,#record-form button,#operator-exit');try{state.saving=true;controls.forEach(e=>e.disabled=true);$('#save-status').textContent='กำลังยืนยัน M365 และส่งข้อมูล กรุณารอสักครู่';const result=await SlitterSharePoint.save(collectGroups());state.saved=true;$('#success-message').textContent=`Document No. ${result.docNo} · บันทึกครบ ${result.count} รายการ`;enableThrough('success');$('#new-record').disabled=false;goStep('success');}catch(error){$('#save-status').textContent=error.message;notify(error.message);}finally{state.saving=false;if(!state.saved)controls.forEach(e=>e.disabled=false);}});
$('#new-record').addEventListener('click',()=>window.location.reload());

$('#microsoft-login').addEventListener('click',async()=>{const button=$('#microsoft-login');button.disabled=true;try{const response=await SlitterSharePoint.token();button.textContent=response.account.name||response.account.username;showView('management-view');$('#operator-exit').hidden=false;}catch(error){notify(error.message);}finally{button.disabled=false;}});
$('#admin-trigger').addEventListener('click',()=>{$('#admin-panel').hidden=!$('#admin-panel').hidden;});
$('#review-record').disabled=true;
$('#apply-filter').addEventListener('click',()=>{$('#confirm-summary');$('.empty-data').innerHTML='<b>พร้อมเชื่อมข้อมูลการผลิต</b><span>ส่วนอ่านข้อมูล SharePoint จะทำในขั้นตอนถัดไป</span>';$('#save-record');$('#management-view').querySelector('.secondary').disabled=false;});
$('#management-view .filter-card .ghost').addEventListener('click',()=>{$$('#management-view input,#management-view select').forEach(e=>e.value='');$('.empty-data').innerHTML='<b>ยังไม่มีข้อมูล</b><span>กรอกตัวกรองแล้วกด “กรองข้อมูล”</span>';});
refreshDocumentNo();renderConditions();
