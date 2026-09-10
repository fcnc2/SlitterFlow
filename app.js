const dialog = document.querySelector('#operator-dialog');
const toast = document.querySelector('#toast');
const microsoftButton = document.querySelector('#microsoft-login');
let operatorId = '';
let saving = false;
let saved = false;

document.querySelector('[data-open="operator-dialog"]').addEventListener('click', () => dialog.showModal());

function notify(message) {
  toast.textContent = message;
  toast.classList.add('show');
  window.setTimeout(() => toast.classList.remove('show'), 3600);
}

document.querySelector('#operator-login').addEventListener('click', () => {
  const employeeId = dialog.querySelector('[name="employeeId"]').value.trim();
  const pin = dialog.querySelector('[name="pin"]').value.trim();
  if (!employeeId || !pin) {
    notify('กรุณากรอก Employee ID และ PIN');
    return;
  }
  operatorId = employeeId;
  dialog.querySelector('[name="pin"]').value = '';
  document.querySelector('#operator-name').textContent = `ID ${employeeId}`;
  document.querySelector('#home-view').hidden = true;
  document.querySelector('#operator-view').hidden = false;
  dialog.close();
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

const stepTabs = [...document.querySelectorAll('.step-tab')];
const stepPanels = [...document.querySelectorAll('.step-panel')];

function goToStep(step) {
  if (saving || saved) return;
  stepTabs.forEach(tab => tab.classList.toggle('active', tab.dataset.step === step));
  stepPanels.forEach(panel => panel.classList.toggle('active', panel.dataset.panel === step));
  window.scrollTo({ top: 80, behavior: 'smooth' });
}

stepTabs.forEach(tab => tab.addEventListener('click', () => !tab.disabled && goToStep(tab.dataset.step)));

document.querySelector('#complete-inspection').addEventListener('click', () => {
  const machinePassed = [...document.querySelectorAll('#machine-checks input')].every(item => item.checked);
  const dollyPassed = ['#dolly-clean', '#dolly-wear', '#dolly-size'].every(id => document.querySelector(id).checked);
  const rawCuts = document.querySelector('#dolly-cuts').value;
  const cuts = rawCuts === '' ? NaN : Number(rawCuts);
  if (!machinePassed || !dollyPassed || !Number.isInteger(cuts) || cuts < 0 || cuts >= 400) {
    notify('ต้องตรวจครบทุกข้อ และจำนวนครั้งตัด Dolly ต้องน้อยกว่า 400');
    return;
  }
  stepTabs.slice(1).forEach(tab => { tab.disabled = false; tab.classList.remove('locked'); });
  notify('การตรวจสอบผ่านแล้ว');
  goToStep('condition');
});

document.querySelectorAll('.next-step').forEach(button => button.addEventListener('click', () => goToStep(button.dataset.next)));

const knifeGrid = document.querySelector('#knife-grid');
for (let number = 1; number <= 50; number += 1) {
  const id = `knife-${number}`;
  knifeGrid.insertAdjacentHTML('beforeend', `<label for="${id}"><input id="${id}" type="checkbox" value="${String(number).padStart(2, '0')}">${String(number).padStart(2, '0')}</label>`);
}

document.querySelector('#copy-knives').addEventListener('click', () => notify('ฟังก์ชันคัดลอกจะใช้งานเมื่อเชื่อม SharePoint'));

document.querySelector('#add-replacement').addEventListener('click', () => {
  const list = document.querySelector('#replacement-list');
  list.querySelector('.empty-row')?.remove();
  const row = document.createElement('div');
  row.className = 'replacement-row';
  row.innerHTML = '<label>เบอร์ใบมีด<input type="number" min="1" max="50" required></label><label>อายุการใช้งาน<input type="number" min="0" required></label><label>สาเหตุ<input required></label><button type="button">ลบ</button>';
  row.querySelector('button').addEventListener('click', () => row.remove());
  list.append(row);
});

function collectRecord() {
  const form = document.querySelector('#record-form');
  for (const input of form.querySelectorAll('input, select, textarea')) {
    if (!input.checkValidity()) {
      goToStep(input.closest('.step-panel').dataset.panel);
      input.reportValidity();
      throw new Error('กรุณากรอกข้อมูลให้ครบและถูกต้อง');
    }
  }
  const checks = [...document.querySelectorAll('#machine-checks input, .dolly-grid input[type="checkbox"]')];
  if (!checks.every(input => input.checked)) {
    goToStep('inspection');
    throw new Error('ต้องผ่านการตรวจสอบครบทุกข้อก่อนบันทึก');
  }
  const knives = [...knifeGrid.querySelectorAll('input:checked')].map(input => input.value);
  if (!knives.length) { goToStep('knife'); throw new Error('กรุณาเลือกใบมีดที่ใช้งานอย่างน้อย 1 ใบ'); }
  const data = Object.fromEntries(new FormData(form));
  if (!operatorId) throw new Error('กรุณาเข้าสู่หน้า Operator');
  if (!({ TH3: ['3A', '3B'], TH4: ['4A', '4B'] })[data.area]?.includes(data.machine)) throw new Error('Area และ Machine ไม่ตรงกัน');
  const docNo = `${data.recordDate.replaceAll('-', '')}${data.shift === 'Day' ? 'D' : 'N'}-${data.machine}`;
  const now = new Date().toISOString();
  const revision = 1;
  const conditionRound = 1;
  const common = { docNo, revision, area: data.area, operatorId };
  const numberValue = key => data[key] === '' ? null : Number(data[key]);
  const productCode = data.productCode.trim().toUpperCase();
  const jumboNo = data.jumboNo.trim();
  const slitNo = data.slitNo.trim();
  const inspectionKey = `${docNo}-R${revision}`;
  const checkResults = checks.map((input, index) => {
    const category = input.closest('.dolly-grid') ? 'Dolly' : 'Machine';
    const itemNo = category === 'Dolly' ? index - document.querySelectorAll('#machine-checks input').length + 1 : index + 1;
    return {
      ...common,
      resultKey: `${inspectionKey}-${category === 'Dolly' ? 'D' : 'M'}${String(itemNo).padStart(2, '0')}`,
      inspectionKey,
      category,
      itemCode: `${category === 'Dolly' ? 'D' : 'M'}${String(itemNo).padStart(2, '0')}`,
      checklistName: input.parentElement.textContent.trim(),
      inputType: 'PassFail',
      resultText: 'Pass',
      resultStatus: 'Pass',
      checkedAt: now,
      remark: data.inspectionRemark.trim()
    };
  });
  checkResults.push({
    ...common,
    resultKey: `${inspectionKey}-D04`, inspectionKey, category: 'Dolly', itemCode: 'D04',
    checklistName: 'จำนวนครั้งตัด Paper Core Cutter', inputType: 'Number',
    resultNumber: Number(document.querySelector('#dolly-cuts').value), resultStatus: 'Pass',
    checkedAt: now, remark: data.inspectionRemark.trim()
  });
  const parameters = [
    ['WIDTH', 'Width on Jumbo', 'Number', numberValue('width'), null, 'mm'],
    ['SPEED', 'Speed', 'Number', numberValue('speed'), null, 'm/min'],
    ['RAMP', 'Ramp Up / Down', 'Text', null, data.ramp.trim(), ''],
    ['TENSION', 'Winder Tension', 'Number', numberValue('tension'), null, ''],
    ['PRESSURE', 'Rider Roll Pressure', 'Number', numberValue('pressure'), null, ''],
    ['TORQUE', 'Torque', 'Number', numberValue('torque'), null, ''],
    ['ROLL_LENGTH', 'Roll Length', 'Number', numberValue('rollLength'), null, 'm'],
    ['ROLL_DIAMETER', 'Roll Diameter', 'Number', numberValue('rollDiameter'), null, 'mm'],
    ['CROSS_SECTION', 'ตรวจสอบหน้าตัดม้วน', 'Choice', null, data.crossSection, '']
  ];
  const conditionRows = parameters.filter(([, , , valueNumber, valueText]) => valueNumber !== null || valueText).map(([code, name, dataType, valueNumber, valueText, unit], index) => {
    const recordKey = `${docNo}-C${conditionRound}-${String(index + 1).padStart(2, '0')}`;
    return { ...common, title: recordKey, recordKey, productCode, jumboNo, conditionRound: String(conditionRound), parameterCode: code, parameterName: name, dataType, valueNumber, valueText, unit, recordedAt: now };
  });
  const replacements = [...document.querySelectorAll('.replacement-row')].map(row => {
    const inputs = row.querySelectorAll('input');
    if (!inputs[2].value.trim()) throw new Error('กรุณาระบุสาเหตุเปลี่ยนใบมีด');
    const knifeNo = Number(inputs[0].value);
    const recordKey = `${docNo}-K${String(knifeNo).padStart(2, '0')}-${Date.now()}`;
    return { ...common, title: recordKey, recordKey, productCode, jumboNo, knifeNo, lifeHours: Number(inputs[1].value), lifeUnit: 'Hours', reason: inputs[2].value.trim(), replacedAt: now };
  });
  return [
    { list: 'ProductionRecords', rows: [{ docNo, revision, isLatest: false, recordDate: data.recordDate, shift: data.shift, team: data.team, operatorId, productCode, jumboNo, conditionRound, slitNo, width: numberValue('width'), speed: numberValue('speed'), ramp: data.ramp.trim(), tension: data.tension, pressure: data.pressure, torque: data.torque, rollLength: numberValue('rollLength'), rollDiameter: numberValue('rollDiameter'), remark: data.jobRemark.trim() }] },
    { list: 'MachineInspections', rows: [{ ...common, inspectionKey, inspectedAt: now, shift: data.shift, team: data.team, overallResult: 'Pass', remark: data.inspectionRemark.trim() }] },
    { list: 'InspectionResults', rows: checkResults },
    { list: 'ConditionValues', rows: conditionRows },
    { list: 'KnifeSelectionRecords', rows: knives.map(value => { const knifeNo = Number(value); const recordKey = `${docNo}-K${String(knifeNo).padStart(2, '0')}`; return { ...common, title: recordKey, recordKey, productCode, jumboNo, widthPattern: data.width, knifeNo, recordedAt: now }; }) },
    { list: 'KnifeReplacements', rows: replacements }
  ];
}

document.querySelector('#record-form').addEventListener('submit', async event => {
  event.preventDefault();
  if (saving || saved) return;
  const status = document.querySelector('#save-status');
  const button = document.querySelector('#save-record');
  let controls = [];
  try {
    const groups = collectRecord();
    saving = true;
    controls = [...document.querySelectorAll('#record-form input, #record-form select, #record-form textarea, #record-form button, #operator-exit')].map(element => [element, element.disabled]);
    controls.forEach(([element]) => { element.disabled = true; });
    status.textContent = 'กำลังยืนยันบัญชี M365 และส่งข้อมูล กรุณารอสักครู่';
    const result = await SlitterSharePoint.save(groups);
    saved = true;
    status.textContent = `บันทึกสำเร็จ: ${result.docNo} ส่งครบ ${result.count} รายการ`;
    button.textContent = 'บันทึกสำเร็จแล้ว';
    document.querySelector('#operator-exit').disabled = false;
  } catch (error) {
    status.textContent = error.message;
    notify(error.message);
  } finally {
    saving = false;
    if (!saved) controls.forEach(([element, disabled]) => { element.disabled = disabled; });
  }
});

document.querySelector('#operator-exit').addEventListener('click', () => {
  if (saving) return;
  if (saved) { window.location.reload(); return; }
  document.querySelector('#operator-view').hidden = true;
  document.querySelector('#home-view').hidden = false;
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

microsoftButton.addEventListener('click', async () => {
  microsoftButton.disabled = true;
  try {
    const response = await SlitterSharePoint.token();
    microsoftButton.textContent = response.account.name || response.account.username;
    notify(`เข้าสู่ระบบแล้ว: ${response.account.username}`);
  } catch (error) {
    notify(error.message);
  } finally { microsoftButton.disabled = false; }
});

const areaSelect = document.querySelector('[name="area"]');
areaSelect.addEventListener('change', () => {
  document.querySelector('[name="machine"]').innerHTML = (areaSelect.value === 'TH3' ? ['3A', '3B'] : ['4A', '4B']).map(value => `<option>${value}</option>`).join('');
});
document.querySelector('[name="team"]').addEventListener('change', event => { document.querySelector('#operator-team').textContent = event.target.value; });
