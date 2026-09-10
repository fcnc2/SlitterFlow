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
  const common = { docNo };
  const condition = { ...common };
  for (const key of ['productCode', 'jumboNo', 'slitNo', 'ramp']) condition[key] = data[key].trim();
  condition.productCode = condition.productCode.toUpperCase();
  for (const key of ['width', 'speed', 'tension', 'pressure', 'torque', 'rollLength', 'rollDiameter']) {
    condition[key] = data[key] === '' ? null : Number(data[key]);
  }
  const replacements = [...document.querySelectorAll('.replacement-row')].map(row => {
    const inputs = row.querySelectorAll('input');
    if (!inputs[2].value.trim()) throw new Error('กรุณาระบุสาเหตุเปลี่ยนใบมีด');
    return { ...common, knifeNo: inputs[0].value.padStart(2, '0'), lifeHours: Number(inputs[1].value), reason: inputs[2].value.trim() };
  });
  return [
    { list: 'SlitterRecord', rows: [{ ...common, area: data.area, machine: data.machine, dayNight: data.shift, team: data.team, operator: operatorId, recordDate: data.recordDate, createdTime: new Date().toISOString(), remark: data.jobRemark, status: 'Pending' }] },
    { list: 'PreStartCheck', rows: checks.map(input => ({ ...common, checklistName: input.parentElement.textContent.trim(), passed: input.checked, dollyCuts: Number(document.querySelector('#dolly-cuts').value), remark: data.inspectionRemark })) },
    { list: 'SlitCondition', rows: [condition] },
    { list: 'KnifeUsage', rows: knives.map(knifeNo => ({ ...common, slitNo: data.slitNo, knifeNo })) },
    { list: 'KnifeReplacement', rows: replacements }
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
