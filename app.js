const dialog = document.querySelector('#operator-dialog');
const toast = document.querySelector('#toast');
const microsoftButton = document.querySelector('#microsoft-login');

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
  document.querySelector('#operator-name').textContent = `ID ${employeeId}`;
  document.querySelector('#home-view').hidden = true;
  document.querySelector('#operator-view').hidden = false;
  dialog.close();
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

const stepTabs = [...document.querySelectorAll('.step-tab')];
const stepPanels = [...document.querySelectorAll('.step-panel')];

function goToStep(step) {
  stepTabs.forEach(tab => tab.classList.toggle('active', tab.dataset.step === step));
  stepPanels.forEach(panel => panel.classList.toggle('active', panel.dataset.panel === step));
  window.scrollTo({ top: 80, behavior: 'smooth' });
}

stepTabs.forEach(tab => tab.addEventListener('click', () => !tab.disabled && goToStep(tab.dataset.step)));

document.querySelector('#complete-inspection').addEventListener('click', () => {
  const machinePassed = [...document.querySelectorAll('#machine-checks input')].every(item => item.checked);
  const dollyPassed = ['#dolly-clean', '#dolly-wear', '#dolly-size'].every(id => document.querySelector(id).checked);
  const cuts = Number(document.querySelector('#dolly-cuts').value);
  if (!machinePassed || !dollyPassed || !Number.isFinite(cuts) || cuts < 0 || cuts >= 400) {
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

document.querySelector('#record-form').addEventListener('submit', event => {
  event.preventDefault();
  notify('ข้อมูลผ่านการตรวจสอบแล้ว รอเชื่อมระบบกลางเพื่อส่งเข้า SharePoint');
});

document.querySelector('#operator-exit').addEventListener('click', () => {
  document.querySelector('#operator-view').hidden = true;
  document.querySelector('#home-view').hidden = false;
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

microsoftButton.addEventListener('click', async () => {
  if (typeof msal === 'undefined') {
    notify('โหลดระบบ Microsoft Login ไม่สำเร็จ กรุณาตรวจสอบเครือข่ายบริษัท');
    return;
  }

  microsoftButton.disabled = true;
  microsoftButton.textContent = 'กำลังเข้าสู่ระบบ...';

  try {
    const msalClient = new msal.PublicClientApplication({
      auth: {
        clientId: '497a8bdf-06dc-4e47-9b25-e3fa8b3248d5',
        authority: 'https://login.microsoftonline.com/464ae0fc-7c03-447d-9776-fbfad0c89bcf',
        redirectUri: 'https://fcnc2.github.io/SlitterFlow/'
      },
      cache: {
        cacheLocation: 'sessionStorage'
      }
    });
    const response = await msalClient.loginPopup({
      scopes: ['User.Read', 'Sites.ReadWrite.All'],
      prompt: 'select_account'
    });
    msalClient.setActiveAccount(response.account);
    microsoftButton.textContent = response.account.name || response.account.username;
    notify(`เข้าสู่ระบบแล้ว: ${response.account.username}`);
  } catch (error) {
    console.error('Microsoft sign-in failed', error);
    microsoftButton.disabled = false;
    microsoftButton.textContent = 'เข้าสู่ระบบ Microsoft 365';
    notify('เข้าสู่ระบบไม่สำเร็จ โปรดตรวจ Redirect URI และสิทธิ์ของแอป');
  }
});
