const dialog = document.querySelector('#operator-dialog');
const toast = document.querySelector('#toast');
const microsoftButton = document.querySelector('#microsoft-login');

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

const msalReady = msalClient.initialize();

document.querySelector('[data-open="operator-dialog"]').addEventListener('click', () => dialog.showModal());

function notify(message) {
  toast.textContent = message;
  toast.classList.add('show');
  window.setTimeout(() => toast.classList.remove('show'), 3600);
}

document.querySelector('#operator-login').addEventListener('click', () => {
  notify('รอเชื่อมระบบตรวจสอบ PIN และ SharePoint');
});

microsoftButton.addEventListener('click', async () => {
  microsoftButton.disabled = true;
  microsoftButton.textContent = 'กำลังเข้าสู่ระบบ...';

  try {
    await msalReady;
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
