const dialog = document.querySelector('#operator-dialog');
const toast = document.querySelector('#toast');

document.querySelector('[data-open="operator-dialog"]').addEventListener('click', () => dialog.showModal());

function notify(message) {
  toast.textContent = message;
  toast.classList.add('show');
  window.setTimeout(() => toast.classList.remove('show'), 3600);
}

document.querySelector('#operator-login').addEventListener('click', () => {
  notify('รอเชื่อมระบบตรวจสอบ PIN และ SharePoint');
});

document.querySelector('#microsoft-login').addEventListener('click', () => {
  notify('รอใส่ Client ID และ Tenant ID จาก Microsoft Entra');
});
