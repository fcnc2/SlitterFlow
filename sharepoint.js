/* No secrets belong in this public repository. M365 supplies delegated access. */
window.SlitterSharePoint = (() => {
  let client;
  const scopes = ['User.Read', 'Sites.ReadWrite.All'];
  function auth() {
    if (!window.msal) throw new Error('โหลด Microsoft Login ไม่สำเร็จ กรุณาตรวจสอบเครือข่าย');
    if (!client) client = new msal.PublicClientApplication({
      auth: {
        clientId: '497a8bdf-06dc-4e47-9b25-e3fa8b3248d5',
        authority: 'https://login.microsoftonline.com/464ae0fc-7c03-447d-9776-fbfad0c89bcf',
        redirectUri: 'https://fcnc2.github.io/SlitterFlow/'
      },
      cache: { cacheLocation: 'sessionStorage' }
    });
    return client;
  }
  async function token() {
    const app = auth();
    const account = app.getActiveAccount();
    let result;
    if (account) {
      try { result = await app.acquireTokenSilent({ scopes, account }); }
      catch (error) {
        if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
        result = await app.acquireTokenPopup({ scopes, account });
      }
    } else {
      result = await app.loginPopup({ scopes, prompt: 'select_account' });
    }
    if (result.account.tenantId !== '464ae0fc-7c03-447d-9776-fbfad0c89bcf') {
      throw new Error('กรุณาใช้บัญชี M365 ของบริษัท');
    }
    app.setActiveAccount(result.account);
    return result;
  }
  async function save(groups) {
    const config = window.SLITTER_SHAREPOINT;
    if (!config?.siteId || !config.schemaVerified) {
      throw new Error('ยังไม่ได้ตั้งค่า SharePoint Site และยืนยันชื่อคอลัมน์ กรุณาติดต่อผู้ดูแล');
    }
    const pendingKey = 'slitterflow-pending-write';
    if (sessionStorage.getItem(pendingKey)) {
      throw new Error('มีการส่งครั้งก่อนที่ยังไม่ยืนยันผล กรุณาให้ผู้ดูแลตรวจ SharePoint ก่อนส่งซ้ำ');
    }
    // Validate all mappings before requesting credentials or writing anything.
    const jobs = groups.flatMap(group => group.rows.map(row => {
      const target = config.lists[group.list];
      if (!target?.id) throw new Error(`ยังไม่ตั้งค่า List: ${group.list}`);
      const fields = {};
      for (const [key, value] of Object.entries(row)) {
        const column = target.fields[key];
        if (!column) throw new Error(`ยังไม่จับคู่คอลัมน์ ${group.list}.${key}`);
        fields[column] = value;
      }
      return { list: group.list, id: target.id, fields };
    }));
    const response = await token();
    const base = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(config.siteId)}/lists/`;
    async function request(path, options = {}) {
      const result = await fetch(base + path, {
        ...options, headers: { Authorization: `Bearer ${response.accessToken}`, 'Content-Type': 'application/json' }
      });
      if (!result.ok) {
        throw new Error(`SharePoint HTTP ${result.status}: ${result.status === 403 ? 'บัญชีหรือแอปไม่มีสิทธิ์บันทึก' : 'ส่งข้อมูลไม่สำเร็จ'}`);
      }
      return result.status === 204 ? null : result.json();
    }
    // Read actual internal names before the first POST. No schema is created automatically.
    for (const id of new Set(jobs.map(job => job.id))) {
      const names = new Set();
      let page = await request(`${encodeURIComponent(id)}/columns?$select=name`);
      for (;;) {
        page.value.forEach(column => names.add(column.name));
        if (!page['@odata.nextLink']) break;
        const next = page['@odata.nextLink'];
        if (!next.startsWith(base)) throw new Error('SharePoint ส่งลิงก์คอลัมน์ไม่ถูกต้อง');
        page = await request(next.slice(base.length));
      }
      for (const job of jobs.filter(job => job.id === id)) {
        for (const name of Object.keys(job.fields)) {
          if (!names.has(name)) throw new Error(`ไม่พบคอลัมน์ ${job.list}.${name}`);
        }
      }
    }
    const journal = { docNo: groups[0].rows[0].docNo, startedAt: new Date().toISOString(), items: [] };
    sessionStorage.setItem(pendingKey, JSON.stringify(journal));
    try {
      for (const job of jobs) {
        const item = await request(`${encodeURIComponent(job.id)}/items`, {
          method: 'POST', body: JSON.stringify({ fields: job.fields })
        });
        if (!item.id) throw new Error('SharePoint ไม่ส่งหมายเลขรายการกลับมา');
        journal.items.push({ list: job.list, id: item.id });
        sessionStorage.setItem(pendingKey, JSON.stringify(journal));
      }
      const header = jobs[0];
      await request(`${encodeURIComponent(header.id)}/items/${encodeURIComponent(journal.items[0].id)}/fields`, {
        method: 'PATCH', body: JSON.stringify({ [config.lists.SlitterRecord.fields.status]: 'Complete' })
      });
      sessionStorage.removeItem(pendingKey);
      return { docNo: journal.docNo, count: journal.items.length };
    } catch (error) {
      throw new Error(`${error.message} — เอกสาร ${journal.docNo} ยืนยันแล้ว ${journal.items.length} รายการ กรุณาตรวจข้อมูลก่อนส่งซ้ำ`);
    }
  }
  return { token, save };
})();
