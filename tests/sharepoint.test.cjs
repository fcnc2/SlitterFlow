const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
function setup(failAt) {
  const state = new Map(), calls = [];
  const account = { tenantId: '464ae0fc-7c03-447d-9776-fbfad0c89bcf', username: 'test' };
  let active, logins = 0;
  const context = { console, sessionStorage: { getItem: k => state.get(k), setItem: (k,v) => state.set(k,v), removeItem: k => state.delete(k) }, msal: {
    InteractionRequiredAuthError: class extends Error {},
    PublicClientApplication: class { getActiveAccount() { return active; } setActiveAccount(a) { active = a; } async loginPopup() { logins++; return { account, accessToken: 'test' }; } async acquireTokenSilent() { return { account, accessToken: 'test' }; } }
  }, SLITTER_SHAREPOINT: { siteId: 'site', schemaVerified: true, lists: { ProductionRecords: { id:'header', fields:{docNo:'Title',isLatest:'IsLatest'} }, KnifeSelectionRecords: {id:'knife', fields:{docNo:'Title',knifeNo:'KnifeNo'}} } },
  fetch: async (url, options) => {
    calls.push({url,options});
    if (failAt === calls.length) throw new Error('network lost');
    const data = options.method === 'POST' ? {id:String(calls.length)} : {value: ['Title','IsLatest','KnifeNo'].map(name => ({name}))};
    return {ok:true,status:200,json:async()=>data};
  }};
  context.window = context;
  vm.runInNewContext(fs.readFileSync('sharepoint.js','utf8'), context);
  return {context,calls,state,logins:()=>logins};
}
const groups = [{list:'ProductionRecords',rows:[{docNo:'20260910D-3A',isLatest:false},{docNo:'20260910D-3A',isLatest:false}]},{list:'KnifeSelectionRecords',rows:[{docNo:'20260910D-3A',knifeNo:1}]}];
test('missing config blocks before login or network', async()=> {
 const s=setup(); s.context.SLITTER_SHAREPOINT.siteId='';
 await assert.rejects(s.context.SlitterSharePoint.save(groups)); assert.equal(s.calls.length,0); assert.equal(s.logins(),0);
});
test('all lists saved before Complete; token reused',async()=>{
 const s=setup(); const result=await s.context.SlitterSharePoint.save(groups);
 assert.equal(result.count,3); assert.equal(s.calls.filter(call=>call.options.method==='PATCH').length,2);
 assert.equal(JSON.parse(s.calls.at(-1).options.body).IsLatest,true);
 assert.equal(s.state.size,0); await s.context.SlitterSharePoint.token(); assert.equal(s.logins(),1);
});
test('partial or uncertain writes retain journal and block repeated POST',async()=>{
 const s=setup(4); await assert.rejects(s.context.SlitterSharePoint.save(groups));
 assert.equal(s.state.size,1); const count=s.calls.length;
 await assert.rejects(s.context.SlitterSharePoint.save(groups)); assert.equal(s.calls.length,count);
 assert.equal(s.calls.some(c=>c.options.method==='PATCH'),false);
});
test('missing mapping blocks every write',async()=>{
 const s=setup(); delete s.context.SLITTER_SHAREPOINT.lists.KnifeSelectionRecords.fields.knifeNo;
 await assert.rejects(s.context.SlitterSharePoint.save(groups)); assert.equal(s.calls.length,0);
});
