# SlitterFlow

ระบบบันทึกและจัดการข้อมูล Slitter สำหรับ Operator, Viewer และ Admin

- Operator ใช้ Employee ID + PIN เพื่อเพิ่มข้อมูลใหม่
- Viewer/Admin ใช้ Microsoft 365 เพื่อดู กรอง Export และสร้าง Revision
- SharePoint เป็นฐานข้อมูลส่วนกลาง

GitHub Pages: `https://fcnc2.github.io/SlitterFlow/`

## M365 save integration (configuration pending)

Operator opens the existing form with Employee ID + PIN, fills it, then presses Save.
PIN remains an explicitly labelled prototype: it is **not verified** and is not sent to Graph or stored.
Employee ID is self-reported; the authenticated M365 account is recorded by SharePoint's Created By.
A real PIN verification service is still required before treating PIN as authentication.

One shared MSAL client obtains a delegated token at save time. Existing active sessions use
silent token acquisition. Microsoft Graph writes the five existing lists; no Power Automate
endpoint or client secret is used. Viewer/Admin navigation is unchanged from the prototype.

### Required before enabling

Provide the SharePoint site URL, list IDs and actual internal column names. Populate
`sharepoint-config.js` and only set `schemaVerified: true` after verifying types and mappings.
The app checks mapped column existence before the first write; it does not create lists or columns.
Use the existing Entra SPA redirect URI `https://fcnc2.github.io/SlitterFlow/` and delegated
`User.Read`, `Sites.ReadWrite.All` consent. Each saving account also needs SharePoint write access.

Map these payload keys under each list's `fields` object (all keys are required in the mapping):

| List | Payload keys |
| --- | --- |
| SlitterRecord | docNo, area, machine, dayNight, team, operator, recordDate, createdTime, remark, status |
| PreStartCheck | docNo, checklistName, passed, dollyCuts, remark |
| SlitCondition | docNo, productCode, jumboNo, slitNo, ramp, width, speed, tension, pressure, torque, rollLength, rollDiameter |
| KnifeUsage | docNo, slitNo, knifeNo |
| KnifeReplacement | docNo, knifeNo, lifeHours, reason |

Numeric condition fields may be null. `passed` is Boolean. `recordDate` is YYYY-MM-DD;
`createdTime` is ISO UTC. `knifeNo` is text (01–50). Other values are text. Configure the actual
column types accordingly, or adapt serialization to the existing schema before enabling.
The header status column must support `Pending` and `Complete`. Production readers must
exclude Pending documents. This prototype contains one condition per submission.

### Failure handling and verification

Save revalidates hidden steps, inspection, cuts below 400 and at least one knife. Replacement
is optional. The form locks during submission. Complete is set only after every item succeeds.
Graph multi-list writes are not transactional: partial records remain Pending for reconciliation.
A sessionStorage journal (`slitterflow-pending-write`) blocks repeat sends after uncertain writes.
Before clearing this journal, an administrator must reconcile ALL items for its document number,
including a request that might have succeeded without returning a response. Closing the tab can
lose this journal; cross-tab/session deduplication and recovery remain a production prerequisite.
No live writes have been tested; site/schema information and real account access are still needed.

Run `node --test tests/sharepoint.test.cjs` for mocked configuration, completion, token reuse,
and partial-failure checks. Perform a controlled live save and confirm all five lists before rollout.

References: [Graph list create](https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0),
[MSAL token acquisition](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/acquire-token).
