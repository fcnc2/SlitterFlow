# SlitterFlow

ระบบบันทึกและจัดการข้อมูล Slitter สำหรับ Operator, Viewer และ Admin

- Operator ใช้ Employee ID + PIN เพื่อเพิ่มข้อมูลใหม่
- Viewer/Admin ใช้ Microsoft 365 เพื่อดู กรอง Export และสร้าง Revision
- SharePoint เป็นฐานข้อมูลส่วนกลาง

GitHub Pages: `https://fcnc2.github.io/SlitterFlow/`

## M365 save integration

Operator opens the existing form with Employee ID + PIN, fills it, then presses Save.
The published interface uses the restored blue-white five-step Operator flow. PIN remains an explicitly labelled prototype: it is **not verified** and is not sent to Graph or stored.
Employee ID is self-reported; the authenticated M365 account is recorded by SharePoint's Created By.
A real PIN verification service is still required before treating PIN as authentication.

One shared MSAL client obtains a delegated token at save time. Existing active sessions use
silent token acquisition. Microsoft Graph writes the six existing production lists; no Power Automate
endpoint or client secret is used. Viewer/Admin navigation is unchanged from the prototype.

### Required before enabling

The SharePoint site, list IDs, internal column names, types and Choice values were read from the
live SlitterFlow site on 10 September 2026 and are populated in `sharepoint-config.js`.
The app checks mapped column existence before the first write; it does not create lists or columns.
Use the existing Entra SPA redirect URI `https://fcnc2.github.io/SlitterFlow/` and delegated
`User.Read`, `Sites.ReadWrite.All` consent. Each saving account also needs SharePoint write access.

Map these payload keys under each list's `fields` object (all keys are required in the mapping):

| List | Payload keys |
| --- | --- |
| ProductionRecords | Document, revision, latest flag, shift/team/operator and condition summary |
| MachineInspections | Inspection header, time, shift/team/operator, area and result |
| InspectionResults | One row for each machine/Dolly check and cutter count |
| ConditionValues | One row for each condition parameter, including cross-section result |
| KnifeSelectionRecords | One row for each selected knife |
| KnifeReplacements | Optional replacement rows with usage life and reason |

Blank optional values are omitted. Dates use ISO format and KnifeNo is sent as a number.
`ProductionRecords.IsLatest` starts false and changes to true only after all related writes succeed.
Production readers must use rows where IsLatest is true. Each document supports up to six conditions per submission.

### Failure handling and verification

Save revalidates hidden steps, inspection, cuts below 400 and at least one knife. Replacement
is optional. The form locks during submission. IsLatest becomes true only after every item succeeds.
Graph multi-list writes are not transactional: partial records remain linked to a ProductionRecords row where IsLatest is false for reconciliation.
A sessionStorage journal (`slitterflow-pending-write`) blocks repeat sends after uncertain writes.
Before clearing this journal, an administrator must reconcile ALL items for its document number,
including a request that might have succeeded without returning a response. Closing the tab can
lose this journal; cross-tab/session deduplication and recovery remain a production prerequisite.
No live writes have been tested; a controlled write with a real account is still needed.

Run `node --test tests/sharepoint.test.cjs` for mocked configuration, completion, token reuse,
and partial-failure checks. Perform a controlled live save and confirm all six lists before rollout.

References: [Graph list create](https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0),
[MSAL token acquisition](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/acquire-token).
