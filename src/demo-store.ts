// Trial data stays in this browser. GitHub Pages has no server or shared database.
const RECORDS_KEY = "slitterflow-demo-records-v1";
const PARAMETERS_KEY = "slitterflow-demo-parameters-v1";
const SUBMISSIONS_KEY = "slitterflow-demo-submissions-v1";
export type DemoSubmission = {
  documentNo: string; area: string; machine: string; operator: string;
  managerName: string; period: string; submittedAt: string;
  reviewedAt: string; reviewedBy: string; reviewNote: string;
  status: "pending" | "approved" | "returned";
  checks: Record<string, "pass" | "fail">;
  conditions: Array<Record<string, string>>;
};
export function demoSubmissions() { return read<DemoSubmission>(SUBMISSIONS_KEY); }
export function reviewDemoSubmission(documentNo: string, reviewer: string, decision: "approved" | "returned", note: string) {
  const items = demoSubmissions();
  const target = items.find(item => item.documentNo === documentNo && item.status === "pending");
  if (!target) throw new Error("เอกสารนี้ไม่ได้อยู่ในสถานะรอตรวจสอบแล้ว");
  if (decision === "returned" && !note.trim()) throw new Error("กรุณาระบุเหตุผลที่ส่งกลับ");
  const next = items.map(item => item === target ? { ...item, status: decision, reviewedBy: reviewer, reviewedAt: new Date().toISOString(), reviewNote: note.trim() } : item);
  localStorage.setItem(SUBMISSIONS_KEY, JSON.stringify(next));
}

export type DemoRecord = {
  id: number; documentNo: string; date: string; area: string; machine: string;
  operator: string; updatedAt: string; conditionNo: number; productCode: string;
  jumboNo: string; slitNo: string; width: string; speed: string; ramp: string;
  tension: string; pressure: string; torque: string; rollLength: string;
  rollDiameter: string; crossSection: string; knifeNumbers: string[];
  replacements: { id: string; knifeNo: string; life: string; reason: string }[];
};

export type DemoParameter = {
  id: string; area: "TH3" | "TH4"; category: string; name: string;
  inputType: string; minValue: string; maxValue: string; unit: string;
  required: boolean; active: boolean;
};

function read<T>(key: string): T[] {
  try { const value = JSON.parse(localStorage.getItem(key) || "[]"); return Array.isArray(value) ? value : []; }
  catch { return []; }
}

export function demoRecords() { return read<DemoRecord>(RECORDS_KEY); }
export function demoParameters() { return read<DemoParameter>(PARAMETERS_KEY); }
export function writeDemoParameters(items: DemoParameter[]) {
  localStorage.setItem(PARAMETERS_KEY, JSON.stringify(items));
}

export function saveDemoRecord(body: {
  documentNo: string; area: string; machine: string; operator: string;
  managerName: string; period: string; checks: Record<string, "pass" | "fail">;
  conditions: Array<Record<string, string>>;
}) {
  const now = new Date().toISOString();
  const existing = demoRecords();
  const submissions = demoSubmissions();
  const previous = submissions.find(item => item.documentNo === body.documentNo);
  if (previous && previous.status !== "returned") throw new Error("เอกสารนี้ส่งแล้วและยังแก้ไขไม่ได้");
  const records: DemoRecord[] = body.conditions.map((condition, index) => {
    let replacements: DemoRecord["replacements"] = [];
    try { replacements = JSON.parse(condition.replacements || "[]"); } catch { /* optional */ }
    return {
      id: Date.now() + index, documentNo: body.documentNo, date: now.slice(0, 10),
      area: body.area, machine: body.machine, operator: body.operator, updatedAt: now,
      conditionNo: (Number(condition.jumboSlot) - 1) * 2 + Number(condition.inspectionNo),
      productCode: condition.productCode || "", jumboNo: condition.jumboNo || "",
      slitNo: condition.slitNo || "", width: condition.width || "", speed: condition.speed || "",
      ramp: condition.ramp || "", tension: condition.tension || "", pressure: condition.pressure || "",
      torque: condition.torque || "", rollLength: condition.rollLength || "",
      rollDiameter: condition.rollDiameter || "", crossSection: condition.crossSection || "",
      knifeNumbers: (condition.knifeNumbers || "").split(",").filter(Boolean), replacements,
    };
  });
  localStorage.setItem(RECORDS_KEY, JSON.stringify([...existing.filter(item => item.documentNo !== body.documentNo), ...records]));
  localStorage.setItem(SUBMISSIONS_KEY, JSON.stringify([
    ...submissions.filter(item => item.documentNo !== body.documentNo),
    { documentNo: body.documentNo, area: body.area, machine: body.machine, operator: body.operator,
      managerName: body.managerName, period: body.period, submittedAt: now,
      reviewedAt: "", reviewedBy: "", reviewNote: "", status: "pending",
      checks: body.checks, conditions: body.conditions },
  ]));
  return records[0]?.id ?? Date.now();
}

export function filterDemoRecords(filters: { date: string; productCode: string; area: string; machine: string }) {
  return demoRecords().filter(record =>
    (!filters.date || record.date === filters.date) &&
    (!filters.productCode || record.productCode.toUpperCase().includes(filters.productCode.toUpperCase())) &&
    (!filters.area || record.area === filters.area) &&
    (!filters.machine || record.machine === filters.machine));
}
