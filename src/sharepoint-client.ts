import type { AccountInfo } from "@azure/msal-browser";

export const SITE_ID =
  "fitesacnc.sharepoint.com,9c41b493-3a6a-4c01-8d43-26b203783d58,1e97269c-5340-47e3-b05b-6170a179d56b";

export const LISTS = {
  users: "e09422fa-1425-4c87-93b6-a06388d122dd",
  production: "cb93c4d0-9fe2-4a54-b867-409a96bf1e0a",
  inspections: "ba3a8dbd-62e6-48d4-8a6e-a6473ce31145",
  inspectionResults: "33f6a027-9be5-41d8-b231-05393d8850b2",
  conditionValues: "17479d1f-056a-4783-a311-ac62613d9ba6",
  knifeSelections: "d0c5e332-16b1-4fed-afd3-8c20d16cd38a",
  knifeReplacements: "925b3412-b79f-4815-9473-27790ac18ac7",
  parameters: "7cf3f2a4-2214-4e75-95ea-ea9b2a504ef6",
} as const;

type ListRow = { id: string; fields: Record<string, unknown> };

export async function graphRequest(
  path: string,
  token: string,
  init: RequestInit = {},
) {
  if (!token) throw new Error("กรุณาเข้าสู่ระบบ Microsoft 365");
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers || {}),
    },
  });
  const data = response.status === 204 ? null : await response.json();
  if (!response.ok) {
    throw new Error(data?.error?.message || "เชื่อมต่อ SharePoint ไม่สำเร็จ");
  }
  return data;
}

export async function listItems(listId: string, token: string) {
  const rows: ListRow[] = [];
  let path = `/sites/${SITE_ID}/lists/${listId}/items?$expand=fields&$top=999`;
  while (path) {
    const data = await graphRequest(path, token);
    rows.push(...(data.value || []));
    const next = data["@odata.nextLink"] as string | undefined;
    path = next ? next.replace("https://graph.microsoft.com/v1.0", "") : "";
  }
  return rows;
}

export async function addListItems(
  items: Array<{ listId: string; fields: Record<string, unknown> }>,
  token: string,
) {
  for (let offset = 0; offset < items.length; offset += 20) {
    const chunk = items.slice(offset, offset + 20);
    const data = await graphRequest("/$batch", token, {
      method: "POST",
      body: JSON.stringify({
        requests: chunk.map((item, index) => ({
          id: String(index + 1),
          method: "POST",
          url: `/sites/${SITE_ID}/lists/${item.listId}/items`,
          headers: { "Content-Type": "application/json" },
          body: { fields: item.fields },
        })),
      }),
    });
    const failure = (data.responses || []).find(
      (response: { status: number }) => response.status < 200 || response.status >= 300,
    );
    if (failure) {
      throw new Error(failure.body?.error?.message || `SharePoint error ${failure.status}`);
    }
  }
}

async function sha256(value: string) {
  const bytes = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

export async function getUserProfile(
  account: AccountInfo,
  token: string,
  username: string,
  password: string,
) {
  const rows = await listItems(LISTS.users, token);
  const accountEmail = account.username.trim().toLowerCase();
  const enteredUser = username.trim().toLowerCase();
  const row = rows.find(({ fields }) => {
    const employeeId = String(fields.Title || "").trim().toLowerCase();
    const email = String(fields.Email || "").trim().toLowerCase();
    return (employeeId === enteredUser || email === enteredUser) && email === accountEmail;
  });
  if (!row || row.fields.Active === false) {
    throw new Error("Username ไม่ตรงกับบัญชี M365 หรือบัญชีถูกระงับ");
  }
  const stored = String(row.fields.PINHash || "").trim().toLowerCase();
  const expected = stored.startsWith("sha256:") ? stored.slice(7) : stored;
  if (!username.trim() || !password || !/^[a-f0-9]{64}$/.test(expected)) {
    throw new Error("Username หรือ Password ยังไม่ได้ตั้งค่าใน Users List");
  }
  if ((await sha256(password)) !== expected) {
    throw new Error("Username หรือ Password ไม่ถูกต้อง");
  }
  const roleValue = String(row.fields.Role || "operator").toLowerCase();
  return {
    username: String(row.fields.Title || account.username),
    name: String(row.fields.FullName || account.name || account.username),
    role: roleValue === "admin" ? "admin" : roleValue === "viewer" ? "viewer" : "operator",
    team: String(row.fields.Team || ""),
    area: String(row.fields.Area || ""),
  } as const;
}

export async function loadParameters(area: string, token: string) {
  const rows = await listItems(LISTS.parameters, token);
  return rows
    .filter(({ fields }) => fields.field_10 !== false && String(fields.field_2 || "") === area)
    .map(({ id, fields }) => ({
      id,
      code: String(fields.field_1 || ""),
      name: String(fields.Title || ""),
      dataType: String(fields.field_3 || "Text"),
      unit: String(fields.field_4 || ""),
      required: fields.field_6 === true,
      minValue: fields.field_7 == null ? null : Number(fields.field_7),
      maxValue: fields.field_8 == null ? null : Number(fields.field_8),
      displayOrder: fields.field_9 == null ? 999 : Number(fields.field_9),
    }))
    .filter((item) => item.code && item.name)
    .sort((left, right) => left.displayOrder - right.displayOrder || left.name.localeCompare(right.name));
}

type SaveCondition = Record<string, string> & {
  jumboSlot: string;
  inspectionNo: string;
  replacements: string;
  parameterValues?: string;
};

export async function saveProductionRecord(
  body: {
    documentNo: string;
    area: string;
    machine: string;
    crew: string;
    period: string;
    operator: string;
    checks: Record<string, "pass" | "fail">;
    conditions: SaveCondition[];
  },
  token: string,
) {
  const now = new Date().toISOString();
  const revision = 1;
  const inspectionKey = `${body.documentNo}-R${revision}`;
  const operations: Array<{ listId: string; fields: Record<string, unknown> }> = [];
  operations.push({
    listId: LISTS.inspections,
    fields: {
      Title: inspectionKey,
      Revision: revision,
      InspectionDateTime: now,
      DocumentNo: body.documentNo,
      Shift: body.period,
      Team: body.crew,
      OperatorID: body.operator,
      Area: body.area,
      OverallResult: Object.values(body.checks).includes("fail") ? "Fail" : "Pass",
      Remark: `Machine ${body.machine}`,
    },
  });
  Object.entries(body.checks).forEach(([name, result], index) => {
    operations.push({
      listId: LISTS.inspectionResults,
      fields: {
        Title: `${inspectionKey}-CHK${String(index + 1).padStart(2, "0")}`,
        ColumnType_x002f__x0e15__x0e31__: inspectionKey,
        DocumentNo: body.documentNo,
        Revision: revision,
        Area: body.area,
        Category: index < 10 ? "Machine" : "Dolly",
        ItemCode: `CHK${String(index + 1).padStart(2, "0")}`,
        ChecklistName: name,
        InputType: "PassFail",
        ResultText: result,
        ResultStatus: result === "pass" ? "Pass" : "Fail",
        OperatorID: body.operator,
        CheckedDateTime: now,
      },
    });
  });
  const numberValue = (value: string) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  };
  for (const condition of body.conditions) {
    const conditionNo = (Number(condition.jumboSlot) - 1) * 2 + Number(condition.inspectionNo);
    const recordKey = `${body.documentNo}-C${conditionNo}-R${revision}`;
    operations.push({
      listId: LISTS.production,
      fields: {
        Title: body.documentNo,
        Revision: revision,
        IsLatest: true,
        RecordDate: now,
        Shift: body.period,
        Team: body.crew,
        OperatorID: body.operator,
        ProductCode: condition.productCode,
        JumboNo: condition.jumboNo,
        ConditionRound: conditionNo,
        SlitNo: condition.slitNo,
        WidthOnJumbo: numberValue(condition.width),
        Speed: numberValue(condition.speed),
        RampUpDown: condition.ramp,
        WinderTension: condition.tension,
        RiderRollPressure: condition.pressure,
        Torque: condition.torque,
        RollLength: numberValue(condition.rollLength),
        RollDiameter: numberValue(condition.rollDiameter),
        Remark: `${body.area}/${body.machine} · หน้าตัดม้วน: ${condition.crossSection}`,
      },
    });
    const standard = [
      ["WIDTH", "Width on Jumbo", "Number", "mm", condition.width],
      ["SPEED", "Speed", "Number", "m/min", condition.speed],
      ["RAMP", "Ramp Up / Down", "Text", "", condition.ramp],
      ["TENSION", "Winder Tension", "Text", "", condition.tension],
      ["PRESSURE", "Rider Roll Pressure", "Text", "", condition.pressure],
      ["TORQUE", "Torque", "Text", "", condition.torque],
      ["ROLL_LENGTH", "Roll Length", "Number", "m", condition.rollLength],
      ["ROLL_DIAMETER", "Roll Diameter", "Number", "mm", condition.rollDiameter],
      ["CROSS_SECTION", "ตรวจสอบหน้าตัดม้วน", "Choice", "", condition.crossSection],
    ];
    let dynamic: Array<{ code: string; name: string; dataType: string; unit: string; value: string }> = [];
    try {
      dynamic = JSON.parse(condition.parameterValues || "[]");
    } catch {
      dynamic = [];
    }
    for (const [code, name, dataType, unit, value] of [
      ...standard,
      ...dynamic.map((item) => [item.code, item.name, item.dataType, item.unit, item.value]),
    ]) {
      operations.push({
        listId: LISTS.conditionValues,
        fields: {
          Title: `${recordKey}-${code}`,
          RecordKey: recordKey,
          DocumentNo: body.documentNo,
          Revision: revision,
          Area: body.area,
          ProductCode: condition.productCode,
          JumboNo: condition.jumboNo,
          ConditionRound: condition.inspectionNo,
          ParameterCode: code,
          ParameterName: name,
          DataType: dataType,
          ValueNumber: dataType === "Number" ? numberValue(value) : null,
          ValueText: dataType === "Number" ? "" : value,
          Unit: unit,
          OperatorID: body.operator,
          RecordedDateTime: now,
        },
      });
    }
    for (const knifeNo of (condition.knifeNumbers || "").split(",").filter(Boolean)) {
      operations.push({
        listId: LISTS.knifeSelections,
        fields: {
          Title: `${recordKey}-K${knifeNo}`,
          RecordKey: recordKey,
          DocumentNo: body.documentNo,
          Revision: revision,
          Area: body.area,
          ProductCode: condition.productCode,
          JumboNo: condition.jumboNo,
          WidthPattern: condition.width,
          KnifeNo: Number(knifeNo),
          OperatorID: body.operator,
          RecordedDateTime: now,
        },
      });
    }
    let replacements: Array<{ knifeNo?: string; life?: string; reason?: string }> = [];
    try {
      replacements = JSON.parse(condition.replacements || "[]");
    } catch {
      replacements = [];
    }
    replacements.forEach((replacement, index) => {
      operations.push({
        listId: LISTS.knifeReplacements,
        fields: {
          Title: `${recordKey}-REP${index + 1}`,
          RecordKey: recordKey,
          DocumentNo: body.documentNo,
          Revision: revision,
          Area: body.area,
          ProductCode: condition.productCode,
          JumboNo: condition.jumboNo,
          KnifeNo: Number(replacement.knifeNo),
          UsageLife: numberValue(replacement.life || ""),
          UsageLifeUnit: "Hours",
          Reason: replacement.reason || "",
          ReplacementDateTime: now,
          OperatorID: body.operator,
        },
      });
    });
  }
  await addListItems(operations, token);
  return { id: body.documentNo, documentNo: body.documentNo };
}

export async function loadAdminRecords(
  filters: { date: string; productCode: string; area: string; machine: string },
  token: string,
) {
  const [productionRows, knifeRows, replacementRows] = await Promise.all([
    listItems(LISTS.production, token),
    listItems(LISTS.knifeSelections, token),
    listItems(LISTS.knifeReplacements, token),
  ]);
  const knivesByKey = new Map<string, string[]>();
  knifeRows.forEach((row) => {
    const key = String(row.fields.RecordKey || "");
    knivesByKey.set(key, [...(knivesByKey.get(key) || []), String(row.fields.KnifeNo || "")].filter(Boolean));
  });
  const replacementsByKey = new Map<string, Array<{ id: string; knifeNo: string; life: string; reason: string }>>();
  replacementRows.forEach((row) => {
    const key = String(row.fields.RecordKey || "");
    replacementsByKey.set(key, [
      ...(replacementsByKey.get(key) || []),
      { id: row.id, knifeNo: String(row.fields.KnifeNo || ""), life: String(row.fields.UsageLife || ""), reason: String(row.fields.Reason || "") },
    ]);
  });
  const date = filters.date.replaceAll("-", "");
  const records = productionRows.flatMap((row) => {
    const f = row.fields;
    if (f.IsLatest === false) return [];
    const documentNo = String(f.Title || "");
    const conditionNo = Number(f.ConditionRound || 1);
    const recordKey = `${documentNo}-C${conditionNo}-R${Number(f.Revision || 1)}`;
    const machine = documentNo.split("-").pop() || "";
    const area = machine.startsWith("3") ? "TH3" : machine.startsWith("4") ? "TH4" : "";
    const rowDate = String(f.RecordDate || "").slice(0, 10).replaceAll("-", "");
    if (date && rowDate !== date && !documentNo.startsWith(date)) return [];
    if (filters.productCode && String(f.ProductCode || "").toUpperCase() !== filters.productCode.toUpperCase()) return [];
    if (filters.area && area !== filters.area) return [];
    if (filters.machine && machine !== filters.machine) return [];
    const remark = String(f.Remark || "");
    return [{
      id: Number(row.id), documentNo, date: String(f.RecordDate || documentNo.slice(0, 8)), area, machine,
      operator: String(f.OperatorID || ""), updatedAt: String(f.Modified || f.RecordDate || ""), conditionNo,
      productCode: String(f.ProductCode || ""), jumboNo: String(f.JumboNo || ""), slitNo: String(f.SlitNo || ""),
      width: String(f.WidthOnJumbo || ""), speed: String(f.Speed || ""), ramp: String(f.RampUpDown || ""),
      tension: String(f.WinderTension || ""), pressure: String(f.RiderRollPressure || ""), torque: String(f.Torque || ""),
      rollLength: String(f.RollLength || ""), rollDiameter: String(f.RollDiameter || ""),
      crossSection: remark.includes("abnormal") ? "abnormal" : remark.includes("normal") ? "normal" : "",
      knifeNumbers: knivesByKey.get(recordKey) || [], replacements: replacementsByKey.get(recordKey) || [],
    }];
  });
  records.sort((left, right) => right.date.localeCompare(left.date) || right.id - left.id);
  return records;
}
