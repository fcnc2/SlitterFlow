"use client";
import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";
import { useEffect, useMemo, useState } from "react";
import {
  getUserProfile,
  loadAdminRecords,
  loadParameters as loadSharePointParameters,
  saveProductionRecord,
} from "./sharepoint-client";
const machineChecks = [
    "Air supply — Top Knife",
    "ท่อดูดฝุ่นใบมีด",
    "ชุดจัดใบมีด (หนา/บาง)",
    "สายพาน Unwinder",
    "รถ Unload Table",
    "Dust Collector",
    "ชุดจัดแกน",
    "ระบบขนส่ง Shaft",
    "ระดับน้ำมัน Hydraulic",
    "สภาพพาเลท",
  ],
  dollyChecks = ["ความสะอาดของ Dolly", "การสึกหรอของ Dolly"],
  coreCutterChecks = [
    "ขนาด Paper Core ถูกต้อง",
    "จำนวนครั้งตัดน้อยกว่า 400 ครั้ง",
  ];
type AppRole = "operator" | "viewer" | "admin";
const conditionFields = [
  ["productCode", "Product Code"],
  ["jumboNo", "Jumbo No."],
  ["slitNo", "Slit No."],
  ["width", "Width on Jumbo"],
  ["speed", "Speed"],
  ["ramp", "Ramp Up / Down"],
  ["tension", "Winder Tension"],
  ["pressure", "Rider Roll Pressure"],
  ["torque", "Torque"],
  ["rollLength", "Roll Length"],
  ["rollDiameter", "Roll Diameter"],
] as const;
type Step = "menu" | "setup" | "check" | "condition" | "confirm" | "success";
type Result = "pass" | "fail";
type Condition = Record<string, string> & {
  jumboSlot: string;
  inspectionNo: string;
  replacements: string;
  parameterValues: string;
};
type Replacement = {
  id: string;
  knifeNo: string;
  life: string;
  reason: string;
};
type AreaParameter = {
  id: string;
  code: string;
  name: string;
  dataType: string;
  unit: string;
  required: boolean;
  minValue: number | null;
  maxValue: number | null;
};
type DynamicParameterValue = {
  code: string;
  name: string;
  dataType: string;
  unit: string;
  value: string;
};
const blankCondition = (slot = "1", round = "1"): Condition => ({
  jumboSlot: slot,
  inspectionNo: round,
  productCode: "",
  jumboNo: "",
  slitNo: "",
  width: "",
  speed: "",
  ramp: "",
  tension: "",
  pressure: "",
  torque: "",
  rollLength: "",
  rollDiameter: "",
  crossSection: "",
  knifeNumbers: "",
  replacements: "[]",
  parameterValues: "[]",
});

let sharedMsalClient: PublicClientApplication | null = null;
async function acquireMicrosoftToken() {
  if (!sharedMsalClient) {
    sharedMsalClient = new PublicClientApplication({
      auth: {
        clientId: "497a8bdf-06dc-4e47-9b25-e3fa8b3248d5",
        authority: "https://login.microsoftonline.com/464ae0fc-7c03-447d-9776-fbfad0c89bcf",
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
    await sharedMsalClient.initialize();
  }
  return sharedMsalClient.loginPopup({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    prompt: "select_account",
  });
}
async function refreshMicrosoftToken(account: MsalAccount | null) {
  if (!sharedMsalClient || !account) return acquireMicrosoftToken();
  try {
    const result = await sharedMsalClient.acquireTokenSilent({
      scopes: ["User.Read", "Sites.ReadWrite.All"],
      account,
    });
    return { ...result, account };
  } catch {
    return acquireMicrosoftToken();
  }
}
export default function SlitterApp() {
  const [role, setRole] = useState<AppRole | null>(null),
    [adminView, setAdminView] = useState<"data" | "users" | "parameters">("data"),
    [login, setLogin] = useState({ user: "", pass: "" }),
    [loginError, setLoginError] = useState(""),
    [step, setStep] = useState<Step>("menu"),
    [area, setArea] = useState(""),
    [machine, setMachine] = useState(""),
    [crew, setCrew] = useState(""),
    [period, setPeriod] = useState(""),
    [checks, setChecks] = useState<Record<string, Result>>({}),
    [condition, setCondition] = useState<Condition>(blankCondition()),
    [saved, setSaved] = useState<Condition[]>([]),
    [conditionCount, setConditionCount] = useState(6),
    [deletedCards, setDeletedCards] = useState<number[]>([]),
    [activeKey, setActiveKey] = useState<string | null>(null),
    [editingCondition, setEditingCondition] = useState(false),
    [copySource, setCopySource] = useState(""),
    [savingCondition, setSavingCondition] = useState(false),
    [saving, setSaving] = useState(false),
    [loadedDocumentNo, setLoadedDocumentNo] = useState(""),
    [loadQuery, setLoadQuery] = useState(""),
    [loadingRecord, setLoadingRecord] = useState(false),
    [loadError, setLoadError] = useState(""),
    [recordId, setRecordId] = useState<number | string | null>(null),
    [accessToken, setAccessToken] = useState(""),
    [microsoftAccount, setMicrosoftAccount] = useState<MsalAccount | null>(null),
    [areaParameters, setAreaParameters] = useState<AreaParameter[]>([]),
    [parameterError, setParameterError] = useState("");
  const operator = login.user.trim(),
    machines =
      area === "TH3" ? ["3A", "3B"] : area === "TH4" ? ["4A", "4B"] : [],
    allChecks = [...machineChecks, ...dollyChecks, ...coreCutterChecks],
    checkedCount = Object.keys(checks).length,
    failCount = Object.values(checks).filter((x) => x === "fail").length,
    complete = checkedCount === allChecks.length;
  const documentNo = useMemo(() => {
    if (loadedDocumentNo) return loadedDocumentNo;
    if (!crew || !period) return "";
    const d = new Date(),
      date = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
    return `${date}${period === "Day" ? "D" : "N"}-${machine || "SL"}`;
  }, [crew, period, machine, loadedDocumentNo]);
  const visibleAreaParameters = useMemo(() => {
    const fixedNames = new Set(conditionFields.map(([, label]) => label.toLowerCase()));
    return areaParameters.filter(
      (parameter) => !fixedNames.has(parameter.name.trim().toLowerCase()),
    );
  }, [areaParameters]);
  const replacements: Replacement[] = useMemo(() => {
      try {
        return JSON.parse(condition.replacements || "[]");
      } catch {
        return [];
      }
    }, [condition.replacements]),
    parameterValues: DynamicParameterValue[] = (() => {
      try {
        return JSON.parse(condition.parameterValues || "[]");
      } catch {
        return [];
      }
    })(),
    replacementComplete = replacements.every(
      (x) => x.knifeNo && x.life.trim() && x.reason.trim(),
    ),
    conditionComplete =
      conditionFields.every(([k]) => condition[k].trim()) &&
      !!condition.crossSection &&
      !!condition.knifeNumbers &&
      visibleAreaParameters.every(
        (parameter) =>
          !parameter.required ||
          !!parameterValues.find((item) => item.code === parameter.code)?.value.trim(),
      ) &&
      replacementComplete,
    conditionKey = (c: Condition) => `${c.jumboSlot}-${c.inspectionNo}`;

  useEffect(() => {
    if (!area || !accessToken) {
      return;
    }
    let active = true;
    loadSharePointParameters(area, accessToken)
      .then((parameters) => {
        if (active) setAreaParameters(parameters);
      })
      .catch((reason) => {
        if (active) setParameterError(reason instanceof Error ? reason.message : "โหลด Parameter ไม่สำเร็จ");
      });
    return () => {
      active = false;
    };
  }, [area, accessToken]);

  function setParameterValue(parameter: AreaParameter, value: string) {
    const next = [
      ...parameterValues.filter((item) => item.code !== parameter.code),
      {
        code: parameter.code,
        name: parameter.name,
        dataType: parameter.dataType,
        unit: parameter.unit,
        value,
      },
    ];
    setCondition({ ...condition, parameterValues: JSON.stringify(next) });
  }
  async function signIn(e: React.FormEvent) {
    e.preventDefault();
    setLoginError("");
    if (!login.user.trim() || !login.pass) {
      setLoginError("กรุณากรอก Username และ Password");
      return;
    }
    try {
      const microsoft = await acquireMicrosoftToken();
      const profile = await getUserProfile(
        microsoft.account,
        microsoft.accessToken,
        login.user,
        login.pass,
      );
      setRole(profile.role);
      setAccessToken(microsoft.accessToken);
      setMicrosoftAccount(microsoft.account);
      setLogin({ user: profile.username || login.user, pass: "" });
      if (profile.team) setCrew(profile.team);
      setAdminView("data");
      setLoginError("");
    } catch (reason) {
      setLoginError(
        reason instanceof Error ? reason.message : "เข้าสู่ระบบไม่สำเร็จ",
      );
    }
  }
  function signOut() {
    setRole(null);
    setAccessToken("");
    setMicrosoftAccount(null);
    setAdminView("data");
  }
  function reset() {
    setStep("menu");
    setArea("");
    setMachine("");
    setCrew("");
    setPeriod("");
    setChecks({});
    setCondition(blankCondition());
    setSaved([]);
    setConditionCount(6);
    setDeletedCards([]);
    setLoadedDocumentNo("");
    setLoadQuery("");
    setLoadError("");
    setActiveKey(null);
    setRecordId(null);
  }
  function startNewRecord() {
    reset();
    setStep("setup");
  }
  async function loadRecord() {
    const query = loadQuery.trim().toUpperCase();
    if (!query) return setLoadError("กรุณากรอกเลขที่เอกสาร");
    setLoadingRecord(true);
    setLoadError("");
    try {
      throw new Error("ฟังก์ชันโหลดเอกสารเดิมจะเปิดใช้ในขั้นตอน Revision");
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : "โหลดเอกสารไม่สำเร็จ");
    } finally {
      setLoadingRecord(false);
    }
  }
  function openCondition(slot: string, round: string) {
    const found = saved.find(
      (x) => x.jumboSlot === slot && x.inspectionNo === round,
    );
    setCondition(found || blankCondition(slot, round));
    setActiveKey(`${slot}-${round}`);
    setEditingCondition(!found);
  }
  function cardNo(c: Condition) {
    return (Number(c.jumboSlot) - 1) * 2 + Number(c.inspectionNo);
  }
  function toggleKnife(no: string) {
    const list = condition.knifeNumbers
        ? condition.knifeNumbers.split(",")
        : [],
      next = list.includes(no)
        ? list.filter((x) => x !== no)
        : [...list, no].sort();
    setCondition({ ...condition, knifeNumbers: next.join(",") });
  }
  function copyKnives() {
    const source = saved.find((x) => conditionKey(x) === copySource);
    if (source)
      setCondition({ ...condition, knifeNumbers: source.knifeNumbers || "" });
  }
  function setReplacements(rows: Replacement[]) {
    setCondition({ ...condition, replacements: JSON.stringify(rows) });
  }
  function addReplacement() {
    setReplacements([
      ...replacements,
      { id: String(Date.now()), knifeNo: "", life: "", reason: "" },
    ]);
  }
  function updateReplacement(
    id: string,
    key: keyof Replacement,
    value: string,
  ) {
    setReplacements(
      replacements.map((x) => (x.id === id ? { ...x, [key]: value } : x)),
    );
  }
  function deleteReplacement(id: string) {
    setReplacements(replacements.filter((x) => x.id !== id));
  }
  async function deleteConditionCard(no: number, slot: string, round: string) {
    if (!confirm(`ลบ Condition ${no} ใช่หรือไม่?`)) return;
    const key = `${slot}-${round}`;
    setSaved((items) =>
      items
        .filter((item) => conditionKey(item) !== key)
        .map((item) => {
          const currentNo = cardNo(item);
          if (currentNo < no) return item;
          const nextNo = currentNo - 1;
          return {
            ...item,
            jumboSlot: String(Math.floor((nextNo - 1) / 2) + 1),
            inspectionNo: String(((nextNo - 1) % 2) + 1),
          };
        }),
    );
    setConditionCount((count) => Math.max(1, count - 1));
    setDeletedCards([]);
    if (activeKey === key) setActiveKey(null);
  }
  async function saveCondition() {
    setSavingCondition(true);
    try {
      setSaved((x) => [
        ...x.filter((y) => conditionKey(y) !== conditionKey(condition)),
        condition,
      ]);
      setEditingCondition(false);
    } catch (e) {
      alert(e instanceof Error ? e.message : "บันทึกไม่สำเร็จ");
    } finally {
      setSavingCondition(false);
    }
  }
  async function submit() {
    setSaving(true);
    try {
      const microsoft = await refreshMicrosoftToken(microsoftAccount);
      setAccessToken(microsoft.accessToken);
      setMicrosoftAccount(microsoft.account);
      const record = await saveProductionRecord(
        {
          documentNo,
          area,
          machine,
          crew,
          period,
          operator,
          checks,
          conditions: saved,
        },
        microsoft.accessToken,
      );
      setRecordId(record.id);
      setStep("success");
    } catch (e) {
      alert(e instanceof Error ? e.message : "บันทึกไม่สำเร็จ");
    } finally {
      setSaving(false);
    }
  }
  if (!role)
    return (
      <Login
        login={login}
        setLogin={setLogin}
        error={loginError}
        submit={signIn}
      />
    );
  if (role === "admin")
    return (
      <>
        <Header
          role={role}
          logout={signOut}
          onAdminHome={() => setAdminView("data")}
          onManageUsers={() => setAdminView("users")}
          onManageParameters={() => setAdminView("parameters")}
        />
        {adminView === "data" ? (
          <main className="container admin-page">
            <AdminDataCenter accessToken={accessToken} />
          </main>
        ) : adminView === "users" ? (
          <AdminPanel onBack={() => setAdminView("data")} />
        ) : (
          <ParameterPanel
            onBack={() => setAdminView("data")}
            initialToken={accessToken}
            initialAccount={microsoftAccount}
          />
        )}
      </>
    );
  if (role === "viewer")
    return (
      <>
        <Header role={role} logout={signOut} />
        <main className="container admin-page">
          <AdminDataCenter accessToken={accessToken} />
        </main>
      </>
    );
  return (
    <>
      <Header role={role} logout={signOut} />
      <main className="container">
        <div className="rolebar">
          <div>
            <b>Operator Workspace</b>
            <small>User: {operator}</small>
          </div>
          <span>OPERATOR</span>
        </div>
        {step === "menu" ? (
          <RecordMenu
            choose={(x) => (x === "slitter" ? startNewRecord() : undefined)}
            query={loadQuery}
            setQuery={setLoadQuery}
            load={loadRecord}
            loading={loadingRecord}
            error={loadError}
          />
        ) : (
          <Steps step={step} />
        )}
        {step === "setup" && (
          <section className="card">
            <p className="kicker">01 · ข้อมูลเริ่มต้น</p>
            <h2>เลือกพื้นที่ เครื่องจักร และกะ</h2>
            <p className="muted">
              กรอกข้อมูลให้ครบเพื่อเริ่มตรวจสอบก่อนเดินเครื่อง
            </p>
            <div className="two options">
              {["TH3", "TH4"].map((x) => (
                <button
                  key={x}
                  className={area === x ? "selected" : ""}
                  onClick={() => {
                    setArea(x);
                    setAreaParameters([]);
                    setParameterError("");
                    setMachine("");
                  }}
                >
                  <strong>{x}</strong>
                  <span>Production Area {x.slice(-1)}</span>
                </button>
              ))}
            </div>
            {area && (
              <>
                <h3>เลือกเครื่องจักร {area}</h3>
                <div className="two compact">
                  {machines.map((x) => (
                    <button
                      key={x}
                      className={machine === x ? "selected" : ""}
                      onClick={() => setMachine(x)}
                    >
                      <strong>{x}</strong>
                      <span>Slitter Machine</span>
                    </button>
                  ))}
                </div>
              </>
            )}
            <div className="form-grid">
              <label>
                ชื่อผู้ปฏิบัติงาน
                <input value={operator} readOnly />
              </label>
              <label>
                กะทีม
                <select value={crew} onChange={(e) => setCrew(e.target.value)}>
                  <option value="">เลือก A–D</option>
                  {["A", "B", "C", "D"].map((x) => (
                    <option key={x}>{x}</option>
                  ))}
                </select>
              </label>
              <label>
                ช่วงเวลา
                <select
                  value={period}
                  onChange={(e) => setPeriod(e.target.value)}
                >
                  <option value="">เลือก Day/Night</option>
                  <option>Day</option>
                  <option>Night</option>
                </select>
              </label>
              <label>
                เลขที่เอกสาร
                <input value={documentNo || "ระบบจะสร้างอัตโนมัติ"} readOnly />
              </label>
            </div>
            <div className="actions top-gap">
              <button className="secondary" onClick={() => setStep("menu")}>
                ← เมนูหลัก
              </button>
              <button
                className="primary"
                disabled={!(area && machine && crew && period)}
                onClick={() => setStep("check")}
              >
                เริ่มตรวจสอบ →
              </button>
            </div>
          </section>
        )}
        {step === "check" && (
          <section className="card">
            <div className="title-row">
              <div>
                <p className="kicker">02 · PRE-START CHECK</p>
                <h2>ตรวจสอบเครื่องจักร, Dolly และชุดตัดแกน</h2>
                <p className="muted">เลือกผลทุกหัวข้อให้ครบ</p>
              </div>
              <button className="link" onClick={() => setStep("setup")}>
                ← กลับ
              </button>
            </div>
            <div className="unit">
              <b>{machine}</b>
              <span>
                {area} · กะ {crew} · {period}
                <small>{documentNo}</small>
              </span>
            </div>
            <div className="check-grid three-groups">
              <CheckGroup
                title="เครื่องจักร"
                items={machineChecks}
                checks={checks}
                setChecks={setChecks}
              />
              <CheckGroup
                title="Dolly"
                items={dollyChecks}
                checks={checks}
                setChecks={setChecks}
              />
              <CheckGroup
                title="ชุดตัดแกน (Paper Core Cutter)"
                items={coreCutterChecks}
                checks={checks}
                setChecks={setChecks}
              />
            </div>
            {failCount > 0 && (
              <div className="fail-notice">
                พบรายการไม่ผ่าน {failCount} รายการ
              </div>
            )}
            <div className="progress-head">
              <span>ตรวจแล้ว</span>
              <b>
                {checkedCount} / {allChecks.length}
              </b>
            </div>
            <div className="progress">
              <i
                style={{ width: `${(checkedCount / allChecks.length) * 100}%` }}
              />
            </div>
            <button
              className="primary right"
              disabled={!complete}
              onClick={() => setStep("condition")}
            >
              ไปบันทึก Condition →
            </button>
          </section>
        )}
        {step === "condition" && (
          <section className="card condition-page">
            <div className="title-row">
              <div>
                <p className="kicker">03 · SLIT CONDITION</p>
                <h2>บันทึก Condition การ Slit</h2>
                <p className="muted">
                  เลือก Condition เพื่อบันทึกข้อมูล
                  หรือเพิ่มการ์ดใหม่ได้ตามต้องการ
                </p>
              </div>
            </div>
            <div className="poster-strip">
              {Array.from({ length: conditionCount }, (_, i) => ({
                slot: String(Math.floor(i / 2) + 1),
                round: String((i % 2) + 1),
                no: i + 1,
              }))
                .filter(({ no }) => !deletedCards.includes(no))
                .map(({ slot, round, no }) => {
                  const item = saved.find(
                      (x) => x.jumboSlot === slot && x.inspectionNo === round,
                    ),
                    key = `${slot}-${round}`;
                  return (
                    <button
                      key={key}
                      className={`condition-poster ${item ? "filled" : "empty"} ${activeKey === key ? "active" : ""}`}
                      onClick={() => openCondition(slot, round)}
                    >
                      <span className="poster-top">
                        CONDITION {no}
                        <span
                          className="delete-card"
                          role="button"
                          tabIndex={0}
                          aria-label={`ลบ Condition ${no}`}
                          onClick={(e) => {
                            e.stopPropagation();
                            deleteConditionCard(no, slot, round);
                          }}
                          onKeyDown={(e) => {
                            if (e.key === "Enter" || e.key === " ") {
                              e.preventDefault();
                              e.stopPropagation();
                              deleteConditionCard(no, slot, round);
                            }
                          }}
                        >
                          ×
                        </span>
                      </span>
                      {item ? (
                        <>
                          <strong>{item.productCode}</strong>
                          <div>
                            <b>Jumbo No.</b>
                            <span>{item.jumboNo}</span>
                          </div>
                          <div>
                            <b>Slit No.</b>
                            <span>{item.slitNo}</span>
                          </div>
                          <i>
                            {item.knifeNumbers
                              ? "ใบมีด " +
                                item.knifeNumbers.split(",").length +
                                " หมายเลข · "
                              : ""}
                            บันทึกแล้ว ✓
                          </i>
                        </>
                      ) : (
                        <>
                          <strong className="plus">+</strong>
                          <p>ข้อมูลว่าง</p>
                          <i>กดเพื่อบันทึก</i>
                        </>
                      )}
                    </button>
                  );
                })}
              <button
                className="condition-poster add-poster"
                onClick={() => setConditionCount((x) => x + 1)}
              >
                <strong className="plus">+</strong>
                <p>เพิ่ม Condition</p>
                <i>สร้างการ์ดใหม่</i>
              </button>
            </div>
            {activeKey && (
              <div className="condition-editor">
                <div className="editor-head">
                  <div>
                    <p className="kicker">CONDITION {cardNo(condition)}</p>
                    <h3>
                      {editingCondition
                        ? saved.some((x) => conditionKey(x) === activeKey)
                          ? "แก้ไข Condition"
                          : "บันทึก Condition"
                        : "ข้อมูลถูกล็อกแล้ว"}
                    </h3>
                  </div>
                  <div>
                    {!editingCondition && (
                      <button
                        className="edit-btn"
                        onClick={() => setEditingCondition(true)}
                      >
                        แก้ไขข้อมูล
                      </button>
                    )}
                    <button
                      className="close-btn"
                      onClick={() => setActiveKey(null)}
                    >
                      ปิด
                    </button>
                  </div>
                </div>
                <div className="condition-grid">
                  {conditionFields.map(([key, label]) => (
                    <label key={key}>
                      {label}
                      <input
                        disabled={!editingCondition}
                        value={condition[key]}
                        onChange={(e) =>
                          setCondition({
                            ...condition,
                            [key]:
                              key === "productCode"
                                ? e.target.value.toUpperCase()
                                : e.target.value,
                          })
                        }
                        className={
                          key === "productCode" ? "uppercase-input" : ""
                        }
                        placeholder={label}
                      />
                    </label>
                  ))}
                  {visibleAreaParameters.map((parameter) => {
                    const value =
                      parameterValues.find((item) => item.code === parameter.code)
                        ?.value || "";
                    return (
                      <label key={parameter.id}>
                        {parameter.name}
                        {parameter.required ? " *" : ""}
                        {parameter.unit ? ` (${parameter.unit})` : ""}
                        <input
                          disabled={!editingCondition}
                          type={parameter.dataType === "Number" ? "number" : "text"}
                          min={parameter.minValue ?? undefined}
                          max={parameter.maxValue ?? undefined}
                          value={value}
                          onChange={(event) =>
                            setParameterValue(parameter, event.target.value)
                          }
                          placeholder={parameter.name}
                        />
                      </label>
                    );
                  })}
                </div>
                {parameterError && (
                  <div className="fail-notice">{parameterError}</div>
                )}
                <div className="cross-section">
                  <span>ตรวจสอบหน้าตัดม้วน</span>
                  <button
                    disabled={!editingCondition}
                    className={
                      condition.crossSection === "normal" ? "normal active" : ""
                    }
                    onClick={() =>
                      setCondition({ ...condition, crossSection: "normal" })
                    }
                  >
                    ✓ ปกติ
                  </button>
                  <button
                    disabled={!editingCondition}
                    className={
                      condition.crossSection === "abnormal"
                        ? "abnormal active"
                        : ""
                    }
                    onClick={() =>
                      setCondition({ ...condition, crossSection: "abnormal" })
                    }
                  >
                    ! ไม่ปกติ
                  </button>
                </div>
                <section className="knife-section">
                  <div className="knife-head">
                    <div>
                      <p className="kicker">KNIFE USAGE · REQUIRED</p>
                      <h3>บันทึกการใช้ใบมีด</h3>
                      <span>เลือกหมายเลขใบมีดที่ใช้ใน Condition นี้</span>
                    </div>
                    <b>
                      {condition.knifeNumbers
                        ? condition.knifeNumbers.split(",").length
                        : 0}{" "}
                      / 50
                    </b>
                  </div>
                  <div className="knife-tools">
                    <div className="copy-row">
                      <select
                        disabled={!editingCondition}
                        value={copySource}
                        onChange={(e) => setCopySource(e.target.value)}
                      >
                        <option value="">
                          คัดลอกหมายเลขจาก Condition อื่น
                        </option>
                        {saved
                          .filter(
                            (x) =>
                              conditionKey(x) !== activeKey && x.knifeNumbers,
                          )
                          .map((x) => (
                            <option
                              key={conditionKey(x)}
                              value={conditionKey(x)}
                            >
                              Condition {cardNo(x)} — {x.productCode} — Jumbo{" "}
                              {x.jumboNo}
                            </option>
                          ))}
                      </select>
                      <button
                        disabled={!editingCondition || !copySource}
                        onClick={copyKnives}
                      >
                        คัดลอก
                      </button>
                    </div>
                    <div className="bulk-buttons">
                      <button
                        disabled={!editingCondition}
                        onClick={() =>
                          setCondition({
                            ...condition,
                            knifeNumbers: Array.from({ length: 50 }, (_, i) =>
                              String(i + 1).padStart(2, "0"),
                            ).join(","),
                          })
                        }
                      >
                        เลือกทั้งหมด
                      </button>
                      <button
                        disabled={!editingCondition}
                        onClick={() =>
                          setCondition({ ...condition, knifeNumbers: "" })
                        }
                      >
                        Clear ทั้งหมด
                      </button>
                    </div>
                  </div>
                  <div className="knife-grid">
                    {Array.from({ length: 50 }, (_, i) =>
                      String(i + 1).padStart(2, "0"),
                    ).map((no) => (
                      <button
                        key={no}
                        disabled={!editingCondition}
                        className={
                          (condition.knifeNumbers || "").split(",").includes(no)
                            ? "selected"
                            : ""
                        }
                        onClick={() => toggleKnife(no)}
                      >
                        {no}
                      </button>
                    ))}
                  </div>
                </section>
                <section className="replacement-section">
                  <div className="replacement-head">
                    <div>
                      <p className="kicker">KNIFE REPLACEMENT · OPTIONAL</p>
                      <h3>บันทึกการเปลี่ยนใบมีด</h3>
                      <span>เพิ่มเฉพาะเมื่อมีการเปลี่ยนใบมีด</span>
                    </div>
                    <button
                      disabled={!editingCondition}
                      onClick={addReplacement}
                    >
                      + เพิ่มบันทึก
                    </button>
                  </div>
                  {replacements.length === 0 ? (
                    <div className="empty-replacement">
                      ยังไม่มีบันทึกการเปลี่ยนใบมีด
                    </div>
                  ) : (
                    <div className="replacement-list">
                      {replacements.map((row, i) => (
                        <div className="replacement-row" key={row.id}>
                          <b>#{i + 1}</b>
                          <label>
                            หมายเลขใบมีด
                            <select
                              disabled={!editingCondition}
                              value={row.knifeNo}
                              onChange={(e) =>
                                updateReplacement(
                                  row.id,
                                  "knifeNo",
                                  e.target.value,
                                )
                              }
                            >
                              <option value="">เลือกหมายเลข</option>
                              {Array.from({ length: 50 }, (_, n) =>
                                String(n + 1).padStart(2, "0"),
                              ).map((no) => (
                                <option key={no}>{no}</option>
                              ))}
                            </select>
                          </label>
                          <label>
                            อายุการใช้งานใบมีด (ชั่วโมง)
                            <input
                              disabled={!editingCondition}
                              value={row.life}
                              onChange={(e) =>
                                updateReplacement(
                                  row.id,
                                  "life",
                                  e.target.value,
                                )
                              }
                              placeholder="ระบุอายุ"
                            />
                          </label>
                          <label>
                            สาเหตุที่เปลี่ยน
                            <input
                              disabled={!editingCondition}
                              value={row.reason}
                              onChange={(e) =>
                                updateReplacement(
                                  row.id,
                                  "reason",
                                  e.target.value,
                                )
                              }
                              placeholder="ระบุสาเหตุ"
                            />
                          </label>
                          <button
                            className="delete-row"
                            disabled={!editingCondition}
                            onClick={() => deleteReplacement(row.id)}
                          >
                            ลบ
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </section>
                {editingCondition ? (
                  <div className="editor-actions">
                    <button
                      className="primary"
                      disabled={!conditionComplete || savingCondition}
                      onClick={saveCondition}
                    >
                      {savingCondition
                        ? "กำลังบันทึก..."
                        : "บันทึกและล็อกข้อมูล"}
                    </button>
                  </div>
                ) : (
                  <div className="locked-note">
                    🔒 บันทึกแล้ว — ต้องกด “แก้ไขข้อมูล” ก่อนจึงจะเปลี่ยนแปลงได้
                  </div>
                )}
              </div>
            )}
            <div className="actions top-gap">
              <button className="secondary" onClick={() => setStep("check")}>
                ← กลับ
              </button>
              <button
                className="primary"
                disabled={saved.length === 0}
                onClick={() => setStep("confirm")}
              >
                ตรวจสอบข้อมูลทั้งหมด ({saved.length}/{conditionCount}) →
              </button>
            </div>
          </section>
        )}
        {step === "confirm" && (
          <section className="card">
            <p className="kicker">04 · CONFIRMATION</p>
            <h2>Confirm ข้อมูลทั้งหมด</h2>
            <p className="muted">ตรวจสอบก่อนส่งข้อมูลเข้า Database</p>
            <div className="summary">
              <Summary label="เลขที่เอกสาร" value={documentNo} />
              <Summary
                label="พื้นที่ / เครื่อง"
                value={`${area} / ${machine}`}
              />
              <Summary label="ผู้ปฏิบัติงาน" value={operator} />
              <Summary label="กะ" value={`${crew} · ${period}`} />
              <Summary
                label="ผลผ่าน"
                value={`${checkedCount - failCount} รายการ`}
              />
              <Summary
                label="ผลไม่ผ่าน"
                value={`${failCount} รายการ`}
                good={failCount === 0}
              />
              <Summary label="Condition" value={`${saved.length} รายการ`} />
              <Summary label="สถานะ" value="พร้อมบันทึก" good />
            </div>
            <div className="condition-review">
              {saved.map((x) => (
                <div key={conditionKey(x)}>
                  <b>
                    Jumbo {x.jumboSlot} · ครั้งที่ {x.inspectionNo}
                  </b>
                  <span>
                    {x.productCode} · Jumbo No. {x.jumboNo} · Slit {x.slitNo}
                  </span>
                </div>
              ))}
            </div>
            <div className="actions">
              <button
                className="secondary"
                onClick={() => setStep("condition")}
              >
                ← กลับไปแก้ไข
              </button>
              <button className="primary" disabled={saving} onClick={submit}>
                {saving ? "กำลังบันทึก..." : "ยืนยันส่งเข้า Database ✓"}
              </button>
            </div>
          </section>
        )}
        {step === "success" && (
          <section className="card success">
            <div>✓</div>
            <p className="kicker">05 · COMPLETE</p>
            <h2>บันทึกข้อมูลสำเร็จ</h2>
            <p>
              เอกสาร <b>{documentNo}</b> ถูกส่งเข้า Database แล้ว
            </p>
            <small>Record ID: {recordId}</small>
            <button className="primary" onClick={reset}>
              กลับเมนูหลัก
            </button>
          </section>
        )}
      </main>
    </>
  );
}
function Login({
  login,
  setLogin,
  error,
  submit,
}: {
  login: { user: string; pass: string };
  setLogin: (x: { user: string; pass: string }) => void;
  error: string;
  submit: (e: React.FormEvent) => void;
}) {
  return (
    <main className="login">
      <section className="login-side">
        <div className="logo">SF</div>
        <p>PRODUCTION RECORD SYSTEM</p>
        <h1>
          SlitterFlow
          <br />
          Control
        </h1>
        <span>ระบบบันทึกและควบคุมงาน Slitter</span>
      </section>
      <form className="login-card" onSubmit={submit}>
        <h2>เข้าสู่ระบบ</h2>
        <p>กรอก Username และ Password จากนั้นยืนยันบัญชี Microsoft 365</p>
        <label>
          Username
          <input
            value={login.user}
            onChange={(e) => setLogin({ ...login, user: e.target.value })}
          />
        </label>
        <label>
          Password
          <input
            type="password"
            value={login.pass}
            onChange={(e) => setLogin({ ...login, pass: e.target.value })}
          />
        </label>
        <small className="error">{error}</small>
        <button className="primary">เข้าสู่ระบบ →</button>
        <aside>
          <b>Username + Password + Microsoft 365</b>
          <span>ระบบจะตรวจสิทธิ์จาก Users List ใน SharePoint</span>
        </aside>
      </form>
    </main>
  );
}
function Header({
  role,
  logout,
  onAdminHome,
  onManageUsers,
  onManageParameters,
}: {
  role: string;
  logout: () => void;
  onAdminHome?: () => void;
  onManageUsers?: () => void;
  onManageParameters?: () => void;
}) {
  const [menuOpen, setMenuOpen] = useState(false);
  return (
    <header>
      <div className="brand">
        <b>SF</b>
        <span>
          <strong>SlitterFlow</strong>
          <small>CONTROL</small>
        </span>
      </div>
      <div className="account">
        {role === "admin" ? (
          <div className="admin-account-menu">
            <button
              className="admin-account-trigger"
              onClick={() => setMenuOpen((value) => !value)}
              aria-expanded={menuOpen}
            >
              ADMIN ▾
            </button>
            {menuOpen && (
              <div className="admin-account-dropdown">
                <button
                  onClick={() => {
                    onAdminHome?.();
                    setMenuOpen(false);
                  }}
                >
                  ข้อมูลการผลิต
                </button>
                <button
                  onClick={() => {
                    onManageUsers?.();
                    setMenuOpen(false);
                  }}
                >
                  จัดการผู้ใช้งาน
                </button>
                <button
                  onClick={() => {
                    onManageParameters?.();
                    setMenuOpen(false);
                  }}
                >
                  จัดการ Parameter
                </button>
                <button className="logout-menu" onClick={logout}>
                  ออกจากระบบ
                </button>
              </div>
            )}
          </div>
        ) : (
          <>
            <span>{role.toUpperCase()}</span>
            <button onClick={logout}>ออกจากระบบ</button>
          </>
        )}
      </div>
    </header>
  );
}
type AdminUser = {
  id: number;
  username: string;
  name: string;
  role: "operator" | "viewer" | "admin";
  active: boolean;
  permissions: {
    slitter3: boolean;
    slitter4: boolean;
    autoPack: boolean;
    loadRecord: boolean;
  };
  teams: string[];
  areas: string[];
};
function AdminPanel({ onBack }: { onBack: () => void }) {
  const baseAccess = {
    slitter3: true,
    slitter4: true,
    autoPack: false,
    loadRecord: true,
  };
  const [users, setUsers] = useState<AdminUser[]>([
    {
      id: 1,
      username: "operator",
      name: "Production Operator",
      role: "operator",
      active: true,
      permissions: baseAccess,
      teams: ["A"],
      areas: ["TH3", "TH4"],
    },
    {
      id: 2,
      username: "viewer",
      name: "Production Viewer",
      role: "viewer",
      active: true,
      permissions: { ...baseAccess, loadRecord: false },
      teams: [],
      areas: ["TH3", "TH4"],
    },
    {
      id: 3,
      username: "admin",
      name: "System Administrator",
      role: "admin",
      active: true,
      permissions: {
        slitter3: true,
        slitter4: true,
        autoPack: true,
        loadRecord: true,
      },
      teams: [],
      areas: ["TH3", "TH4"],
    },
  ]);
  const [search, setSearch] = useState("");
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [tempPassword, setTempPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [newUser, setNewUser] = useState({
    username: "",
    name: "",
    role: "operator" as AdminUser["role"],
    team: "A",
  });
  const filtered = users.filter((user) =>
    `${user.username} ${user.name} ${user.role}`
      .toLowerCase()
      .includes(search.toLowerCase()),
  );
  function addUser() {
    if (!newUser.username.trim() || !newUser.name.trim()) return;
    setUsers((items) => [
      ...items,
      {
        id: Date.now(),
        username: newUser.username.trim().toLowerCase(),
        name: newUser.name.trim(),
        role: newUser.role,
        active: true,
        permissions: baseAccess,
        teams: newUser.role === "operator" ? [newUser.team] : [],
        areas: ["TH3", "TH4"],
      },
    ]);
    setNewUser({ username: "", name: "", role: "operator", team: "A" });
  }
  const selected = users.find((user) => user.id === selectedId);
  function updateSelected(change: Partial<AdminUser>) {
    if (!selectedId) return;
    setUsers((items) =>
      items.map((user) =>
        user.id === selectedId ? { ...user, ...change } : user,
      ),
    );
  }
  function toggleList(key: "areas", value: string) {
    if (!selected) return;
    const list = selected[key];
    updateSelected({
      [key]: list.includes(value)
        ? list.filter((item) => item !== value)
        : [...list, value],
    });
  }
  function generatePassword() {
    const value = `SF-${crypto.getRandomValues(new Uint32Array(1))[0].toString(36).slice(0, 7).toUpperCase()}!`;
    setTempPassword(value);
    setShowPassword(true);
  }
  function deleteUser(user: AdminUser) {
    const adminCount = users.filter((item) => item.role === "admin").length;
    if (user.role === "admin" && adminCount === 1) {
      alert("ไม่สามารถลบ Admin คนสุดท้ายของระบบได้");
      return;
    }
    if (
      !window.confirm(
        `ยืนยันลบผู้ใช้ ${user.name} (@${user.username}) ออกจากระบบ?`,
      )
    ) {
      return;
    }
    const remaining = users.filter((item) => item.id !== user.id);
    setUsers(remaining);
    if (selectedId === user.id) {
      setSelectedId(remaining[0]?.id ?? null);
      setTempPassword("");
    }
  }
  return (
    <main className="container admin-page">
      <button className="admin-back" onClick={onBack}>
        ← กลับหน้าข้อมูลการผลิต
      </button>
      <div className="admin-hero">
        <div>
          <p className="kicker">ADMIN CONTROL</p>
          <h1>จัดการผู้ใช้งานและสิทธิ์</h1>
          <p>กำหนดบทบาทและสถานะการเข้าใช้งานของผู้ใช้ทั้งหมด</p>
        </div>
        <div className="admin-stats">
          <span>
            <b>{users.length}</b>ผู้ใช้ทั้งหมด
          </span>
          <span>
            <b>{users.filter((x) => x.active).length}</b>เปิดใช้งาน
          </span>
          <span>
            <b>{users.filter((x) => x.role === "admin").length}</b>ผู้ดูแลระบบ
          </span>
        </div>
      </div>
      <section className="card admin-card">
        <div className="admin-toolbar">
          <div>
            <h2>บัญชีผู้ใช้งาน</h2>
            <span>เลือกผู้ใช้เพื่อกำหนดสิทธิ์ ทีม พื้นที่ และรหัสผ่าน</span>
          </div>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="ค้นหาชื่อหรือบทบาท"
          />
        </div>
        <div className="user-table">
          <div className="user-row user-head">
            <span>ผู้ใช้งาน</span>
            <span>บทบาท</span>
            <span>สถานะ</span>
            <span>การจัดการ</span>
          </div>
          {filtered.map((user) => (
            <div className="user-row" key={user.id}>
              <div className="user-name">
                <b>{user.name}</b>
                <small>@{user.username}</small>
              </div>
              <select
                value={user.role}
                onChange={(e) =>
                  setUsers((items) =>
                    items.map((x) =>
                      x.id === user.id
                        ? (() => {
                            const nextRole = e.target
                              .value as AdminUser["role"];
                            return {
                              ...x,
                              role: nextRole,
                              teams:
                                nextRole === "operator"
                                  ? [x.teams[0] || "A"]
                                  : [],
                            };
                          })()
                        : x,
                    ),
                  )
                }
              >
                <option value="operator">Operator</option>
                <option value="viewer">Viewer</option>
                <option value="admin">Admin</option>
              </select>
              <span className={`status-pill ${user.active ? "on" : "off"}`}>
                {user.active ? "ใช้งานอยู่" : "ระงับแล้ว"}
              </span>
              <div className="user-actions">
                <button
                  className="manage"
                  onClick={() => {
                    setSelectedId(user.id);
                    setTempPassword("");
                  }}
                >
                  กำหนดสิทธิ์
                </button>
                <button
                  className={user.active ? "suspend" : "activate"}
                  onClick={() =>
                    setUsers((items) =>
                      items.map((x) =>
                        x.id === user.id ? { ...x, active: !x.active } : x,
                      ),
                    )
                  }
                >
                  {user.active ? "ระงับ" : "เปิด"}
                </button>
                <button
                  className="delete-user"
                  onClick={() => deleteUser(user)}
                  aria-label={`ลบผู้ใช้ ${user.name}`}
                >
                  ลบ
                </button>
              </div>
            </div>
          ))}
        </div>
      </section>
      {selected && (
        <section className="card permission-card">
          <div className="permission-title">
            <div>
              <p className="kicker">USER PERMISSIONS</p>
              <h2>{selected.name}</h2>
              <span>
                @{selected.username} · {selected.role.toUpperCase()}
              </span>
            </div>
            <span className={`status-pill ${selected.active ? "on" : "off"}`}>
              {selected.active ? "ใช้งานอยู่" : "ระงับแล้ว"}
            </span>
          </div>
          <div className="permission-grid">
            <div className="permission-block">
              <h3>สิทธิ์การใช้งานระบบ</h3>
              {(
                [
                  ["slitter3", "Slitter Record — TH3"],
                  ["slitter4", "Slitter Record — TH4"],
                  ["autoPack", "Auto Pack Record"],
                  ["loadRecord", "โหลดเอกสารเพื่อแก้ไข"],
                ] as const
              ).map(([key, label]) => (
                <button
                  key={key}
                  className={selected.permissions[key] ? "enabled" : ""}
                  onClick={() =>
                    updateSelected({
                      permissions: {
                        ...selected.permissions,
                        [key]: !selected.permissions[key],
                      },
                    })
                  }
                >
                  <span>{label}</span>
                  <b>{selected.permissions[key] ? "อนุญาต" : "ไม่อนุญาต"}</b>
                </button>
              ))}
            </div>
            <div className="permission-block">
              {selected.role === "operator" && (
                <>
                  <h3>ทีมที่รับผิดชอบ (เลือกได้ 1 ทีม)</h3>
                  <div className="choice-pills">
                    {["A", "B", "C", "D"].map((team) => (
                      <button
                        key={team}
                        className={selected.teams[0] === team ? "enabled" : ""}
                        onClick={() => updateSelected({ teams: [team] })}
                      >
                        Team {team}
                      </button>
                    ))}
                  </div>
                </>
              )}
              <h3>พื้นที่การผลิต</h3>
              <div className="choice-pills">
                {["TH3", "TH4"].map((area) => (
                  <button
                    key={area}
                    className={selected.areas.includes(area) ? "enabled" : ""}
                    onClick={() => toggleList("areas", area)}
                  >
                    {area}
                  </button>
                ))}
              </div>
            </div>
            <div className="password-block">
              <h3>จัดการรหัสผ่าน</h3>
              <p>ตั้งรหัสใหม่หรือสร้างรหัสชั่วคราว ระบบจะไม่แสดงรหัสผ่านเดิม</p>
              <div className="password-field">
                <input
                  type={showPassword ? "text" : "password"}
                  value={tempPassword}
                  onChange={(e) => setTempPassword(e.target.value)}
                  placeholder="รหัสผ่านใหม่"
                />
                <button onClick={() => setShowPassword((value) => !value)}>
                  {showPassword ? "ซ่อน" : "ดู"}
                </button>
              </div>
              <div className="password-actions">
                <button onClick={generatePassword}>สร้างรหัสชั่วคราว</button>
                <button
                  className="primary"
                  disabled={!tempPassword.trim()}
                  onClick={() => {
                    alert("ตั้งรหัสผ่านใหม่แล้ว");
                    setTempPassword("");
                  }}
                >
                  บันทึกรหัสใหม่
                </button>
              </div>
            </div>
          </div>
        </section>
      )}
      <section className="card add-user-card">
        <div>
          <p className="kicker">ADD USER</p>
          <h2>เพิ่มผู้ใช้งานใหม่</h2>
        </div>
        <input
          value={newUser.name}
          onChange={(e) => setNewUser({ ...newUser, name: e.target.value })}
          placeholder="ชื่อผู้ใช้งาน"
        />
        <input
          value={newUser.username}
          onChange={(e) => setNewUser({ ...newUser, username: e.target.value })}
          placeholder="Username"
        />
        <select
          value={newUser.role}
          onChange={(e) =>
            setNewUser({
              ...newUser,
              role: e.target.value as AdminUser["role"],
            })
          }
        >
          <option value="operator">Operator</option>
          <option value="viewer">Viewer</option>
          <option value="admin">Admin</option>
        </select>
        {newUser.role === "operator" && (
          <select
            value={newUser.team}
            onChange={(e) => setNewUser({ ...newUser, team: e.target.value })}
            aria-label="ทีมของ Operator"
          >
            {["A", "B", "C", "D"].map((team) => (
              <option key={team} value={team}>
                Team {team}
              </option>
            ))}
          </select>
        )}
        <button className="primary" onClick={addUser}>
          + เพิ่มผู้ใช้
        </button>
      </section>
    </main>
  );
}

type ParameterItem = {
  id: string;
  area: "TH3" | "TH4";
  category: string;
  name: string;
  inputType: string;
  minValue: string;
  maxValue: string;
  unit: string;
  active: boolean;
};

type MsalAccount = AccountInfo;

function ParameterPanel({
  onBack,
  initialToken,
  initialAccount,
}: {
  onBack: () => void;
  initialToken: string;
  initialAccount: MsalAccount | null;
}) {
  const [area, setArea] = useState<"TH3" | "TH4">("TH3");
  const [items, setItems] = useState<ParameterItem[]>([]);
  const [token, setToken] = useState(initialToken);
  const [account, setAccount] = useState<MsalAccount | null>(initialAccount);
  const [listId, setListId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [form, setForm] = useState({
    name: "",
    code: "",
    inputType: "Text",
    minValue: "",
    maxValue: "",
    unit: "",
    required: true,
  });
  const siteId = "fitesacnc.sharepoint.com,9c41b493-3a6a-4c01-8d43-26b203783d58,1e97269c-5340-47e3-b05b-6170a179d56b";

  useEffect(() => {
    if (initialToken) void loadParameters(initialToken, "");
  }, [initialToken]); // eslint-disable-line react-hooks/exhaustive-deps

  async function graph(path: string, accessToken: string, init: RequestInit = {}) {
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...(init.headers || {}),
      },
    });
    const data = response.status === 204 ? null : await response.json();
    if (!response.ok) throw new Error(data?.error?.message || "เชื่อมต่อ SharePoint ไม่สำเร็จ");
    return data;
  }

  async function findParameterList(accessToken: string) {
    const data = await graph(`/sites/${siteId}/lists?$select=id,displayName`, accessToken);
    const found = (data.value || []).find((list: { id: string; displayName: string }) => {
      const name = list.displayName.toLowerCase().replace(/[ _-]/g, "");
      return name === "parametermaster" || name === "parameters";
    });
    if (!found) throw new Error("ไม่พบ List: Parameter Master ใน SharePoint");
    setListId(found.id);
    return found.id as string;
  }

  async function loadParameters(accessToken = token, targetListId = listId) {
    if (!accessToken) return;
    setLoading(true);
    setError("");
    try {
      const id = targetListId || (await findParameterList(accessToken));
      const data = await graph(
        `/sites/${siteId}/lists/${id}/items?$expand=fields&$top=999`,
        accessToken,
      );
      setItems(
        (data.value || []).map((row: { id: string; fields: Record<string, unknown> }) => ({
          id: row.id,
          area: String(row.fields.field_2 || "TH3") as "TH3" | "TH4",
          category: String(row.fields.field_1 || ""),
          name: String(row.fields.Title || row.fields.ChecklistName || ""),
          inputType: String(row.fields.field_3 || "Text"),
          minValue: row.fields.field_7 == null ? "" : String(row.fields.field_7),
          maxValue: row.fields.field_8 == null ? "" : String(row.fields.field_8),
          unit: String(row.fields.field_4 || ""),
          active: row.fields.field_10 !== false,
        })),
      );
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "โหลด Parameter ไม่สำเร็จ");
    } finally {
      setLoading(false);
    }
  }

  async function connectMicrosoft() {
    setLoading(true);
    setError("");
    try {
      const result = await acquireMicrosoftToken();
      setAccount(result.account);
      setToken(result.accessToken);
      const id = await findParameterList(result.accessToken);
      await loadParameters(result.accessToken, id);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "เข้าสู่ระบบ Microsoft 365 ไม่สำเร็จ");
      setLoading(false);
    }
  }

  async function addParameter() {
    if (!form.name.trim() || !token || !listId) return;
    setLoading(true);
    setError("");
    try {
      const fields: Record<string, string | number | boolean | null> = {
        Title: form.name.trim(),
        field_1: form.code.trim(),
        field_2: area,
        field_3: form.inputType,
        field_4: form.unit.trim(),
        field_6: form.required,
        field_10: true,
        InactiveDate: null,
      };
      if (form.minValue !== "") fields.field_7 = Number(form.minValue);
      if (form.maxValue !== "") fields.field_8 = Number(form.maxValue);
      await graph(`/sites/${siteId}/lists/${listId}/items`, token, {
        method: "POST",
        body: JSON.stringify({ fields }),
      });
      setForm({ name: "", code: "", inputType: "Text", minValue: "", maxValue: "", unit: "", required: true });
      await loadParameters();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "เพิ่ม Parameter ไม่สำเร็จ");
      setLoading(false);
    }
  }

  async function setParameterActive(item: ParameterItem, active: boolean) {
    if (!token || !listId) return;
    if (!active && !window.confirm(`ยกเลิกการใช้งาน ${item.name} ใน ${item.area}?`)) return;
    setLoading(true);
    setError("");
    try {
      await graph(`/sites/${siteId}/lists/${listId}/items/${item.id}/fields`, token, {
        method: "PATCH",
        body: JSON.stringify({ field_10: active, InactiveDate: active ? null : new Date().toISOString(), field_13: active ? "" : (account?.username || "Admin") }),
      });
      await loadParameters();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "แก้ไขสถานะไม่สำเร็จ");
      setLoading(false);
    }
  }

  const visible = items.filter((item) => item.area === area);
  return (
    <main className="container admin-page parameter-page">
      <button className="admin-back" onClick={onBack}>← กลับหน้าข้อมูลการผลิต</button>
      <div className="admin-hero">
        <div><p className="kicker">PARAMETER MASTER</p><h1>จัดการ Parameter</h1><p>เพิ่มหรือยกเลิก Parameter แยกตามพื้นที่ TH3 และ TH4</p></div>
        <div className="admin-stats"><span><b>{visible.length}</b>รายการใน {area}</span><span><b>{visible.filter((x) => x.active).length}</b>ใช้งานอยู่</span></div>
      </div>
      {!token ? (
        <section className="card parameter-connect">
          <h2>เชื่อมต่อ SharePoint</h2>
          <p>Admin ต้องยืนยันด้วยบัญชี Microsoft 365 ก่อนจัดการ Parameter</p>
          <button className="primary" onClick={connectMicrosoft} disabled={loading}>{loading ? "กำลังเชื่อมต่อ..." : "เข้าสู่ระบบ Microsoft 365"}</button>
          {error && <p className="error">{error}</p>}
        </section>
      ) : (
        <>
          <section className="card parameter-toolbar">
            <div className="parameter-area-tabs">
              {(["TH3", "TH4"] as const).map((value) => <button key={value} className={area === value ? "selected" : ""} onClick={() => setArea(value)}>{value}</button>)}
            </div>
            <span>เชื่อมต่อแล้ว: {account?.name || account?.username}</span>
          </section>
          <section className="card parameter-form">
            <div><p className="kicker">ADD PARAMETER · {area}</p><h2>เพิ่ม Parameter ใหม่</h2></div>
            <div className="parameter-form-grid">
              <label>ชื่อ Parameter<input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></label>
              <label>Parameter Code<input value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value.toUpperCase() })} /></label>
              <label>รูปแบบข้อมูล<select value={form.inputType} onChange={(e) => setForm({ ...form, inputType: e.target.value })}>{["Text", "Number", "Choice"].map((x) => <option key={x}>{x}</option>)}</select></label>
              <label>ค่าต่ำสุด<input type="number" value={form.minValue} onChange={(e) => setForm({ ...form, minValue: e.target.value })} /></label>
              <label>ค่าสูงสุด<input type="number" value={form.maxValue} onChange={(e) => setForm({ ...form, maxValue: e.target.value })} /></label>
              <label>หน่วย<input value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} /></label>
              <label>จำเป็นต้องกรอก<select value={form.required ? "yes" : "no"} onChange={(e) => setForm({ ...form, required: e.target.value === "yes" })}><option value="yes">ใช่</option><option value="no">ไม่ใช่</option></select></label>
              <button className="primary" onClick={addParameter} disabled={loading || !form.name.trim() || !form.code.trim()}>+ เพิ่มใน {area}</button>
            </div>
          </section>
          <section className="card parameter-list">
            <div className="admin-toolbar"><div><h2>Parameter ของ {area}</h2><span>รายการที่ยกเลิกจะถูกเก็บไว้เป็นประวัติ</span></div><button className="manage" onClick={() => loadParameters()} disabled={loading}>รีเฟรช</button></div>
            {error && <p className="error">{error}</p>}
            {loading ? <div className="data-state">กำลังโหลดข้อมูล...</div> : visible.map((item) => (
              <div className={`parameter-row ${item.active ? "" : "inactive"}`} key={item.id}>
                <div><b>{item.name}</b><small>{item.category} · {item.inputType}{item.unit ? ` · ${item.unit}` : ""}</small></div>
                <span className={`status-pill ${item.active ? "on" : "off"}`}>{item.active ? "ใช้งานอยู่" : "ยกเลิกแล้ว"}</span>
                <button className={item.active ? "suspend" : "activate"} onClick={() => setParameterActive(item, !item.active)}>{item.active ? "ยกเลิก" : "เปิดใช้งาน"}</button>
              </div>
            ))}
            {!loading && visible.length === 0 && <div className="data-state">ยังไม่มี Parameter ใน {area}</div>}
          </section>
        </>
      )}
    </main>
  );
}

type AdminConditionRecord = {
  id: number;
  documentNo: string;
  date: string;
  area: string;
  machine: string;
  operator: string;
  updatedAt: string;
  conditionNo: number;
  productCode: string;
  jumboNo: string;
  slitNo: string;
  width: string;
  speed: string;
  ramp: string;
  tension: string;
  pressure: string;
  torque: string;
  rollLength: string;
  rollDiameter: string;
  crossSection: string;
  knifeNumbers: string[];
  replacements: Replacement[];
};

function AdminDataCenter({ accessToken }: { accessToken: string }) {
  const [tab, setTab] = useState<"condition" | "knives" | "replacement">(
    "condition",
  );
  const [filters, setFilters] = useState({
    date: "",
    productCode: "",
    area: "",
    machine: "",
  });
  const [records, setRecords] = useState<AdminConditionRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);
  const [error, setError] = useState("");

  async function loadData(activeFilters = filters) {
    setHasSearched(true);
    setLoading(true);
    setError("");
    try {
      setRecords(await loadAdminRecords(activeFilters, accessToken));
    } catch (reason) {
      setError(
        reason instanceof Error ? reason.message : "โหลดข้อมูลไม่สำเร็จ",
      );
    } finally {
      setLoading(false);
    }
  }

  const machineOptions =
    filters.area === "TH3"
      ? ["3A", "3B"]
      : filters.area === "TH4"
        ? ["4A", "4B"]
        : ["3A", "3B", "4A", "4B"];
  const replacementRows = records.flatMap((record) =>
    record.replacements.map((replacement, index) => ({
      record,
      replacement,
      key: `${record.id}-${replacement.id || index}`,
    })),
  );
  const conditionHeaders = [
    "วันที่",
    "เลขที่เอกสาร",
    "พื้นที่ / เครื่อง",
    "Condition",
    "Product Code",
    "Jumbo No.",
    "Slit No.",
    "Width on Jumbo",
    "Speed",
    "Ramp Up / Down",
    "Winder Tension",
    "Rider Roll Pressure",
    "Torque",
    "Roll Length",
    "Roll Diameter",
    "หน้าตัดม้วน",
    "ผู้ปฏิบัติงาน",
  ];
  const conditionRows = records.map((record) => [
    formatRecordDate(record.date),
    record.documentNo,
    `${record.area} / ${record.machine}`,
    `Condition ${record.conditionNo}`,
    record.productCode,
    record.jumboNo,
    record.slitNo,
    record.width,
    record.speed,
    record.ramp,
    record.tension,
    record.pressure,
    record.torque,
    record.rollLength,
    record.rollDiameter,
    record.crossSection,
    record.operator,
  ]);
  const knifeHeaders = [
    "วันที่",
    "เลขที่เอกสาร",
    "พื้นที่ / เครื่อง",
    "Condition",
    "Product Code",
    "Jumbo No.",
    "หมายเลขใบมีดที่ใช้",
  ];
  const knifeRows = records.map((record) => [
    formatRecordDate(record.date),
    record.documentNo,
    `${record.area} / ${record.machine}`,
    `Condition ${record.conditionNo}`,
    record.productCode,
    record.jumboNo,
    record.knifeNumbers.join(", ") || "—",
  ]);
  const replacementHeaders = [
    "วันที่",
    "เลขที่เอกสาร",
    "พื้นที่ / เครื่อง",
    "Condition",
    "Product Code",
    "หมายเลขใบมีด",
    "อายุการใช้งาน (ชั่วโมง)",
    "สาเหตุที่เปลี่ยน",
  ];
  const knifeReplacementRows = replacementRows.map(
    ({ record, replacement }) => [
      formatRecordDate(record.date),
      record.documentNo,
      `${record.area} / ${record.machine}`,
      `Condition ${record.conditionNo}`,
      record.productCode,
      replacement.knifeNo || "—",
      replacement.life || "—",
      replacement.reason || "—",
    ],
  );
  const activeTable =
    tab === "condition"
      ? { headers: conditionHeaders, rows: conditionRows, name: "condition" }
      : tab === "knives"
        ? { headers: knifeHeaders, rows: knifeRows, name: "knife-usage" }
        : {
            headers: replacementHeaders,
            rows: knifeReplacementRows,
            name: "knife-replacement",
          };

  function exportCsv() {
    if (!activeTable.rows.length) return;
    const escapeCsv = (value: string) =>
      `"${String(value ?? "").replaceAll('"', '""')}"`;
    const csv = [activeTable.headers, ...activeTable.rows]
      .map((row) => row.map(escapeCsv).join(","))
      .join("\r\n");
    const blob = new Blob(["\uFEFF", csv], {
      type: "text/csv;charset=utf-8;",
    });
    const link = document.createElement("a");
    const filterDate = filters.date || new Date().toISOString().slice(0, 10);
    link.href = URL.createObjectURL(blob);
    link.download = `slitter-${activeTable.name}-${filterDate}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
  }

  return (
    <section className="card admin-data-card">
      <div className="admin-data-title">
        <div>
          <p className="kicker">ADMIN DATA CENTER</p>
          <h2>ข้อมูล Slitter ทั้งหมด</h2>
          <span>สำหรับตรวจสอบข้อมูลเท่านั้น ไม่สามารถแก้ไขได้</span>
        </div>
        <span className="readonly-badge">READ ONLY</span>
      </div>

      <div className="admin-data-filters">
        <label>
          วันที่
          <input
            type="date"
            value={filters.date}
            onChange={(event) =>
              setFilters({ ...filters, date: event.target.value })
            }
          />
        </label>
        <label>
          Product Code
          <input
            value={filters.productCode}
            onChange={(event) =>
              setFilters({
                ...filters,
                productCode: event.target.value.toUpperCase(),
              })
            }
            placeholder="PRODUCT CODE"
          />
        </label>
        <label>
          พื้นที่
          <select
            value={filters.area}
            onChange={(event) =>
              setFilters({ ...filters, area: event.target.value, machine: "" })
            }
          >
            <option value="">ทั้งหมด</option>
            <option value="TH3">TH3</option>
            <option value="TH4">TH4</option>
          </select>
        </label>
        <label>
          เครื่องจักร
          <select
            value={filters.machine}
            onChange={(event) =>
              setFilters({ ...filters, machine: event.target.value })
            }
          >
            <option value="">ทั้งหมด</option>
            {machineOptions.map((item) => (
              <option key={item} value={item}>
                {item}
              </option>
            ))}
          </select>
        </label>
        <button
          className="primary"
          onClick={() => void loadData()}
          disabled={loading}
        >
          {loading ? "กำลังโหลด..." : "กรองข้อมูล"}
        </button>
        <button
          className="clear-filter"
          onClick={() => {
            const empty = {
              date: "",
              productCode: "",
              area: "",
              machine: "",
            };
            setFilters(empty);
            setRecords([]);
            setError("");
            setHasSearched(false);
          }}
        >
          ล้างตัวกรอง
        </button>
      </div>

      <div className="data-category-tabs">
        <button
          className={tab === "condition" ? "active" : ""}
          onClick={() => setTab("condition")}
        >
          Condition <b>{records.length}</b>
        </button>
        <button
          className={tab === "knives" ? "active" : ""}
          onClick={() => setTab("knives")}
        >
          การใช้ใบมีด <b>{records.length}</b>
        </button>
        <button
          className={tab === "replacement" ? "active" : ""}
          onClick={() => setTab("replacement")}
        >
          Knife Replacement <b>{replacementRows.length}</b>
        </button>
      </div>

      <div className="data-export-row">
        <span>ส่งออกข้อมูลตามหมวดและตัวกรองที่เลือก</span>
        <button
          className="export-csv"
          onClick={exportCsv}
          disabled={
            !hasSearched || loading || !!error || !activeTable.rows.length
          }
        >
          Export CSV ({activeTable.rows.length})
        </button>
      </div>

      {!hasSearched ? (
        <div className="data-state">
          เลือกเงื่อนไขที่ต้องการ แล้วกด “กรองข้อมูล” เพื่อเริ่มดึงข้อมูล
        </div>
      ) : error ? (
        <div className="data-state error">
          <span>{error}</span>
        </div>
      ) : loading ? (
        <div className="data-state">กำลังดึงข้อมูลจากฐานข้อมูล...</div>
      ) : (
        <AdminDataTable headers={activeTable.headers} rows={activeTable.rows} />
      )}
    </section>
  );
}

function AdminDataTable({
  headers,
  rows,
}: {
  headers: string[];
  rows: string[][];
}) {
  if (!rows.length)
    return <div className="data-state">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</div>;
  return (
    <div className="admin-data-table-wrap">
      <table className="admin-data-table">
        <thead>
          <tr>
            {headers.map((header) => (
              <th key={header}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) => (
                <td key={`${rowIndex}-${cellIndex}`}>{cell || "—"}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function formatRecordDate(value: string) {
  if (value.length !== 8) return value;
  return `${value.slice(6, 8)}/${value.slice(4, 6)}/${value.slice(0, 4)}`;
}
function RecordMenu({
  choose,
  query,
  setQuery,
  load,
  loading,
  error,
}: {
  choose: (x: string) => void;
  query: string;
  setQuery: (x: string) => void;
  load: () => void;
  loading: boolean;
  error: string;
}) {
  return (
    <section className="card menu-card">
      <p className="kicker">SELECT RECORD</p>
      <h2>เลือกประเภทบันทึก</h2>
      <p className="muted">เลือกระบบที่ต้องการใช้งาน</p>
      <div className="record-menu">
        <button onClick={() => choose("slitter")}>
          <i>SL</i>
          <b>Slitter Record</b>
          <span>บันทึกการตรวจสอบและ Condition การ Slit</span>
          <strong>เริ่มใช้งาน →</strong>
        </button>
        <button className="coming" onClick={() => choose("auto")}>
          <i>AP</i>
          <b>Auto Pack Record</b>
          <span>บันทึกการทำงาน Auto Packing</span>
          <strong>เตรียมระบบ</strong>
        </button>
      </div>
      <div className="load-record">
        <div className="load-icon">↻</div>
        <div className="load-copy">
          <b>โหลดเอกสารเพื่อแก้ไข</b>
          <span>กรอกเลขที่เอกสาร Slitter Record ที่เคยบันทึกไว้</span>
        </div>
        <div className="load-controls">
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value.toUpperCase())}
            onKeyDown={(e) => e.key === "Enter" && load()}
            placeholder="เช่น 20260908D-3A"
            aria-label="เลขที่เอกสาร"
          />
          <button onClick={load} disabled={loading}>
            {loading ? "กำลังโหลด..." : "โหลดข้อมูล"}
          </button>
        </div>
        {error && <small className="load-error">{error}</small>}
      </div>
    </section>
  );
}
function Steps({ step }: { step: Step }) {
  const n = {
    menu: 0,
    setup: 1,
    check: 2,
    condition: 3,
    confirm: 4,
    success: 5,
  }[step];
  return (
    <nav className="steps five">
      {["ข้อมูลเริ่มต้น", "ตรวจสอบ", "Slit Condition", "Confirm", "สำเร็จ"].map(
        (x, i) => (
          <div key={x} className={n >= i + 1 ? "active" : ""}>
            <b>{i + 1}</b>
            <span>{x}</span>
          </div>
        ),
      )}
    </nav>
  );
}
function CheckGroup({
  title,
  items,
  checks,
  setChecks,
}: {
  title: string;
  items: string[];
  checks: Record<string, Result>;
  setChecks: (x: Record<string, Result>) => void;
}) {
  return (
    <div>
      <h3>{title}</h3>
      {items.map((x) => (
        <div className={`check ${checks[x] || ""}`} key={x}>
          <span>{x}</span>
          <div>
            <button
              className={checks[x] === "pass" ? "active" : ""}
              onClick={() => setChecks({ ...checks, [x]: "pass" })}
            >
              ผ่าน
            </button>
            <button
              className={checks[x] === "fail" ? "active" : ""}
              onClick={() => setChecks({ ...checks, [x]: "fail" })}
            >
              ไม่ผ่าน
            </button>
          </div>
        </div>
      ))}
    </div>
  );
}
function Summary({
  label,
  value,
  good,
}: {
  label: string;
  value: string;
  good?: boolean;
}) {
  return (
    <div>
      <span>{label}</span>
      <b className={good ? "good" : ""}>{value}</b>
    </div>
  );
}
