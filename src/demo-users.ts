// Trial accounts only. Local browser storage is not production authentication.
export type TrialRole = "operator" | "manager" | "viewer" | "admin";
export type TrialUser = {
  id: number; username: string; name: string; password: string; role: TrialRole;
  active: boolean; permissions: { slitter3: boolean; slitter4: boolean; autoPack: boolean; loadRecord: boolean; review: boolean };
  areas: string[]; managerName: string;
};
const KEY = "slitterflow-trial-users-v1";
const access = { slitter3: true, slitter4: true, autoPack: false, loadRecord: false, review: false };
const defaults: TrialUser[] = [
  { id: 1, username: "operator", name: "Production Operator", password: "operator123", role: "operator", active: true, permissions: access, areas: ["TH3", "TH4"], managerName: "" },
  { id: 2, username: "manager", name: "Shift Manager", password: "manager123", role: "manager", active: true, permissions: { ...access, review: true }, areas: ["TH3", "TH4"], managerName: "" },
  { id: 3, username: "viewer", name: "Production Viewer", password: "viewer123", role: "viewer", active: true, permissions: access, areas: ["TH3", "TH4"], managerName: "" },
  { id: 4, username: "admin", name: "System Administrator", password: "admin123", role: "admin", active: true, permissions: { ...access, autoPack: true, loadRecord: true }, areas: ["TH3", "TH4"], managerName: "" },
];
export function trialUsers(): TrialUser[] {
  try {
    const stored = JSON.parse(localStorage.getItem(KEY) || "null");
    return Array.isArray(stored) ? stored : defaults;
  } catch { return defaults; }
}
export function writeTrialUsers(users: TrialUser[]) { localStorage.setItem(KEY, JSON.stringify(users)); }
export function managerNames() { return trialUsers().filter(user => user.active && user.role === "manager").map(user => user.name); }
