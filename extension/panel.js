// Side panel: choose the workbook and client, preview the claims, then fill the form
// in the claim website's tab. The reading logic is in src/core.js, the form filling in src/content.js.
import * as core from "./src/core.js";

const DEFAULT_URL = "https://mock-website-theta.vercel.app";
const $ = (id) => document.getElementById(id);

const state = {
  source: null,        // "xlsx" | "json"
  fileName: "",
  sheet: null,         // loadSheet() result (xlsx only)
  clients: [],         // findClients() result (xlsx only)
  dateColumns: [],
  doc: null,           // parsed export (json only)
  result: null,        // {client, rate, claims, problems, warnings}
  groups: [],
  range: [null, null],
  running: false,
  port: null,
};

function log(text) {
  const box = $("log");
  box.textContent += text + "\n";
  box.scrollTop = box.scrollHeight;
}

// --- loading files --------------------------------------------------------------------
$("file").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  try {
    const sheet = core.loadSheet(XLSX, await file.arrayBuffer());
    const clients = core.findClients(sheet);
    if (!clients.length) throw new Error(`No '${core.CLIENT_LABEL}' labels found on sheet '${sheet.title}'.`);
    const { columns } = core.findDateColumns(sheet);
    Object.assign(state, { source: "xlsx", fileName: file.name, sheet, clients, dateColumns: columns, doc: null });
  } catch (err) {
    alert(`Could not read the file.\n\n${err.message}`);
    return;
  }
  $("fileName").textContent = file.name;
  const list = $("clients");
  list.replaceChildren(...state.clients.map((c, i) => new Option(`${i + 1}. ${c.name}`, String(i))));
  $("start").disabled = $("end").disabled = false;
  clearPreview();
  log(`Found ${state.clients.length} client(s). Pick one from the list.`);
});

$("jsonfile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  try {
    const doc = core.validateExport(JSON.parse(await file.text()));
    Object.assign(state, { source: "json", fileName: file.name, sheet: null, clients: [], dateColumns: [], doc });
    $("fileName").textContent = `${file.name} (exported file)`;
    $("clients").replaceChildren(new Option(`1. ${doc.client}`, "0"));
    $("clients").selectedIndex = 0;
    $("start").value = $("end").value = "";
    $("start").disabled = $("end").disabled = true;
    state.result = {
      client: doc.client, rate: doc.rate,
      claims: doc.groups.flatMap((g) => g.claims),
      problems: doc.problems || [], warnings: doc.warnings || [],
    };
    state.groups = doc.groups.map((g) => g.claims);
    state.range = [null, null];
    renderPreview();
    log(`Loaded ${state.result.claims.length} claims for ${doc.client} from ${file.name}.`);
  } catch (err) {
    alert(`Could not read the file.\n\n${err.message}`);
  }
});

// --- preview --------------------------------------------------------------------------
function clearPreview() {
  state.result = null;
  state.groups = [];
  $("clientTitle").textContent = "No client selected";
  $("group").replaceChildren();
  $("group").disabled = true;
  $("claims").tBodies[0].replaceChildren();
  $("summary").textContent = "";
  $("problems").textContent = "";
  updateButtons();
}

function onClientOrRangeChanged() {
  if (state.source !== "xlsx") return;
  const index = $("clients").selectedIndex;
  if (index < 0) return;
  clearPreview();
  const client = state.clients[index];
  $("clientTitle").textContent = `${index + 1}. ${client.name}`;
  let start, end;
  try {
    start = core.parseDay($("start").value);
    end = core.parseDay($("end").value);
  } catch (err) {
    $("problems").textContent = err.message;
    return;
  }
  const iso = (d) => `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  if (start && end && iso(start) > iso(end)) {
    $("problems").textContent = "'From' is after 'To'.";
    return;
  }
  state.range = [start, end];
  state.result = core.readClient(state.sheet, client, state.dateColumns, start, end);
  state.groups = core.groupClaims(state.result.claims);
  renderPreview();
}
$("clients").addEventListener("change", onClientOrRangeChanged);
$("start").addEventListener("input", onClientOrRangeChanged);
$("end").addEventListener("input", onClientOrRangeChanged);

function renderPreview() {
  const { result, groups } = state;
  $("clientTitle").textContent = `${$("clients").selectedIndex + 1}. ${result.client}`;
  const rows = [];
  groups.forEach((group, n) => {
    for (const c of group) {
      const tr = document.createElement("tr");
      const cells = [n + 1, c.service_date, c.service, c.hours, c.total_cost === null ? "" : c.total_cost.toFixed(2)];
      cells.forEach((v, i) => {
        const td = document.createElement("td");
        td.textContent = String(v);
        if (i >= 3) td.className = "num";
        tr.append(td);
      });
      rows.push(tr);
    }
  });
  $("claims").tBodies[0].replaceChildren(...rows);
  const group = $("group");
  group.replaceChildren(...groups.map((_, i) => new Option(core.groupLabel(groups, i + 1), String(i))));
  group.disabled = groups.length < 2;
  const hours = result.claims.reduce((a, c) => a + c.hours, 0);
  const cost = result.claims.reduce((a, c) => a + (c.total_cost || 0), 0);
  $("summary").textContent = `${result.claims.length} claims, ${+hours.toFixed(2)} hours, $${cost.toFixed(2)}` +
    (result.rate ? `   (rate $${result.rate}/hr)` : "");
  $("problems").textContent = [...result.problems, ...result.warnings].join("\n");
  updateButtons();
}

function updateButtons() {
  const r = state.result;
  const ready = !!(r && r.claims.length && !r.problems.length);
  $("fill").disabled = !ready || state.running;
  $("export").disabled = !ready || state.running;
  $("stop").hidden = !state.running;
  $("open").disabled = state.running;
}

// --- website address, permissions, tabs --------------------------------------------------
async function loadUrl() {
  const { url } = await chrome.storage.local.get("url");
  $("url").value = url || DEFAULT_URL;
}

function normalizedUrl() {
  let url = $("url").value.trim();
  if (!url) throw new Error("Paste the claim website address first.");
  if (!url.includes("://")) url = "https://" + url;
  $("url").value = url;
  return new URL(url);
}

async function ensurePermission(url) {
  const origins = [`${url.origin}/*`];
  if (await chrome.permissions.contains({ origins })) return;
  if (!(await chrome.permissions.request({ origins }))) {
    throw new Error(`Permission to use ${url.origin} was not given, so the form cannot be filled.`);
  }
}

$("open").addEventListener("click", async () => {
  try {
    const url = normalizedUrl();
    await chrome.storage.local.set({ url: url.href });
    await ensurePermission(url); // must run first: it needs the click
    await chrome.tabs.create({ url: url.href });
    log(`Opened ${url.href}. Log in if needed and go to the Claim details form, then click 'Fill form'.`);
  } catch (err) {
    alert(err.message);
  }
});

/** The tab on the claim website that has the Claim details form open. */
async function findFormTab(url, options) {
  const tabs = await chrome.tabs.query({ url: `${url.origin}/*` });
  if (!tabs.length) throw new Error("The claim website is not open. Click 'Open claim website' first.");
  tabs.sort((a, b) => Number(b.active) - Number(a.active) || (b.lastAccessed || 0) - (a.lastAccessed || 0));
  let lastMissing = [];
  for (const tab of tabs) {
    await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["src/content.js"] });
    const reply = await roundTrip(tab.id, { type: "check", options });
    if (reply.ready) return tab;
    lastMissing = reply.missing;
  }
  throw new Error(`The Claim details form was not found (missing: ${lastMissing.join(", ")}). Open the form in the claim website tab and try again.`);
}

function roundTrip(tabId, message) {
  return new Promise((resolve, reject) => {
    const port = chrome.tabs.connect(tabId, { name: "claim-filler" });
    const timer = setTimeout(() => { port.disconnect(); reject(new Error("The page did not answer.")); }, 5000);
    port.onMessage.addListener((m) => { clearTimeout(timer); port.disconnect(); resolve(m); });
    port.postMessage(message);
  });
}

// --- fill -------------------------------------------------------------------------------
async function engineOptions() {
  const { options } = await chrome.storage.local.get("options");
  return options || {}; // optional wording overrides, see README
}

$("fill").addEventListener("click", async () => {
  let url;
  try {
    url = normalizedUrl();
    await ensurePermission(url); // first, while the click is fresh
  } catch (err) {
    alert(err.message);
    return;
  }
  const group = state.groups[$("group").selectedIndex] || [];
  const testOnly = $("firstOnly").checked;
  const claims = testOnly ? group.slice(0, 1) : group;
  const which = state.groups.length > 1 ? ` (Group ${$("group").selectedIndex + 1})` : "";
  const note = testOnly
    ? "This is a test, so 'Predetermine' will NOT be clicked."
    : "When every claim is added and the totals match, 'Predetermine' will be clicked.";
  if (!confirm(`Enter ${claims.length} claim(s) for ${state.result.client}${which}?\n\n${note}\n\nCheck that the browser tab is showing the Claim details form.`)) return;

  state.running = true;
  updateButtons();
  try {
    const options = await engineOptions();
    const tab = await findFormTab(url, options);
    await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["src/content.js"] });
    const outcome = await runFill(tab.id, claims, !testOnly, options);
    log(outcome.ok
      ? `Done for ${state.result.client}. Pick another group or client.`
      : "STOPPED. " + outcome.message);
  } catch (err) {
    log("Stopped: " + err.message);
  } finally {
    state.running = false;
    state.port = null;
    updateButtons();
  }
});

function runFill(tabId, claims, predetermine, options) {
  return new Promise((resolve, reject) => {
    const port = chrome.tabs.connect(tabId, { name: "claim-filler" });
    state.port = port;
    port.onMessage.addListener((m) => {
      if (m.type === "log") log(m.text);
      else if (m.type === "done") { port.disconnect(); resolve(m.outcome); }
    });
    port.onDisconnect.addListener(() => reject(new Error("The page was closed or reloaded while filling.")));
    port.postMessage({ type: "fill", claims, predetermine, options });
  });
}

$("stop").addEventListener("click", () => {
  if (state.port) { state.port.postMessage({ type: "stop" }); log("Stopping after the current claim…"); }
});

// --- export -----------------------------------------------------------------------------
$("export").addEventListener("click", () => {
  const [start, end] = state.range;
  const doc = state.source === "json"
    ? state.doc
    : core.buildExport(state.result, state.groups, start, end, state.fileName, state.sheet.title);
  const name = core.exportFilename(state.result.client, start, end);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([JSON.stringify(doc, null, 2)], { type: "application/json" }));
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
  log(`Saved ${name} (${state.result.claims.length} claims in ${state.groups.length} group(s)).`);
});

$("version").textContent = chrome.runtime.getManifest().version;
loadUrl();
updateButtons();
log("Start by choosing the Excel file.");
