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
  const t = new Date().toLocaleTimeString([], { hour12: false });
  box.textContent += `[${t}] ${text}\n`;
  box.scrollTop = box.scrollHeight;
}

// Log lines sent from the claim page (it may reload at any moment, so it reports by message).
chrome.runtime.onMessage.addListener((m) => {
  if (m && m.type === "claim-filler-log") log(m.text);
});

$("copyLog").addEventListener("click", async () => {
  await navigator.clipboard.writeText($("log").textContent);
  log("Log copied to the clipboard.");
});
$("saveLog").addEventListener("click", () => {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([$("log").textContent], { type: "text/plain" }));
  a.download = `claim-filler-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`;
  a.click();
  URL.revokeObjectURL(a.href);
});

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
  list.replaceChildren(...state.clients.map((c, i) => new Option(c.name, String(i))));
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
    $("clients").replaceChildren(new Option(doc.client, "0"));
    $("clients").selectedIndex = 0;
    $("start").value = $("end").value = "";
    $("start").disabled = $("end").disabled = true;
    state.result = {
      client: doc.client, clientName: doc.client_name ?? doc.client,
      familyNumber: doc.family_number ?? null, clientNumber: doc.client_number ?? null, rate: doc.rate,
      claims: doc.groups.flatMap((g) => g.claims),
      problems: doc.problems || [], warnings: doc.warnings || [],
    };
    state.groups = core.groupsFromExport(doc);
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
  $("clientTitle").textContent = client.name;
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
  $("clientTitle").textContent = result.client;
  const group = $("group");
  group.replaceChildren(...groups.map((_, i) => new Option(core.groupLabel(groups, i + 1), String(i))));
  group.disabled = groups.length < 2;
  $("problems").textContent = [...result.problems, ...result.warnings].join("\n");
  renderGroup();
}

/** Table and summary for the selected group only (the website takes one group at a time). */
function renderGroup() {
  const { result, groups } = state;
  const index = Math.max(0, $("group").selectedIndex);
  const info = groups[index];
  const claims = info ? info.claims : [];
  const rows = claims.map((c) => {
    const tr = document.createElement("tr");
    [info.number, c.service_date, c.service, c.hours, c.total_cost === null ? "" : c.total_cost.toFixed(2)].forEach((v, i) => {
      const td = document.createElement("td");
      td.textContent = String(v);
      if (i >= 3) td.className = "num";
      tr.append(td);
    });
    return tr;
  });
  $("claims").tBodies[0].replaceChildren(...rows);
  const hours = claims.reduce((a, c) => a + c.hours, 0);
  const cost = claims.reduce((a, c) => a + (c.total_cost || 0), 0);
  const where = info ? `${core.groupTitle(info)}: ` : "";
  $("summary").textContent = `${where}${claims.length} claims, ${+hours.toFixed(2)} hours, $${cost.toFixed(2)}` +
    (result.rate ? `   (rate $${result.rate}/hr)` : "");
  updateButtons();
}
$("group").addEventListener("change", renderGroup);

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
  const { pattern, label } = core.siteScope(url.href);
  const origins = [pattern];
  if (await chrome.permissions.contains({ origins })) return;
  if (!(await chrome.permissions.request({ origins }))) {
    throw new Error(`Permission to use ${label} was not given, so the form cannot be filled.`);
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

// --- talking to the claim page ----------------------------------------------------------
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const withTimeout = (promise, ms, what) => Promise.race([
  promise,
  new Promise((_, reject) => setTimeout(() => reject(new Error(`${what} did not finish within ${Math.round(ms / 1000)} s`)), ms)),
]);

/** Wait for the tab to finish loading (the site reloads the page after every Add claim). */
async function waitForLoaded(tabId, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const tab = await chrome.tabs.get(tabId);
    if (tab.status === "complete") return;
    await sleep(200);
  }
  throw new Error("The page did not finish loading.");
}

async function inject(tabId) {
  await chrome.scripting.executeScript({ target: { tabId }, files: ["src/content.js"] });
}

/** Run one operation of the engine in the page. Resolves with its result; rejects if the page reloads mid-way. */
async function callEngine(tabId, op, payload, options, timeoutMs = 60000) {
  await inject(tabId);
  const [res] = await withTimeout(chrome.scripting.executeScript({
    target: { tabId },
    func: (o, p, opt) => globalThis.ClaimFillerEngine.run(o, p, opt),
    args: [op, payload ?? null, options],
  }), timeoutMs, `'${op}' in the page`);
  return res ? res.result : undefined;
}

/** Ask the page for its state, riding out a reload in progress. */
async function readPageState(tabId, options, maxWaitMs = 15000) {
  const deadline = Date.now() + maxWaitMs;
  let lastError;
  while (Date.now() < deadline) {
    try {
      await waitForLoaded(tabId, 10000);
      return await callEngine(tabId, "state", null, options, 10000);
    } catch (e) {
      lastError = e;
      await sleep(300);
    }
  }
  throw lastError || new Error("Could not read the page.");
}

/** The tab on the claim website that has the Claim details form open. */
async function findFormTab(url, options) {
  const { pattern, label } = core.siteScope(url.href);
  const tabs = await chrome.tabs.query({ url: pattern });
  if (!tabs.length) {
    throw new Error(`No open tab on ${label}. Click 'Open claim website', log in and open the Claim details form in that tab first.`);
  }
  tabs.sort((a, b) => Number(b.active) - Number(a.active) || (b.lastAccessed || 0) - (a.lastAccessed || 0));
  log(`Found ${tabs.length} tab(s) on ${label}.`);
  let lastMissing = [];
  for (const tab of tabs) {
    log(`Checking tab ${tab.id} (${tab.title || tab.url}) for the Claim details form…`);
    try {
      const st = await readPageState(tab.id, options, 8000);
      if (st.ready) {
        log(`Form found. Website total cost: ${st.total === null ? "none shown" : "$" + (st.total / 100).toFixed(2)}. Form: ${JSON.stringify(st.form)}`);
        return tab;
      }
      lastMissing = st.missing;
      log(`  not the form (missing: ${st.missing.join("; ")})`);
    } catch (e) {
      log(`  could not check this tab: ${e.message}`);
    }
  }
  throw new Error(`The Claim details form was not found (${lastMissing.join("; ") || "no usable tab"}). Open the form in the claim website tab and try again.`);
}

// --- fill -------------------------------------------------------------------------------
const MAX_ATTEMPTS = 3;
const TOTALS_WAIT_MS = 15000; // time for the site to answer Add claim (and reload the page)
const dollars = (c) => (c / 100).toFixed(2);

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
  const info = state.groups[$("group").selectedIndex];
  const group = info ? info.claims : [];
  const testOnly = $("firstOnly").checked;
  const claims = testOnly ? group.slice(0, 1) : group;
  const which = info ? ` (${core.groupTitle(info)})` : "";
  const note = testOnly
    ? "This is a test, so 'Predetermine' will NOT be clicked."
    : "When every claim is added and the totals match, 'Predetermine' will be clicked.";
  if (!confirm(`Enter ${claims.length} claim(s) for ${state.result.client}${which}?\n\n${note}\n\nCheck that the browser tab is showing the Claim details form.`)) return;

  state.running = true;
  state.stopRequested = false;
  updateButtons();
  log(`--- Filling ${claims.length} claim(s) for ${state.result.client}${which}${testOnly ? " (test: first claim only)" : ""} ---`);
  try {
    const options = await engineOptions();
    const tab = await findFormTab(url, options);
    const outcome = await runClaims(tab.id, claims, !testOnly, options);
    log(outcome.ok ? `Done for ${state.result.client}. Pick another group or client.` : "STOPPED. " + outcome.message);
  } catch (err) {
    log("Stopped: " + err.message);
  } finally {
    state.running = false;
    updateButtons();
  }
});

$("stop").addEventListener("click", () => {
  state.stopRequested = true;
  log("Stop requested: finishing the current step first…");
});

/**
 * The claim loop. It lives here (not in the page) because the site reloads the page after
 * every Add claim, which destroys the page script. For each claim: fill + click in the page,
 * then wait for the reload and read the site's running total cost; the claim counts as added
 * only when the total equals the rolling sum. Up to MAX_ATTEMPTS tries per claim; a third
 * failure (or an unexpected total) stops everything.
 */
async function runClaims(tabId, claims, predetermine, options) {
  let st = await readPageState(tabId, options);
  let expected = st.total ?? 0;
  log(`Website total cost before adding: $${dollars(expected)}`);
  for (let i = 0; i < claims.length; i++) {
    const c = claims[i];
    const n = i + 1;
    const label = `${c.service_date} ${c.service} ${c.hours}h`;
    const before = expected;
    expected += Math.round((c.total_cost || 0) * 100);
    let added = false;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS && !added; attempt++) {
      if (state.stopRequested) return { ok: false, added: i, message: `Stopped by you after ${i} of ${claims.length} claims.` };
      log(`Claim ${n}/${claims.length}: ${label}, attempt ${attempt}/${MAX_ATTEMPTS}`);
      let result;
      try {
        result = await callEngine(tabId, "enter", c, options, 90000);
      } catch (e) {
        // Expected on the real site: the page reloaded right after Add claim and cut the call off.
        log(`  the page reloaded or the call ended (${e.message.split("\n")[0]}); checking the total…`);
        result = { status: "unknown" };
      }
      if (result && result.status === "failed") {
        log(`  could not enter the claim: ${result.reason}`);
        if (result.errors && result.errors.length) log(`  the site says: ${result.errors.join(" | ")}`);
        if (result.form) log(`  ${result.form}`);
      }
      // Wait for the site to answer (and reload) and for its total to reach the expected value.
      let seen = null;
      const deadline = Date.now() + (result && result.status === "failed" ? 2000 : TOTALS_WAIT_MS);
      while (Date.now() < deadline) {
        await sleep(400);
        try {
          await waitForLoaded(tabId, 10000);
          const s = await callEngine(tabId, "state", null, options, 10000);
          seen = s.total;
          if (seen === expected) break;
        } catch (e) { /* mid-reload: try again */ }
      }
      if (seen === expected) {
        log(`  added ${n}/${claims.length}: ${label} (website total now $${dollars(seen)})`);
        added = true;
      } else if (seen !== null && seen !== before) {
        const message = `Stopped at claim ${n} (${label}): the website total is $${dollars(seen)} but $${dollars(expected)} was expected. Check the website before continuing.`;
        log(message);
        return { ok: false, added: i, message };
      } else {
        let extra = "";
        try {
          const s = await readPageState(tabId, options, 5000);
          if (s.errors && s.errors.length) extra += ` The site says: ${s.errors.join(" | ")}.`;
          if (s.form) extra += ` Form now holds: ${JSON.stringify(s.form)}.`;
        } catch (e) { /* ignore */ }
        log(`  attempt ${attempt}/${MAX_ATTEMPTS} failed: the website total did not change (still ${seen === null ? "none shown" : "$" + dollars(seen)}, wanted $${dollars(expected)}).${extra}`);
      }
    }
    if (!added) {
      const message = `Stopped: claim ${n} (${label}) failed ${MAX_ATTEMPTS} times. ${i} of ${claims.length} claims were added.`;
      log(message);
      return { ok: false, added: i, message };
    }
  }
  let message = `All ${claims.length} claims added; total cost matches ($${dollars(expected)}).`;
  if (predetermine) {
    let r;
    try {
      r = await callEngine(tabId, "predetermine", null, options, 20000);
    } catch (e) {
      r = { clicked: true }; // the click started a page change that cut the call off
    }
    if (!r.clicked) {
      message += ` But ${r.reason}.`;
      log(message);
      return { ok: false, added: claims.length, message };
    }
    message += " Clicked 'Predetermine'.";
  }
  log(message);
  return { ok: true, added: claims.length, message };
}

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
