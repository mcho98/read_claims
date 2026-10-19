// Reads the timesheet workbook and turns one client into claims.
// A port of python-app/read_claims.py (find_date_columns, find_clients, read_client,
// group_claims, build_export) so both versions give the same result for the same file.
// Pure logic: no DOM and no chrome.* calls, so it runs in the side panel and in Node tests.
// The SheetJS library is passed in (XLSX) instead of imported, because it is a plain script.

export const SERVICES = {
  "homemaking hours": "Homemaking",
  "personal care hours": "Personal Care",
  "respite hours": "Respite",
};
export const CLIENT_LABEL = "client name";
export const RATE_LABEL = "rate $/hr";
export const MAX_HOURS_PER_DAY = 24;
export const MAX_CLAIMS_PER_GROUP = 10; // the website accepts at most this many claims at a time
export const EXPORT_FORMAT = "claimfiller-claims";
export const EXPORT_VERSION = 2; // 2: groups are fixed day-of-month ranges and carry days/part/parts

const pad = (n, w = 2) => String(n).padStart(w, "0");
const iso = (d) => `${pad(d.y, 4)}-${pad(d.m)}-${pad(d.d)}`;

function addDays(d, n) {
  const t = new Date(Date.UTC(d.y, d.m - 1, d.d + n));
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

/** "A".."Z","AA".. for a 0-based column index; same as Excel's letters. */
function colLetters(c) {
  let s = "";
  for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

const norm = (v) => (typeof v === "string" ? v.trim().toLowerCase() : null);

/**
 * Parse an .xlsx file (ArrayBuffer) and wrap the sheet in a small helper object.
 * cellDates stays off: dates are read from the serial numbers, which avoids the
 * time-zone shifts SheetJS can introduce when it builds JavaScript Dates.
 */
export function loadSheet(XLSX, arrayBuffer, sheetName) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellNF: true });
  const name = sheetName || wb.SheetNames[0];
  const sheet = wb.Sheets[name];
  if (!sheet) throw new Error(`Sheet '${name}' was not found.`);
  const ref = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const isDate = (cell) => cell.t === "n" && cell.z && XLSX.SSF.is_date(cell.z);
  return {
    title: name,
    maxRow: ref.e.r + 1, // 1-based like openpyxl
    maxCol: ref.e.c + 1,
    /** value at 1-based (row, col): number, string, boolean, {date:{y,m,d}} or undefined */
    get(row, col) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
      if (!cell || cell.v === undefined || cell.v === null) return undefined;
      if (isDate(cell)) {
        const p = XLSX.SSF.parse_date_code(cell.v);
        if (p) return { date: { y: p.y, m: p.m, d: p.d } };
      }
      return cell.v;
    },
    coordinate: (row, col) => `${colLetters(col - 1)}${row}`,
  };
}

const asDate = (v) => (v && typeof v === "object" && v.date ? v.date : null);

/** [{col, day:{y,m,d}}] for each day of the month in the date row (first date cell in the first 20 rows). */
export function findDateColumns(sheet) {
  for (let row = 1; row <= Math.min(sheet.maxRow, 20); row++) {
    for (let col = 1; col <= sheet.maxCol; col++) {
      const start = asDate(sheet.get(row, col));
      if (!start) continue;
      // Following cells are formulas; use their cached value when present, else count up a day per column.
      const columns = [];
      let day = start;
      let c = col;
      while (day.m === start.m) {
        columns.push({ col: c, day: asDate(sheet.get(row, c)) || day });
        c += 1;
        day = addDays(day, 1);
      }
      return { row, columns };
    }
  }
  throw new Error("Could not find the date row (no date cell in the first 20 rows).");
}

/** [{name, firstRow, lastRow}] one per "Client Name" label. */
export function findClients(sheet) {
  const starts = [];
  for (let row = 1; row <= sheet.maxRow; row++) {
    for (let col = 1; col <= sheet.maxCol; col++) {
      if (norm(sheet.get(row, col)) === CLIENT_LABEL) {
        const name = sheet.get(row, col + 1);
        starts.push({ row, name: name ? String(name).trim() : `(unnamed, row ${row})` });
      }
    }
  }
  return starts.map((s, i) => ({
    name: s.name,
    firstRow: s.row,
    lastRow: i + 1 < starts.length ? starts[i + 1].row - 1 : sheet.maxRow,
  }));
}

/** Round half up to cents using the shortest decimal text of the number (matches Python's Decimal(str(x))). */
export function toMoney(value) {
  const m = /^(-?)(\d+)(?:\.(\d+))?$/.exec(String(value));
  if (!m) return Math.round(value * 100) / 100; // exponent form etc.: not expected for money
  const [, sign, whole, frac = ""] = m;
  const padded = (frac + "00").slice(0, 2);
  let cents = BigInt(whole + padded);
  if (frac.length > 2 && frac[2] >= "5") cents += 1n;
  const text = cents.toString().padStart(3, "0");
  return Number(`${sign}${text.slice(0, -2)}.${text.slice(-2)}`);
}

/** "20260815" or "2026-08-15" -> {y,m,d}; blank -> null; anything else throws. */
export function parseDay(text) {
  const t = (text || "").trim();
  if (!t) return null;
  const digits = t.replace(/-/g, "");
  const bad = () => new Error(`'${t}' is not a valid date. Use YYYYMMDD, e.g. 20260815.`);
  if (!/^\d{8}$/.test(digits)) throw bad();
  const d = { y: +digits.slice(0, 4), m: +digits.slice(4, 6), d: +digits.slice(6, 8) };
  const check = new Date(Date.UTC(d.y, d.m - 1, d.d));
  if (check.getUTCFullYear() !== d.y || check.getUTCMonth() + 1 !== d.m || check.getUTCDate() !== d.d) throw bad();
  return d;
}

const cmpDay = (a, b) => (iso(a) < iso(b) ? -1 : iso(a) > iso(b) ? 1 : 0);

/** Read one client. start/end are optional {y,m,d}; both ends inclusive. */
export function readClient(sheet, client, dateColumns, start = null, end = null) {
  let rate;
  const serviceRows = new Map();
  for (let row = client.firstRow; row <= client.lastRow; row++) {
    for (let col = 1; col <= sheet.maxCol; col++) {
      const label = norm(sheet.get(row, col));
      if (label === RATE_LABEL && rate === undefined) rate = sheet.get(row, col + 1);
      else if (label !== null && Object.hasOwn(SERVICES, label) && !serviceRows.has(row)) serviceRows.set(row, SERVICES[label]);
    }
  }
  const problems = [];
  const warnings = [];
  if (typeof rate !== "number" || !(rate > 0)) {
    problems.push(`No valid hourly rate next to '${RATE_LABEL}' (found ${rate === undefined ? "None" : JSON.stringify(rate)}).`);
    rate = null;
  }
  if (serviceRows.size === 0) problems.push("No service rows found (Homemaking / Personal Care / Respite hours).");

  const claims = [];
  for (const row of [...serviceRows.keys()].sort((a, b) => a - b)) {
    for (const { col, day } of dateColumns) {
      if ((start && cmpDay(day, start) < 0) || (end && cmpDay(day, end) > 0)) continue;
      const hours = sheet.get(row, col);
      if (hours === undefined || hours === "" || hours === 0) continue;
      const where = sheet.coordinate(row, col);
      if (typeof hours !== "number") {
        problems.push(`${where}: hours is not a number (${JSON.stringify(hours)}).`);
        continue;
      }
      if (hours < 0 || hours > MAX_HOURS_PER_DAY) {
        problems.push(`${where}: ${hours} hours is outside 0-${MAX_HOURS_PER_DAY}.`);
        continue;
      }
      claims.push({
        service_date: iso(day),
        service: serviceRows.get(row),
        hours,
        total_cost: rate ? toMoney(hours * rate) : null,
        cell: where,
      });
    }
  }
  const byText = (a, b) => (a < b ? -1 : a > b ? 1 : 0); // code-point order, like Python
  claims.sort((a, b) => byText(a.service_date, b.service_date) || byText(a.service, b.service));
  return { client: client.name, rate, claims, problems, warnings };
}

// Groups are fixed by day of the month, not by how many claims there are. Days with no claims
// are simply absent, so a group often holds fewer than 10 claims.
export const GROUP_DAY_RANGES = [[1, 10], [11, 15], [16, 25], [26, 31]];

/**
 * Group claims by day of the month: days 1-10, 11-15, 16-25 and 26-end of month. Only groups that
 * have claims are returned, each as {number: 1-4, days: "1-10", part, parts, claims}. The website
 * takes at most `size` claims at a time, so a day range holding more than that (several services
 * per day) is split into consecutive parts rather than overflowing.
 */
export function groupClaims(claims, size = MAX_CLAIMS_PER_GROUP) {
  const groups = [];
  GROUP_DAY_RANGES.forEach(([low, high], i) => {
    const inside = claims.filter((c) => {
      const day = Number(c.service_date.slice(8, 10));
      return day >= low && day <= high;
    });
    const parts = Math.ceil(inside.length / size);
    for (let p = 0; p < parts; p++) {
      groups.push({
        number: i + 1,
        days: high === 31 ? `${low}-end` : `${low}-${high}`,
        part: p + 1,
        parts,
        claims: inside.slice(p * size, (p + 1) * size),
      });
    }
  });
  return groups;
}

/** "Group 2 (days 11-15)", with " - part 1 of 2" when a day range had to be split. */
export function groupTitle(group) {
  return `Group ${group.number} (days ${group.days})` + (group.parts > 1 ? ` - part ${group.part} of ${group.parts}` : "");
}

/** Text for group n (1-based) in a list, e.g. "Group 2 (days 11-15): 5 claims". */
export function groupLabel(groups, n) {
  const g = groups[n - 1];
  return `${groupTitle(g)}: ${g.claims.length} claims`;
}

/** The groups of an exported file, in the same shape groupClaims() returns. */
export function groupsFromExport(doc) {
  return doc.groups.map((g, i) => ({
    number: g.group ?? i + 1, days: g.days ?? "", part: g.part ?? 1, parts: g.parts ?? 1, claims: g.claims,
  }));
}

/** The same document python-app writes with --json (format claimfiller-claims, version 1). */
export function buildExport(result, groups, start, end, sourceFile, sheetTitle, now = new Date()) {
  const sum = (f) => result.claims.reduce((a, c) => a + f(c), 0);
  return {
    format: EXPORT_FORMAT,
    version: EXPORT_VERSION,
    generated_at: now.toISOString(),
    source_file: sourceFile,
    sheet: sheetTitle,
    client: result.client,
    rate: result.rate,
    date_range: { start: start ? iso(start) : null, end: end ? iso(end) : null },
    max_claims_per_group: MAX_CLAIMS_PER_GROUP,
    totals: {
      claims: result.claims.length,
      hours: Math.round(sum((c) => c.hours) * 100) / 100,
      total_cost: toMoney(sum((c) => c.total_cost || 0)),
    },
    problems: result.problems,
    warnings: result.warnings,
    groups: groups.map((g) => ({
      group: g.number,
      days: g.days,
      part: g.part,
      parts: g.parts,
      claims: g.claims.map((c) => ({ service_date: c.service_date, service: c.service, hours: c.hours, total_cost: c.total_cost, cell: c.cell })),
    })),
  };
}

/** Suggested file name, e.g. John_Doe_20260810-20260815.json */
export function exportFilename(client, start, end) {
  let name = client.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "client";
  if (start || end) {
    const compact = (d) => iso(d).replace(/-/g, "");
    name += `_${start ? compact(start) : "start"}-${end ? compact(end) : "end"}`;
  }
  return `${name}.json`;
}

/** Check a parsed JSON file is a claimfiller export we understand; returns it or throws a readable error. */
export function validateExport(doc) {
  if (!doc || doc.format !== EXPORT_FORMAT) throw new Error("This is not a Claim Filler export file.");
  if (doc.version !== EXPORT_VERSION) {
    throw new Error(`This file is version ${doc.version}; this extension understands version ${EXPORT_VERSION}.`);
  }
  if (!Array.isArray(doc.groups)) throw new Error("The file has no claim groups.");
  return doc;
}

/**
 * Which sites the extension may use for an address. The login page and the claim form are often
 * on different sub-domains (login.ab.bluecross.ca vs provider.ab.bluecross.ca), so a host with
 * 4 or more labels is widened to its parent domain (*.ab.bluecross.ca). Shorter hosts, IP
 * addresses and localhost stay exact, so a shared host such as xyz.vercel.app is not widened.
 * Returns {pattern, label}; pattern is a Chrome match pattern.
 */
export function siteScope(address) {
  const u = new URL(address);
  const labels = u.hostname.split(".");
  const isIp = /^\d+(\.\d+){3}$/.test(u.hostname) || u.hostname.includes(":");
  if (!isIp && labels.length >= 4) {
    const parent = labels.slice(1).join(".");
    return { pattern: `${u.protocol}//*.${parent}/*`, label: `*.${parent}` };
  }
  return { pattern: `${u.origin}/*`, label: u.host };
}
