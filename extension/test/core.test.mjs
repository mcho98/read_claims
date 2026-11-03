// Checks the JavaScript reader gives the same answer as the Python script.
// shared/out.json is what python-app/read_claims.py produced from shared/mock_data.xlsx.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";
import * as core from "../src/core.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const XLSX = createRequire(import.meta.url)(path.join(here, "../vendor/xlsx.full.min.js"));
const shared = (f) => path.join(here, "../../shared", f);
const bufferOf = (f) => {
  const b = readFileSync(shared(f));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
};

function load() {
  const sheet = core.loadSheet(XLSX, bufferOf("mock_data.xlsx"));
  const clients = core.findClients(sheet);
  const { columns } = core.findDateColumns(sheet);
  return { sheet, clients, columns };
}

test("full month for the first client matches out.json", () => {
  const expected = JSON.parse(readFileSync(shared("out.json"), "utf8"));
  const { sheet, clients, columns } = load();
  assert.equal(clients.length, 3);
  const result = core.readClient(sheet, clients[0], columns);
  const doc = core.buildExport(result, core.groupClaims(result.claims), null, null, "mock_data.xlsx", sheet.title);
  for (const d of [doc, expected]) delete d.generated_at;
  assert.deepEqual(doc, expected);
});

test("date range keeps only the days inside it", () => {
  const { sheet, clients, columns } = load();
  const all = core.readClient(sheet, clients[0], columns).claims;
  const days = [...new Set(all.map((c) => c.service_date))].sort();
  assert.ok(days.length >= 4, "the fixture needs at least 4 different days");
  const [from, to] = [days[1], days[days.length - 2]];
  const r = core.readClient(sheet, clients[0], columns, core.parseDay(from.replace(/-/g, "")), core.parseDay(to.replace(/-/g, "")));
  const want = all.filter((c) => c.service_date >= from && c.service_date <= to);
  assert.ok(want.length > 0 && want.length < all.length);
  assert.deepEqual(r.claims, want);
});

test("groups are fixed day-of-month ranges, empty ones are left out", () => {
  const claim = (day) => ({ service_date: `2026-10-${String(day).padStart(2, "0")}` });
  const days = [1, 2, 4, 10, 11, 15, 16, 17, 25, 26, 31]; // day 3 has no claim
  const groups = core.groupClaims(days.map(claim));
  assert.deepEqual(groups.map((g) => [g.number, g.days, g.claims.length]),
    [[1, "1-10", 4], [2, "11-15", 2], [3, "16-25", 3], [4, "26-end", 2]]);
  // a custom range of days 7-17 leaves three groups
  const picked = core.groupClaims([7, 9, 12, 14, 16, 17].map(claim));
  assert.deepEqual(picked.map((g) => g.number), [1, 2, 3]);
  assert.equal(core.groupLabel(groups, 2), "Group 2 (days 11-15): 2 claims");
});

test("a day range with more than 10 claims is split into parts", () => {
  const groups = core.groupClaims(Array.from({ length: 12 }, () => ({ service_date: "2026-10-03" })));
  assert.deepEqual(groups.map((g) => [g.number, g.part, g.parts, g.claims.length]), [[1, 1, 2, 10], [1, 2, 2, 2]]);
  assert.equal(core.groupTitle(groups[1]), "Group 1 (days 1-10) - part 2 of 2");
});

test("money rounds half up like Python", () => {
  assert.equal(core.toMoney(3.5 * 34.03), 119.11);
  assert.equal(core.toMoney(1.53 * 34.03), 52.07);
  assert.equal(core.toMoney(4 * 34.03), 136.12);
});

test("dates: YYYYMMDD accepted, short or impossible rejected", () => {
  assert.deepEqual(core.parseDay("20260815"), { y: 2026, m: 8, d: 15 });
  assert.equal(core.parseDay(""), null);
  assert.throws(() => core.parseDay("2026081"));
  assert.throws(() => core.parseDay("20260231"));
});

test("export file name", () => {
  assert.equal(core.exportFilename("John Doe", null, null), "John_Doe.json");
  assert.equal(core.exportFilename("John Doe", { y: 2026, m: 8, d: 10 }, null), "John_Doe_20260810-end.json");
});

test("site scope widens login/provider sub-domains but not shared hosts", () => {
  assert.equal(core.siteScope("https://login.ab.bluecross.ca/iamapp/login?x=1").pattern, "https://*.ab.bluecross.ca/*");
  assert.equal(core.siteScope("https://provider.ab.bluecross.ca/health/x.jspx").pattern, "https://*.ab.bluecross.ca/*");
  assert.equal(core.siteScope("https://mock-website-theta.vercel.app/").pattern, "https://mock-website-theta.vercel.app/*");
  assert.equal(core.siteScope("http://127.0.0.1:8765/").pattern, "http://127.0.0.1:8765/*");
  assert.equal(core.siteScope("http://localhost:3000").pattern, "http://localhost:3000/*");
});

test("client menu text is 'family - client number - name' from the mock workbook", () => {
  const { clients } = load();
  assert.deepEqual(clients.map((c) => c.name), ["1 - 1 - John Doe", "4 - 5 - Albert Albertson", "5 - 6 - Bob Barker"]);
  assert.deepEqual(clients.map((c) => [c.familyNumber, c.clientNumber]), [[1, 1], [4, 5], [5, 6]]);
});

/** A tiny fake sheet: cells given as {"B3": 1, ...}. */
function fakeSheet(cells) {
  const letters = (s) => s.charCodeAt(0) - 64;
  const map = new Map(Object.entries(cells).map(([k, v]) => [`${k.match(/\d+/)[0]}:${letters(k[0])}`, v]));
  return { maxRow: 12, maxCol: 6, get: (r, c) => map.get(`${r}:${c}`) };
}

test("name is the cell under the AHS ID when it has text; otherwise the Client Name cell without its leading number", () => {
  const base = { A2: "Family Number", B2: "Client Number", A3: 4, B3: 5, C3: 234234, B5: "Client Name", C5: "4 Albert Albertson" };
  assert.equal(core.findClients(fakeSheet({ ...base, C4: "Albert A. Albertson" }))[0].name, "4 - 5 - Albert A. Albertson");
  for (const under of [undefined, "", "   "]) {
    assert.equal(core.findClients(fakeSheet({ ...base, C4: under }))[0].name, "4 - 5 - Albert Albertson");
  }
});

test("sheets without the number headers show the plain name", () => {
  const sheet = fakeSheet({ B5: "Client Name", C5: "3 Jane Roe" });
  assert.equal(core.findClients(sheet)[0].name, "Jane Roe");
});

test("money product is exact and rounds half up", () => {
  assert.equal(core.moneyProduct(3.5, 35.03), 122.61); // 122.605 exactly; a float product would give 122.60
  assert.equal(core.moneyProduct(1.23, 35.03), 43.09);
  assert.equal(core.moneyProduct(4, 34.03), 136.12);
  assert.equal(core.moneyProduct(0.01, 0.5), 0.01);
});

test("hours with many decimals are rounded to 2 before display, cost and form-fill", () => {
  // E column = 2026-10-01..03; one service row with long decimals, a tiny value and a normal one
  const cells = { B5: "Client Name", C5: "Test", C7: "rate $/hr", D7: 35.03, B9: "Personal Care hours", E9: 1.23456, F9: 0.004, G9: 3.335 };
  const sheet = {
    maxRow: 12, maxCol: 8,
    get: (r, c) => cells[String.fromCharCode(64 + c) + r],
    coordinate: (r, c) => String.fromCharCode(64 + c) + r,
  };
  const days = [1, 2, 3].map((d, i) => ({ col: 5 + i, day: { y: 2026, m: 10, d } }));
  const r = core.readClient(sheet, { name: "Test", firstRow: 5, lastRow: 12 }, days);
  assert.deepEqual(r.claims.map((c) => [c.service_date, c.hours, c.total_cost]),
    [["2026-10-01", 1.23, 43.09], ["2026-10-03", 3.34, 117.0]]); // 0.004 rounds to 0 and is skipped; 3.34 x 35.03 = 117.0002
  assert.deepEqual(r.problems, []);
});
