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
  assert.equal(clients.length, 1);
  const result = core.readClient(sheet, clients[0], columns);
  const doc = core.buildExport(result, core.groupClaims(result.claims), null, null, "mock_data.xlsx", sheet.title);
  for (const d of [doc, expected]) delete d.generated_at;
  assert.deepEqual(doc, expected);
});

test("date range keeps only the days inside it", () => {
  const { sheet, clients, columns } = load();
  const r = core.readClient(sheet, clients[0], columns, core.parseDay("20260810"), core.parseDay("20260815"));
  assert.equal(r.claims.length, 6);
  assert.equal(core.buildExport(r, core.groupClaims(r.claims), null, null, "x", "y").totals.total_cost, 408.36);
});

test("grouping 15 gives 10 + 5", () => {
  const groups = core.groupClaims(Array.from({ length: 15 }, (_, i) => i));
  assert.deepEqual(groups.map((g) => g.length), [10, 5]);
  assert.equal(core.groupLabel(groups, 2), "Group 2 of 2 (claims 11-15)");
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
