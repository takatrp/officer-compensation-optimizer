import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const html = readFileSync(resolve(root, "index.html"), "utf8");
const css = readFileSync(resolve(root, "styles.css"), "utf8");
const js = readFileSync(resolve(root, "app.js"), "utf8");

for (const localAsset of ["styles.css", "app.js"]) {
  assert.ok(existsSync(resolve(root, localAsset)), `${localAsset} exists`);
  assert.ok(html.includes(localAsset), `${localAsset} is referenced`);
}

assert.ok(html.includes("役員報酬 最適化シミュレーター"));
assert.ok(html.includes("data-for=\"preProfit\""));
assert.ok(html.includes("data-for=\"currentMonthly\""));
assert.ok(html.includes("令和8年度料率"));
assert.ok(js.includes("renderBalance"));
assert.ok(js.includes("updateDependentControls"));
assert.ok(js.includes("formatMoneyLive"));
assert.ok(js.includes("createCalculator"));
assert.ok(css.includes(".balance-track"));
assert.doesNotMatch(html, /<script[^>]+src=["']https?:/i);
assert.doesNotMatch(html, /<link[^>]+href=["']https?:/i);

console.log("Build check passed.");
