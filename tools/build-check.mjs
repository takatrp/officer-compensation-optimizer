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
assert.ok(html.includes("v1.3"));
assert.ok(html.includes("strategyPreset"));
assert.ok(html.includes("バランス型"));
assert.ok(html.includes("法人資金温存"));
assert.ok(html.includes("data-for=\"preProfit\""));
assert.ok(html.includes("data-for=\"currentMonthly\""));
assert.ok(html.includes("令和8・9年分"));
assert.ok(html.includes("令和10年分以後"));
assert.ok(html.includes("国民健康保険料・国民年金保険料"));
assert.ok(html.includes("CPI連動見直し"));
assert.ok(html.includes("令和8年度料率"));
assert.ok(js.includes("renderBalance"));
assert.ok(js.includes("renderRecommendationReason"));
assert.ok(js.includes("strategyPresets"));
assert.ok(js.includes("applyStrategyPreset"));
assert.ok(js.includes("updateDependentControls"));
assert.ok(js.includes("formatMoneyLive"));
assert.ok(js.includes("createCalculator"));
assert.ok(js.includes("1040000"));
assert.ok(js.includes("990000"));
assert.ok(css.includes(".balance-track"));
assert.ok(css.includes(".recommendation-reason"));
assert.ok(css.includes(".reason-list"));
assert.doesNotMatch(html, /<script[^>]+src=["']https?:/i);
assert.doesNotMatch(html, /<link[^>]+href=["']https?:/i);

console.log("Build check passed.");
