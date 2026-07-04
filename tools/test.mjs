import assert from "node:assert/strict";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const app = require("../app.js");

assert.equal(app.parseNumber("20,000,000"), 20000000);
assert.equal(app.formatMoneyText("3000000"), "3,000,000");
assert.equal(app.formatMoneyText("3,000,000"), "3,000,000");
assert.equal(app.formatMoneyText(""), "");
assert.equal(app.formatInput("shareRate", 100), "100");
assert.equal(app.formatInput("fixedPayout", 50), "50");
assert.equal(app.formatInput("healthRate", 10.12), "10.12");
assert.equal(app.salaryDeduction(1900000), 650000);
assert.equal(app.salaryDeduction(3600000), 1160000);
assert.equal(app.basicNational(1320000, "r8r9"), 950000);
assert.equal(app.basicNational(1320001, "r8r9"), 580000);
assert.equal(app.basicNational(1320001, "r7"), 880000);
assert.equal(app.nationalTaxBeforeCredits(7000000), 974000);

const calculator = app.createCalculator();
const params = app.defaultParams();
const sample = calculator.simulate(800000, 0, params);

assert.equal(sample.feasible, true);
assert.equal(sample.annualSalary, 9600000);
assert.ok(sample.employeeSI > 0);
assert.ok(sample.employerSI > 0);
assert.ok(sample.personalTakeHome > 0);

const rows = calculator.runSearch("optimize", params);
assert.ok(rows.length > 100);
assert.ok(rows[0].feasible);
assert.ok(rows[0].metric >= rows[1].metric);
assert.equal(rows[0].monthly % params.step, 0);

const csv = app.buildCsv(rows.slice(0, 2));
assert.ok(csv.includes("月額役員報酬"));
assert.equal(csv.split("\n").length, 3);

console.log("All tests passed.");
