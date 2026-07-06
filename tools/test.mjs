import assert from "node:assert/strict";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const app = require("../app.js");

assert.equal(app.VERSION, "1.7");
assert.equal(app.parseNumber("20,000,000"), 20000000);
assert.equal(app.formatMoneyText("3000000"), "3,000,000");
assert.equal(app.formatMoneyText("3,000,000"), "3,000,000");
assert.equal(app.formatMoneyText(""), "");
assert.equal(app.formatInput("shareRate", 100), "100");
assert.equal(app.formatInput("fixedPayout", 50), "50");
assert.equal(app.formatInput("healthRate", 10.12), "10.12");
assert.equal(app.formatInput("pensionRate", 18.3), "18.30");
assert.equal(app.prefectureRates.length, 47);
assert.equal(app.getPrefectureRate("兵庫").healthRate, 10.12);
assert.equal(app.getPrefectureRate("佐賀").healthRate, 10.55);
assert.equal(app.getPrefectureRate("存在しない支部"), null);
assert.equal(Object.keys(app.strategyPresets).length, 5);
assert.equal(app.strategyPresets.balanced.divPolicy, "fixed");
assert.equal(app.strategyPresets.companyReserve.divPolicy, "none");
assert.equal(app.strategyPresets.personalCash.objective, "personalCash");
assert.deepEqual(app.presetControlledIds, ["objective","divPolicy","fixedPayout","minRetained","noLoss","applyDividendCredit"]);
assert.equal(app.salaryDeduction(1900000, "r7"), 650000);
assert.equal(app.salaryDeduction(1900000, "r8r9"), 740000);
assert.equal(app.salaryDeduction(2200000, "r8r9"), 740000);
assert.equal(app.salaryDeduction(1900000, "r10"), 690000);
assert.equal(app.salaryDeduction(3600000, "r8r9"), 1160000);
assert.equal(app.basicNational(4000000, "r8r9"), 1040000);
assert.equal(app.basicNational(4890000, "r8r9"), 1040000);
assert.equal(app.basicNational(4890001, "r8r9"), 670000);
assert.equal(app.basicNational(6550000, "r8r9"), 670000);
assert.equal(app.basicNational(6550001, "r8r9"), 620000);
assert.equal(app.basicNational(1320000, "r10"), 990000);
assert.equal(app.basicNational(1320001, "r10"), 620000);
assert.equal(app.basicNational(1320001, "r7"), 880000);
assert.equal(app.nationalTaxBeforeCredits(7000000), 974000);

const calculator = app.createCalculator();
const params = app.defaultParams();
assert.equal(params.step, 50000);
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

const bucketed = app.bestByMonthlyBucket(rows, 100000);
assert.ok(bucketed.length > 5);
assert.equal(new Set(bucketed.slice(0, 12).map((row) => row.monthlyBucket)).size, 12);
assert.match(app.monthlyBucketLabel({...bucketed[0], monthlyBucket:100000, monthlyBucketSize:100000, monthly:160000}), /月額10万〜19万（16万が最良）/);

const csv = app.buildCsv(rows.slice(0, 2));
assert.ok(csv.includes("月額役員報酬"));
assert.equal(csv.split("\n").length, 3);

console.log("All tests passed.");
