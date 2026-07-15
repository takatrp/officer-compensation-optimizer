import assert from "node:assert/strict";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const app = require("../app.js");

assert.equal(app.VERSION, "1.18");
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
assert.equal(Object.keys(app.strategyPresets).length, 6);
assert.equal(app.strategyPresets.balanced.includeDividend, false);
assert.equal(app.strategyPresets.retainedTarget.objective, "retainedTarget");
assert.equal(app.strategyPresets.retainedTarget.includeDividend, false);
assert.equal(app.strategyPresets.dividendUse.includeDividend, true);
assert.equal(app.strategyPresets.personalCash.objective, "personalCash");
assert.deepEqual(app.presetControlledIds, ["objective","includeDividend","divPolicy","fixedPayout","targetRetained","minRetained","noLoss","applyDividendCredit"]);
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
assert.equal(params.roleMode, "single");
assert.equal(params.step, 50000);
assert.equal(params.includeDividend, false);
assert.equal(params.divPolicy, "none");
const sample = calculator.simulate(800000, 0, params);

assert.equal(sample.feasible, true);
assert.equal(sample.annualSalary, 9600000);
assert.ok(sample.employeeSI > 0);
assert.ok(sample.employerSI > 0);
assert.ok(sample.personalTakeHome > 0);

const rows = calculator.runSearch("optimize", params);
assert.ok(rows.length > 10);
assert.ok(rows[0].feasible);
assert.ok(rows[0].metric >= rows[1].metric);
assert.equal(rows[0].monthly % params.step, 0);
assert.ok(rows.every((row) => row.payoutRate === 0));

const dividendParams = {...params, includeDividend:true, divPolicy:"optimize"};
const dividendRows = calculator.runSearch("optimize", dividendParams);
assert.ok(dividendRows.length > rows.length);
assert.ok(dividendRows.some((row) => row.payoutRate > 0));

const retainedTargetParams = {
  ...params,
  includeDividend:true,
  objective:"retainedTarget",
  divPolicy:"optimize",
  targetRetained:5000000,
  minRetained:0
};
const retainedTargetRows = calculator.runSearch("optimize", retainedTargetParams);
assert.ok(retainedTargetRows.length > 100);
assert.ok(retainedTargetRows[0].retainedGap <= retainedTargetRows[1].retainedGap + 1);
assert.ok(Math.abs(retainedTargetRows[0].retained - retainedTargetParams.targetRetained) <= retainedTargetRows[0].retainedGap + 1);

const targetVisual = app.targetVisualData(5000000, 4846000, 1650000);
assert.equal(targetVisual.difference, -154000);
assert.equal(targetVisual.gap, 154000);
assert.ok(Math.abs(targetVisual.targetRatio - 0.9692) < 0.00001);
assert.ok(targetVisual.bestGapWidth < targetVisual.currentGapWidth);
assert.equal(targetVisual.improvement, 1496000);
const negativeTargetVisual = app.targetVisualData(5000000, -1000000, 2000000);
assert.equal(negativeTargetVisual.difference, -6000000);
assert.equal(negativeTargetVisual.actualWidth, 0);

const coupleParams = {
  ...params,
  roleMode:"coupleSplit",
  coupleTotalMonthly:1200000,
  couplePrimaryMonthly:800000,
  minRetained:0,
  share:0.5,
  spouseShare:0.5,
  spouseOtherIncome:0,
  spouseOtherDedN:0,
  spouseOtherDedR:0,
  spouseCareApplicable:true,
  spouseSocialApplicable:true
};
const coupleSample = calculator.simulateCoupleSplit(700000, 0, coupleParams);
assert.equal(coupleSample.roleMode, "coupleSplit");
assert.equal(coupleSample.primaryMonthly, 700000);
assert.equal(coupleSample.spouseMonthly, 500000);
assert.equal(coupleSample.totalMonthly, 1200000);
assert.equal(coupleSample.annualSalary, 14400000);
assert.equal(coupleSample.people.length, 2);
assert.ok(coupleSample.employeeSI > sample.employeeSI);
assert.ok(coupleSample.personalTakeHome > 0);

const coupleRows = calculator.runSearch("none", coupleParams);
assert.ok(coupleRows.length > 10);
assert.equal(coupleRows[0].totalMonthly, 1200000);
assert.equal(coupleRows[0].primaryMonthly + coupleRows[0].spouseMonthly, 1200000);

const coupleGridParams = {
  ...coupleParams,
  roleMode:"coupleGrid",
  couplePrimaryMonthly:600000,
  coupleSpouseMonthly:400000,
  couplePrimaryMaxMonthly:800000,
  coupleSpouseMaxMonthly:600000,
  step:100000
};
const coupleGridSample = calculator.simulateCoupleGrid(700000, 300000, 0, coupleGridParams);
assert.equal(coupleGridSample.roleMode, "coupleGrid");
assert.equal(coupleGridSample.primaryMonthly, 700000);
assert.equal(coupleGridSample.spouseMonthly, 300000);
assert.equal(coupleGridSample.totalMonthly, 1000000);
assert.equal(coupleGridSample.people.length, 2);

const coupleGridRows = calculator.runSearch("none", coupleGridParams);
assert.ok(coupleGridRows.length > 20);
assert.ok(coupleGridRows.every((row) => row.primaryMonthly <= 800000));
assert.ok(coupleGridRows.every((row) => row.spouseMonthly <= 600000));
assert.ok(new Set(coupleGridRows.map((row) => row.totalMonthly)).size > 1);

const bucketed = app.bestByMonthlyBucket(rows, 100000);
assert.ok(bucketed.length > 5);
assert.equal(new Set(bucketed.slice(0, 12).map((row) => row.monthlyBucket)).size, 12);
assert.match(app.monthlyBucketLabel({...bucketed[0], monthlyBucket:100000, monthlyBucketSize:100000, monthly:160000}), /月額10万〜19万（16万が最良）/);

const csv = app.buildCsv(rows.slice(0, 2));
assert.ok(csv.includes("月額役員報酬"));
assert.equal(csv.split("\n").length, 3);

const coupleCsv = app.buildCsv(coupleGridRows.slice(0, 1));
assert.ok(coupleCsv.includes("夫婦役員（個別探索）"));

console.log("All tests passed.");
