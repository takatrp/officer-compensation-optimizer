(function(root, factory){
  const api = factory();
  if(typeof module === "object" && module.exports){
    module.exports = api;
  }
  root.OfficerCompensation = api;
  if(root.document){
    root.addEventListener("DOMContentLoaded", api.boot);
  }
})(typeof globalThis !== "undefined" ? globalThis : window, function(){
  "use strict";

  const VERSION = "1.5";

  const healthBands = [
    [0,63000,58000],[63000,73000,68000],[73000,83000,78000],[83000,93000,88000],[93000,101000,98000],
    [101000,107000,104000],[107000,114000,110000],[114000,122000,118000],[122000,130000,126000],[130000,138000,134000],
    [138000,146000,142000],[146000,155000,150000],[155000,165000,160000],[165000,175000,170000],[175000,185000,180000],
    [185000,195000,190000],[195000,210000,200000],[210000,230000,220000],[230000,250000,240000],[250000,270000,260000],
    [270000,290000,280000],[290000,310000,300000],[310000,330000,320000],[330000,350000,340000],[350000,370000,360000],
    [370000,395000,380000],[395000,425000,410000],[425000,455000,440000],[455000,485000,470000],[485000,515000,500000],
    [515000,545000,530000],[545000,575000,560000],[575000,605000,590000],[605000,635000,620000],[635000,665000,650000],
    [665000,695000,680000],[695000,730000,710000],[730000,770000,750000],[770000,810000,790000],[810000,855000,830000],
    [855000,905000,880000],[905000,955000,930000],[955000,1005000,980000],[1005000,1055000,1030000],[1055000,1115000,1090000],
    [1115000,1175000,1150000],[1175000,1235000,1210000],[1235000,1295000,1270000],[1295000,1355000,1330000],[1355000,Infinity,1390000]
  ];

  const pensionBands = [
    [0,93000,88000],[93000,101000,98000],[101000,107000,104000],[107000,114000,110000],[114000,122000,118000],
    [122000,130000,126000],[130000,138000,134000],[138000,146000,142000],[146000,155000,150000],[155000,165000,160000],
    [165000,175000,170000],[175000,185000,180000],[185000,195000,190000],[195000,210000,200000],[210000,230000,220000],
    [230000,250000,240000],[250000,270000,260000],[270000,290000,280000],[290000,310000,300000],[310000,330000,320000],
    [330000,350000,340000],[350000,370000,360000],[370000,395000,380000],[395000,425000,410000],[425000,455000,440000],
    [455000,485000,470000],[485000,515000,500000],[515000,545000,530000],[545000,575000,560000],[575000,605000,590000],
    [605000,635000,620000],[635000,Infinity,650000]
  ];

  const moneyIds = new Set([
    "preProfit","currentMonthly","maxMonthly","step","otherIncome","minRetained",
    "residentPerCapita","otherDedN","otherDedR","corpThreshold","corpFixedTax"
  ]);

  const percentIds = new Set([
    "shareRate","fixedPayout","surtaxRate","residentRate","healthRate","careRate",
    "supportRate","pensionRate","childContributionRate","corpLowRate","corpHighRate",
    "divCreditNLow","divCreditNHigh","divCreditRLow","divCreditRHigh"
  ]);

  const controlIds = [
    "preProfit","currentMonthly","maxMonthly","step","shareRate","otherIncome",
    "objective","divPolicy","fixedPayout","minRetained","noLoss","applyDividendCredit",
    "taxYear","surtaxRate","residentRate","residentPerCapita","otherDedN","otherDedR",
    "healthRate","careRate","supportRate","pensionRate","childContributionRate",
    "careApplicable","socialApplicable","corpLowRate","corpHighRate","corpThreshold",
    "corpFixedTax","divCreditNLow","divCreditNHigh","divCreditRLow","divCreditRHigh"
  ];

  const strategyPresets = Object.freeze({
    balanced: Object.freeze({
      objective:"ownerTotal",
      divPolicy:"fixed",
      fixedPayout:"30",
      minRetained:"3,000,000",
      noLoss:true,
      applyDividendCredit:true
    }),
    ownerTotal: Object.freeze({
      objective:"ownerTotal",
      divPolicy:"optimize",
      fixedPayout:"0",
      minRetained:"0",
      noLoss:true,
      applyDividendCredit:true
    }),
    personalCash: Object.freeze({
      objective:"personalCash",
      divPolicy:"optimize",
      fixedPayout:"100",
      minRetained:"0",
      noLoss:true,
      applyDividendCredit:true
    }),
    companyReserve: Object.freeze({
      objective:"ownerTotal",
      divPolicy:"none",
      fixedPayout:"0",
      minRetained:"5,000,000",
      noLoss:true,
      applyDividendCredit:true
    }),
    dividendUse: Object.freeze({
      objective:"personalCash",
      divPolicy:"fixed",
      fixedPayout:"50",
      minRetained:"1,000,000",
      noLoss:true,
      applyDividendCredit:true
    })
  });

  const presetControlledIds = Object.freeze([
    "objective","divPolicy","fixedPayout","minRetained","noLoss","applyDividendCredit"
  ]);

  const defaults = Object.freeze({
    preProfit:"20,000,000",
    currentMonthly:"800,000",
    maxMonthly:"2,000,000",
    step:"50,000",
    shareRate:"100",
    otherIncome:"0",
    strategyPreset:"balanced",
    objective:"ownerTotal",
    divPolicy:"fixed",
    fixedPayout:"30",
    minRetained:"3,000,000",
    noLoss:true,
    applyDividendCredit:true,
    taxYear:"r8r9",
    surtaxRate:"2.1",
    residentRate:"10.0",
    residentPerCapita:"5,000",
    otherDedN:"0",
    otherDedR:"0",
    healthRate:"10.12",
    careRate:"1.62",
    supportRate:"0.23",
    pensionRate:"18.30",
    childContributionRate:"0.36",
    careApplicable:true,
    socialApplicable:true,
    corpLowRate:"25.0",
    corpHighRate:"34.0",
    corpThreshold:"8,000,000",
    corpFixedTax:"70,000",
    divCreditNLow:"10.0",
    divCreditNHigh:"5.0",
    divCreditRLow:"2.8",
    divCreditRHigh:"1.4"
  });

  const balanceColors = {
    personal:"#047857",
    retained:"#2458d3",
    personalLoad:"#b45309",
    companyLoad:"#6d4aff"
  };

  let latestRows = [];
  let latestState = null;

  function parseNumber(value){
    const normalized = String(value || "")
      .replace(/,/g,"")
      .replace(/円/g,"")
      .replace(/万円/g,"0000")
      .replace(/%/g,"")
      .replace(/％/g,"")
      .trim();
    const n = Number(normalized);
    return Number.isFinite(n) ? n : 0;
  }

  function clamp(value, min, max){
    return Math.min(max, Math.max(min, value));
  }

  function formatPlain(n, digits){
    if(!Number.isFinite(n)) return "0";
    return n.toLocaleString("ja-JP", {
      maximumFractionDigits: digits,
      minimumFractionDigits: digits
    });
  }

  function formatInput(id, n){
    if(moneyIds.has(id)) return Math.round(n).toLocaleString("ja-JP");
    if(percentIds.has(id)){
      const formatted = formatPlain(n, Number.isInteger(n) ? 0 : 2);
      return formatted.includes(".") ? formatted.replace(/\.?0+$/,"") : formatted;
    }
    return String(n);
  }

  function formatMoneyText(value){
    const digits = String(value || "").replace(/[^\d]/g, "");
    if(!digits) return "";
    return Number(digits).toLocaleString("ja-JP");
  }

  function formatMoneyLive(input){
    const oldValue = input.value;
    const caret = input.selectionStart ?? oldValue.length;
    const digitsBeforeCaret = oldValue.slice(0, caret).replace(/[^\d]/g, "").length;
    const formatted = formatMoneyText(oldValue);
    input.value = formatted;

    let nextCaret = 0;
    let seenDigits = 0;
    while(nextCaret < formatted.length && seenDigits < digitsBeforeCaret){
      if(/\d/.test(formatted[nextCaret])) seenDigits += 1;
      nextCaret += 1;
    }

    try{
      input.setSelectionRange(nextCaret, nextCaret);
    }catch(error){
      // Some virtual keyboards do not expose a text selection range.
    }
  }

  function yen(n){
    if(!Number.isFinite(n)) return "-";
    const sign = n < 0 ? "▲" : "";
    return sign + Math.abs(Math.round(n)).toLocaleString("ja-JP") + "円";
  }

  function man(n){
    if(!Number.isFinite(n)) return "-";
    const sign = n < 0 ? "▲" : "";
    return sign + (Math.abs(n) / 10000).toLocaleString("ja-JP", {maximumFractionDigits:1}) + "万円";
  }

  function pct(n){
    return (n * 100).toLocaleString("ja-JP", {maximumFractionDigits:1}) + "%";
  }

  function rateLabel(n){
    return Number(n).toLocaleString("ja-JP", {maximumFractionDigits:2}) + "%";
  }

  function standard(monthly, bands){
    if(monthly <= 0) return 0;
    for(const [lo, hi, std] of bands){
      if(monthly >= lo && monthly < hi) return std;
    }
    return bands[bands.length - 1][2];
  }

  function salaryDeduction(salary, year){
    if(salary <= 0) return 0;
    let deduction;

    if(year === "r8r9" && salary <= 2200000){
      deduction = 740000;
    }else if(salary <= 1900000){
      deduction = year === "r10" ? Math.max(salary * 0.3 + 80000, 690000) : 650000;
    }
    else if(salary <= 3600000) deduction = salary * 0.3 + 80000;
    else if(salary <= 6600000) deduction = salary * 0.2 + 440000;
    else if(salary <= 8500000) deduction = salary * 0.1 + 1100000;
    else deduction = 1950000;

    if(year === "r10") deduction = Math.max(deduction, 690000);
    return Math.min(salary, deduction);
  }

  function salaryIncome(salary, year){
    return Math.max(0, salary - salaryDeduction(salary, year));
  }

  function basicNational(totalIncome, year){
    if(year === "r7"){
      if(totalIncome <= 1320000) return 950000;
      if(totalIncome <= 3360000) return 880000;
      if(totalIncome <= 4890000) return 680000;
      if(totalIncome <= 6550000) return 630000;
      if(totalIncome <= 23500000) return 580000;
      if(totalIncome <= 24000000) return 480000;
      if(totalIncome <= 24500000) return 320000;
      if(totalIncome <= 25000000) return 160000;
      return 0;
    }

    if(year === "r8r9"){
      if(totalIncome <= 4890000) return 1040000;
      if(totalIncome <= 6550000) return 670000;
      if(totalIncome <= 23500000) return 620000;
      if(totalIncome <= 24000000) return 480000;
      if(totalIncome <= 24500000) return 320000;
      if(totalIncome <= 25000000) return 160000;
      return 0;
    }

    if(totalIncome <= 1320000) return 990000;
    if(totalIncome <= 23500000) return 620000;
    if(totalIncome <= 24000000) return 480000;
    if(totalIncome <= 24500000) return 320000;
    if(totalIncome <= 25000000) return 160000;
    return 0;
  }

  function basicResident(totalIncome){
    if(totalIncome <= 24000000) return 430000;
    if(totalIncome <= 24500000) return 290000;
    if(totalIncome <= 25000000) return 150000;
    return 0;
  }

  function nationalTaxBeforeCredits(taxable){
    const x = Math.floor(Math.max(0, taxable) / 1000) * 1000;
    if(x <= 0) return 0;
    if(x <= 1949000) return x * 0.05;
    if(x <= 3299000) return x * 0.10 - 97500;
    if(x <= 6949000) return x * 0.20 - 427500;
    if(x <= 8999000) return x * 0.23 - 636000;
    if(x <= 17999000) return x * 0.33 - 1536000;
    if(x <= 39999000) return x * 0.40 - 2796000;
    return x * 0.45 - 4796000;
  }

  function dividendCredit(dividend, taxable, lowRate, highRate){
    if(dividend <= 0 || taxable <= 0) return 0;
    const threshold = 10000000;
    if(taxable <= threshold) return dividend * lowRate;
    if(taxable - dividend >= threshold) return dividend * highRate;
    const lowPortion = Math.max(0, threshold - (taxable - dividend));
    const highPortion = Math.max(0, dividend - lowPortion);
    return lowPortion * lowRate + highPortion * highRate;
  }

  function defaultParams(){
    return normalizeParams({
      preProfit:parseNumber(defaults.preProfit),
      currentMonthly:parseNumber(defaults.currentMonthly),
      maxMonthly:parseNumber(defaults.maxMonthly),
      step:parseNumber(defaults.step),
      share:parseNumber(defaults.shareRate) / 100,
      otherIncome:parseNumber(defaults.otherIncome),
      objective:defaults.objective,
      divPolicy:defaults.divPolicy,
      fixedPayout:parseNumber(defaults.fixedPayout) / 100,
      minRetained:parseNumber(defaults.minRetained),
      noLoss:defaults.noLoss,
      applyDividendCredit:defaults.applyDividendCredit,
      taxYear:defaults.taxYear,
      surtaxRate:parseNumber(defaults.surtaxRate) / 100,
      residentRate:parseNumber(defaults.residentRate) / 100,
      residentPerCapita:parseNumber(defaults.residentPerCapita),
      otherDedN:parseNumber(defaults.otherDedN),
      otherDedR:parseNumber(defaults.otherDedR),
      healthRate:parseNumber(defaults.healthRate) / 100,
      careRate:parseNumber(defaults.careRate) / 100,
      supportRate:parseNumber(defaults.supportRate) / 100,
      pensionRate:parseNumber(defaults.pensionRate) / 100,
      childContributionRate:parseNumber(defaults.childContributionRate) / 100,
      careApplicable:defaults.careApplicable,
      socialApplicable:defaults.socialApplicable,
      corpLowRate:parseNumber(defaults.corpLowRate) / 100,
      corpHighRate:parseNumber(defaults.corpHighRate) / 100,
      corpThreshold:parseNumber(defaults.corpThreshold),
      corpFixedTax:parseNumber(defaults.corpFixedTax),
      divCreditNLow:parseNumber(defaults.divCreditNLow) / 100,
      divCreditNHigh:parseNumber(defaults.divCreditNHigh) / 100,
      divCreditRLow:parseNumber(defaults.divCreditRLow) / 100,
      divCreditRHigh:parseNumber(defaults.divCreditRHigh) / 100
    });
  }

  function normalizeParams(p){
    return {
      ...p,
      preProfit:Math.max(0, p.preProfit),
      currentMonthly:Math.max(0, p.currentMonthly),
      maxMonthly:Math.max(1000, p.maxMonthly),
      step:clamp(Math.round(p.step || 10000), 1000, 500000),
      share:clamp(p.share, 0, 1),
      fixedPayout:clamp(p.fixedPayout, 0, 1),
      minRetained:Math.max(0, p.minRetained),
      surtaxRate:Math.max(0, p.surtaxRate),
      residentRate:Math.max(0, p.residentRate),
      healthRate:Math.max(0, p.healthRate),
      careRate:Math.max(0, p.careRate),
      supportRate:Math.max(0, p.supportRate),
      pensionRate:Math.max(0, p.pensionRate),
      childContributionRate:Math.max(0, p.childContributionRate),
      corpLowRate:Math.max(0, p.corpLowRate),
      corpHighRate:Math.max(0, p.corpHighRate),
      corpThreshold:Math.max(0, p.corpThreshold),
      corpFixedTax:Math.max(0, p.corpFixedTax),
      divCreditNLow:Math.max(0, p.divCreditNLow),
      divCreditNHigh:Math.max(0, p.divCreditNHigh),
      divCreditRLow:Math.max(0, p.divCreditRLow),
      divCreditRHigh:Math.max(0, p.divCreditRHigh)
    };
  }

  function createCalculator(){
    function socialInsurance(monthly, p){
      if(!p.socialApplicable || monthly <= 0){
        return {employee:0, employer:0, healthStd:0, pensionStd:0};
      }

      const healthStd = standard(monthly, healthBands);
      const pensionStd = standard(monthly, pensionBands);
      const healthTotal = p.healthRate + p.supportRate + (p.careApplicable ? p.careRate : 0);

      const employeeMonthly = healthStd * healthTotal / 2 + pensionStd * p.pensionRate / 2;
      const employerMonthly = healthStd * healthTotal / 2 + pensionStd * p.pensionRate / 2 + pensionStd * p.childContributionRate;

      return {
        employee:employeeMonthly * 12,
        employer:employerMonthly * 12,
        healthStd,
        pensionStd
      };
    }

    function corporateTax(income, p){
      let tax = Math.max(0, p.corpFixedTax);
      if(income > 0){
        tax += Math.min(income, p.corpThreshold) * p.corpLowRate;
        tax += Math.max(0, income - p.corpThreshold) * p.corpHighRate;
      }
      return tax;
    }

    function simulate(monthly, payoutRate, rawParams){
      const p = normalizeParams(rawParams);
      const annualSalary = monthly * 12;
      const si = socialInsurance(monthly, p);

      const companyTaxable = p.preProfit - annualSalary - si.employer;
      const corpTax = corporateTax(companyTaxable, p);
      const companyAfterTax = companyTaxable - corpTax;

      const effectivePayout = clamp(payoutRate, 0, 1);
      const totalDividend = companyAfterTax > 0 ? companyAfterTax * effectivePayout : 0;
      const retained = companyAfterTax - totalDividend;
      const ownerDividend = totalDividend * p.share;

      const salIncome = salaryIncome(annualSalary, p.taxYear);
      const totalIncome = salIncome + ownerDividend + p.otherIncome;

      const basicN = basicNational(totalIncome, p.taxYear);
      const taxableN = Math.max(0, totalIncome - si.employee - p.otherDedN - basicN);
      const baseN = nationalTaxBeforeCredits(taxableN);
      const creditNRaw = p.applyDividendCredit
        ? dividendCredit(ownerDividend, taxableN, p.divCreditNLow, p.divCreditNHigh)
        : 0;
      const creditN = Math.min(baseN, creditNRaw);
      const incomeTaxAfterCredit = Math.max(0, baseN - creditN);
      const surtax = incomeTaxAfterCredit * p.surtaxRate;
      const nationalTax = incomeTaxAfterCredit + surtax;

      const basicR = basicResident(totalIncome);
      const taxableR = Math.max(0, totalIncome - si.employee - p.otherDedR - basicR);
      const residentBefore = taxableR * p.residentRate;
      const creditRRaw = p.applyDividendCredit
        ? dividendCredit(ownerDividend, taxableR, p.divCreditRLow, p.divCreditRHigh)
        : 0;
      const creditR = Math.min(residentBefore, creditRRaw);
      const residentTax = Math.max(0, residentBefore - creditR) + (totalIncome > 0 ? p.residentPerCapita : 0);

      const personalTax = nationalTax + residentTax;
      const personalTakeHome = annualSalary + ownerDividend - si.employee - personalTax;
      const ownerRetainedValue = retained * p.share;
      const ownerTotal = personalTakeHome + ownerRetainedValue;
      const metric = p.objective === "personalCash" ? personalTakeHome : ownerTotal;

      const feasible =
        (!p.noLoss || companyAfterTax >= -1) &&
        retained >= p.minRetained - 1;

      return {
        monthly,
        annualSalary,
        payoutRate:effectivePayout,
        totalDividend,
        ownerDividend,
        retained,
        companyTaxable,
        corpTax,
        companyAfterTax,
        salIncome,
        totalIncome,
        employeeSI:si.employee,
        employerSI:si.employer,
        healthStd:si.healthStd,
        pensionStd:si.pensionStd,
        taxableN,
        taxableR,
        nationalTax,
        residentTax,
        personalTax,
        creditN,
        creditR,
        personalTakeHome,
        ownerRetainedValue,
        ownerTotal,
        metric,
        feasible
      };
    }

    function payoutRatesFor(policy, p){
      if(policy === "none") return [0];
      if(policy === "all") return [1];
      if(policy === "fixed") return [clamp(p.fixedPayout, 0, 1)];
      const rates = [];
      for(let i = 0; i <= 100; i += 1) rates.push(i / 100);
      return rates;
    }

    function runSearch(policy, rawParams){
      const p = normalizeParams(rawParams);
      const rows = [];
      const rates = payoutRatesFor(policy, p);
      const maxMonthly = Math.max(p.maxMonthly, p.currentMonthly);

      for(let monthly = 0; monthly <= maxMonthly + 0.1; monthly += p.step){
        for(const rate of rates){
          const row = simulate(monthly, rate, p);
          if(row.feasible) rows.push(row);
        }
      }

      rows.sort((a, b) => b.metric - a.metric);
      return rows;
    }

    return {socialInsurance, corporateTax, simulate, payoutRatesFor, runSearch};
  }

  function readParams(documentRef){
    const doc = documentRef || document;
    const n = (id) => parseNumber(doc.getElementById(id).value);
    const v = (id) => doc.getElementById(id).value;
    const checked = (id) => doc.getElementById(id).checked;

    return normalizeParams({
      preProfit:n("preProfit"),
      currentMonthly:n("currentMonthly"),
      maxMonthly:n("maxMonthly"),
      step:n("step"),
      share:n("shareRate") / 100,
      otherIncome:n("otherIncome"),
      objective:v("objective"),
      divPolicy:v("divPolicy"),
      fixedPayout:n("fixedPayout") / 100,
      minRetained:n("minRetained"),
      noLoss:checked("noLoss"),
      applyDividendCredit:checked("applyDividendCredit"),
      taxYear:v("taxYear"),
      surtaxRate:n("surtaxRate") / 100,
      residentRate:n("residentRate") / 100,
      residentPerCapita:n("residentPerCapita"),
      otherDedN:n("otherDedN"),
      otherDedR:n("otherDedR"),
      healthRate:n("healthRate") / 100,
      careRate:n("careRate") / 100,
      supportRate:n("supportRate") / 100,
      pensionRate:n("pensionRate") / 100,
      childContributionRate:n("childContributionRate") / 100,
      careApplicable:checked("careApplicable"),
      socialApplicable:checked("socialApplicable"),
      corpLowRate:n("corpLowRate") / 100,
      corpHighRate:n("corpHighRate") / 100,
      corpThreshold:n("corpThreshold"),
      corpFixedTax:n("corpFixedTax"),
      divCreditNLow:n("divCreditNLow") / 100,
      divCreditNHigh:n("divCreditNHigh") / 100,
      divCreditRLow:n("divCreditRLow") / 100,
      divCreditRHigh:n("divCreditRHigh") / 100
    });
  }

  function bestByMonth(rows){
    const map = new Map();
    for(const row of rows){
      const old = map.get(row.monthly);
      if(!old || row.metric > old.metric) map.set(row.monthly, row);
    }
    return Array.from(map.values()).sort((a, b) => a.monthly - b.monthly);
  }

  function bestByMonthlyBucket(rows, bucketSize){
    const map = new Map();
    for(const row of rows){
      const bucket = Math.floor(row.monthly / bucketSize) * bucketSize;
      const old = map.get(bucket);
      if(!old || row.metric > old.metric){
        map.set(bucket, {...row, monthlyBucket:bucket, monthlyBucketSize:bucketSize});
      }
    }
    return Array.from(map.values()).sort((a, b) => b.metric - a.metric);
  }

  function monthlyBucketLabel(row){
    if(!Number.isFinite(row.monthlyBucket) || !Number.isFinite(row.monthlyBucketSize)){
      return yen(row.monthly);
    }
    const bucketStart = row.monthlyBucket;
    const bucketEnd = bucketStart + row.monthlyBucketSize - 1;
    const startMan = Math.floor(bucketStart / 10000).toLocaleString("ja-JP");
    const endMan = Math.floor(bucketEnd / 10000).toLocaleString("ja-JP");
    const bestMan = Math.round(row.monthly / 10000).toLocaleString("ja-JP");
    return `月額${startMan}万〜${endMan}万（${bestMan}万が最良）`;
  }

  function bestCurrentScenario(calculator, p){
    const rates = calculator.payoutRatesFor(p.divPolicy, p);
    return rates
      .map((rate) => calculator.simulate(p.currentMonthly, rate, p))
      .filter((row) => row.feasible)
      .sort((a, b) => b.metric - a.metric)[0] || null;
  }

  function classifyDelta(n){
    if(n > 0) return "good";
    if(n < 0) return "bad";
    return "";
  }

  function signedMan(n){
    if(!Number.isFinite(n)) return "-";
    return `${n >= 0 ? "+" : ""}${man(n)}`;
  }

  function renderRecommendationReason(best, current, p){
    const metricLine = p.objective === "ownerTotal"
      ? `個人手取り ${man(best.personalTakeHome)} ＋ 持分相当の法人留保 ${man(best.ownerRetainedValue)} ＝ ${man(best.metric)}`
      : `個人手取り ${man(best.personalTakeHome)} ＝ ${man(best.metric)}`;
    const comparisonLine = current
      ? `比較月額比：評価指標 ${signedMan(best.metric - current.metric)}、個人手取り ${signedMan(best.personalTakeHome - current.personalTakeHome)}、法人留保 ${signedMan(best.retained - current.retained)}`
      : "比較月額は現在の条件では候補外です。";

    return `
      <div class="recommendation-reason">
        <p class="reason-title">なぜこの金額か</p>
        <p class="reason-lead">探索条件を満たす候補のうち、評価指標が最も高い月額です。</p>
        <dl class="reason-list">
          <div>
            <dt>評価指標</dt>
            <dd>${metricLine}</dd>
          </div>
          <div>
            <dt>個人手取り</dt>
            <dd>給与 ${man(best.annualSalary)} ＋ 本人配当 ${man(best.ownerDividend)} − 個人税 ${man(best.personalTax)} − 本人社保 ${man(best.employeeSI)} ＝ ${man(best.personalTakeHome)}</dd>
          </div>
          <div>
            <dt>法人側</dt>
            <dd>前利益 ${man(p.preProfit)} − 年間役員報酬 ${man(best.annualSalary)} − 会社社保 ${man(best.employerSI)} − 法人税等 ${man(best.corpTax)} ＝ 税引後利益 ${man(best.companyAfterTax)}<br>
            税引後利益 ${man(best.companyAfterTax)} − 配当総額 ${man(best.totalDividend)} ＝ 法人留保 ${man(best.retained)}</dd>
          </div>
          <div>
            <dt>比較</dt>
            <dd>${comparisonLine}</dd>
          </div>
        </dl>
      </div>
    `;
  }

  function renderResult(best, current, noDivBest, allDivBest, p){
    const box = document.getElementById("resultBox");
    const status = document.getElementById("calcStatus");

    if(!best){
      status.textContent = "候補なし";
      box.innerHTML = '<div class="message bad">条件を満たす候補がありません。探索上限、法人留保下限、赤字除外条件を見直してください。</div>';
      document.getElementById("balanceBox").innerHTML = "";
      document.getElementById("deltaBox").innerHTML = "";
      return;
    }

    status.textContent = "計算済み";
    const objectiveText = p.objective === "ownerTotal" ? "個人手取り＋持分相当の法人留保" : "個人キャッシュ";
    let messageClass = "";
    let message = "税・社会保険の概算だけで見ると、この条件では上記の月額が最も有利です。";

    if(p.objective === "ownerTotal" && best.payoutRate > 0.001 && p.share >= .999){
      messageClass = "warn";
      message = "100%オーナーで総手残りを重視する場合、配当は法人留保を個人へ移すたびに追加課税を受けます。個人資金化を重視する場面か確認してください。";
    }else if(p.objective === "personalCash" && best.totalDividend > 0){
      messageClass = "warn";
      message = "個人キャッシュ最大では法人留保より個人手取りを優先するため、配当が多く出やすくなります。運転資金と金融機関評価を別途確認してください。";
    }
    const shareNote = p.share < 0.999
      ? `<div class="message">持株割合 ${(p.share * 100).toFixed(0)}% のため、配当総額 ${man(best.totalDividend)} のうち本人受取は ${man(best.ownerDividend)} です。残りは他の株主に配当されます。</div>`
      : "";

    box.innerHTML = `
      <div class="hero-result">
        <div class="recommendation-main">
          <p class="label">おすすめ月額役員報酬</p>
          <p class="amount">${yen(best.monthly)}</p>
          <p class="sub">年間 ${yen(best.annualSalary)} ／ 配当率 ${pct(best.payoutRate)} ／ 評価基準 ${objectiveText}</p>
        </div>
        ${renderRecommendationReason(best, current, p)}
      </div>

      <div class="metric-grid">
        <div class="metric-card"><p class="k">個人手取り</p><p class="v">${man(best.personalTakeHome)}</p><p class="s">給与＋配当−個人税−本人社保</p></div>
        <div class="metric-card"><p class="k">法人留保</p><p class="v">${man(best.retained)}</p><p class="s">法人税後利益−配当総額</p></div>
        <div class="metric-card"><p class="k">税・社保合計</p><p class="v">${man(best.corpTax + best.personalTax + best.employeeSI + best.employerSI)}</p><p class="s">法人税、個人税、本人・会社社保</p></div>
        <div class="metric-card"><p class="k">配当総額</p><p class="v">${man(best.totalDividend)}</p><p class="s">本人受取 ${man(best.ownerDividend)}</p></div>
        <div class="metric-card"><p class="k">法人税等</p><p class="v">${man(best.corpTax)}</p><p class="s">課税所得 ${man(best.companyTaxable)}</p></div>
        <div class="metric-card"><p class="k">標準報酬月額</p><p class="v">${man(best.healthStd)}</p><p class="s">厚年 ${man(best.pensionStd)}</p></div>
      </div>

      <div class="message ${messageClass}">${message}</div>
      <div class="message">配当なし最適：${noDivBest ? `月額 ${yen(noDivBest.monthly)}、評価指標 ${man(noDivBest.metric)}` : "候補なし"}。 全額配当最適：${allDivBest ? `月額 ${yen(allDivBest.monthly)}、評価指標 ${man(allDivBest.metric)}` : "候補なし"}。</div>
      ${shareNote}
    `;
  }

  function balanceParts(row){
    return [
      {key:"personal", label:"個人手取り", value:Math.max(0, row.personalTakeHome), color:balanceColors.personal},
      {key:"retained", label:"法人留保", value:Math.max(0, row.retained), color:balanceColors.retained},
      {key:"personalLoad", label:"個人税・本人社保", value:Math.max(0, row.personalTax + row.employeeSI), color:balanceColors.personalLoad},
      {key:"companyLoad", label:"法人税・会社社保", value:Math.max(0, row.corpTax + row.employerSI), color:balanceColors.companyLoad}
    ];
  }

  function renderBalanceRow(title, row){
    if(!row){
      return `<div class="balance-row"><div class="message warn">${title}は現在の制約条件では候補外です。</div></div>`;
    }

    const parts = balanceParts(row);
    const total = parts.reduce((sum, part) => sum + part.value, 0) || 1;
    const segments = parts.map((part) => {
      const width = Math.max(0, part.value / total * 100);
      return `<div class="balance-part" title="${part.label} ${man(part.value)}" style="width:${width}%;background:${part.color}"></div>`;
    }).join("");
    const legend = parts.map((part) => `
      <div class="legend-item">
        <span class="legend-name"><i class="legend-swatch" style="background:${part.color}"></i>${part.label}</span>
        <strong>${man(part.value)}</strong>
      </div>
    `).join("");

    return `
      <div class="balance-row">
        <div class="balance-head"><span>${title}</span><span>月額 ${yen(row.monthly)} ／ 配当率 ${pct(row.payoutRate)}</span></div>
        <div class="balance-track">${segments}</div>
        <div class="balance-legend">${legend}</div>
      </div>
    `;
  }

  function renderBalance(best, current){
    const box = document.getElementById("balanceBox");
    box.innerHTML = `
      <div class="panel-head">
        <div>
          <p class="section-kicker">バランス</p>
          <h2>利益の配分イメージ</h2>
        </div>
      </div>
      ${renderBalanceRow("おすすめ案", best)}
      ${renderBalanceRow("比較月額", current)}
    `;
  }

  function renderDelta(best, current){
    const box = document.getElementById("deltaBox");
    if(!best || !current){
      box.innerHTML = "";
      return;
    }

    const deltas = [
      {label:"評価指標の差", value:best.metric - current.metric},
      {label:"個人手取りの差", value:best.personalTakeHome - current.personalTakeHome},
      {label:"法人留保の差", value:best.retained - current.retained}
    ];

    box.innerHTML = deltas.map((item) => `
      <div class="delta-card ${classifyDelta(item.value)}">
        <p class="k">${item.label}</p>
        <p class="v">${item.value >= 0 ? "+" : ""}${man(item.value)}</p>
      </div>
    `).join("");
  }

  function renderTable(rows){
    latestRows = rows;
    const box = document.getElementById("tableBox");
    if(!rows.length){
      box.innerHTML = '<div class="message bad">上位候補を表示できません。</div>';
      return;
    }

    const top = bestByMonthlyBucket(rows, 100000).slice(0, 12);
    box.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>順位</th>
            <th>月額</th>
            <th>配当率</th>
            <th>個人手取り</th>
            <th>法人留保</th>
            <th>総手残り</th>
            <th>税・社保</th>
          </tr>
        </thead>
        <tbody>
          ${top.map((row, index) => `
            <tr>
              <td>${index + 1}</td>
              <td>${monthlyBucketLabel(row)}</td>
              <td>${pct(row.payoutRate)}</td>
              <td>${man(row.personalTakeHome)}</td>
              <td>${man(row.retained)}</td>
              <td>${man(row.ownerTotal)}</td>
              <td>${man(row.corpTax + row.personalTax + row.employeeSI + row.employerSI)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function drawChart(rows, current){
    const canvas = document.getElementById("chart");
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const series = bestByMonth(rows);
    if(series.length < 2){
      ctx.fillStyle = "#667085";
      ctx.font = "16px sans-serif";
      ctx.fillText("条件を満たす候補がありません", 32, 50);
      return;
    }

    const pad = {l:74, r:54, t:30, b:54};
    const w = canvas.width - pad.l - pad.r;
    const h = canvas.height - pad.t - pad.b;
    const xs = series.map((row) => row.monthly);
    const values = series.flatMap((row) => [row.metric, row.personalTakeHome, row.retained]);
    const xmin = Math.min(...xs);
    const xmax = Math.max(...xs);
    const ymin = Math.min(...values);
    const ymax = Math.max(...values);
    const ypad = Math.max(100000, (ymax - ymin) * 0.08);

    const xScale = (x) => pad.l + (x - xmin) / (xmax - xmin || 1) * w;
    const yScale = (y) => pad.t + h - (y - (ymin - ypad)) / ((ymax + ypad) - (ymin - ypad) || 1) * h;

    ctx.lineWidth = 1;
    ctx.strokeStyle = "#d0d5dd";
    ctx.beginPath();
    ctx.moveTo(pad.l, pad.t);
    ctx.lineTo(pad.l, pad.t + h);
    ctx.lineTo(pad.l + w, pad.t + h);
    ctx.stroke();

    ctx.font = "12px sans-serif";
    ctx.fillStyle = "#667085";
    for(let i = 0; i <= 4; i += 1){
      const value = ymin - ypad + ((ymax + ypad) - (ymin - ypad)) * i / 4;
      const y = yScale(value);
      ctx.strokeStyle = "#eef0f4";
      ctx.beginPath();
      ctx.moveTo(pad.l, y);
      ctx.lineTo(pad.l + w, y);
      ctx.stroke();
      ctx.fillStyle = "#667085";
      ctx.fillText((value / 10000).toFixed(0) + "万円", 10, y + 4);
    }

    function drawLine(getValue, color, width){
      ctx.strokeStyle = color;
      ctx.lineWidth = width;
      ctx.setLineDash([]);
      ctx.beginPath();
      series.forEach((row, index) => {
        const x = xScale(row.monthly);
        const y = yScale(getValue(row));
        if(index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
    }

    drawLine((row) => row.personalTakeHome, "#047857", 2);
    drawLine((row) => row.retained, "#6d4aff", 2);
    drawLine((row) => row.metric, "#2458d3", 3);

    const best = rows[0];
    const bestX = xScale(best.monthly);
    const bestY = yScale(best.metric);
    ctx.fillStyle = "#2458d3";
    ctx.beginPath();
    ctx.arc(bestX, bestY, 5, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = "#172033";
    ctx.font = "13px sans-serif";
    const bestLabel = "おすすめ";
    const bestLabelWidth = ctx.measureText(bestLabel).width;
    const labelRightX = bestX + 8;
    const labelLeftX = bestX - bestLabelWidth - 8;
    const bestLabelX = labelRightX + bestLabelWidth <= pad.l + w ? labelRightX : Math.max(pad.l + 4, labelLeftX);
    ctx.fillText(bestLabel, bestLabelX, Math.max(pad.t + 13, bestY - 8));

    if(current){
      const x = xScale(current.monthly);
      ctx.strokeStyle = "#b45309";
      ctx.lineWidth = 2;
      ctx.setLineDash([5, 5]);
      ctx.beginPath();
      ctx.moveTo(x, pad.t);
      ctx.lineTo(x, pad.t + h);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = "#b45309";
      ctx.fillText("比較月額", x + 8, pad.t + 16);
    }

    ctx.fillStyle = "#667085";
    ctx.font = "12px sans-serif";
    ctx.fillText("月額役員報酬", pad.l + w - 82, canvas.height - 18);
    ctx.fillText("金額", pad.l, 18);
  }

  function buildCsv(rows){
    const headers = [
      "順位","月額役員報酬","年間役員報酬","配当率","配当総額","本人配当",
      "個人手取り","法人留保","オーナー総手残り","法人税等","法人課税所得",
      "本人社保","会社社保","個人税","所得税等","住民税"
    ];
    const lines = [headers.join(",")];
    rows.forEach((row, index) => {
      lines.push([
        index + 1,
        row.monthly,
        row.annualSalary,
        row.payoutRate,
        row.totalDividend,
        row.ownerDividend,
        row.personalTakeHome,
        row.retained,
        row.ownerTotal,
        row.corpTax,
        row.companyTaxable,
        row.employeeSI,
        row.employerSI,
        row.personalTax,
        row.nationalTax,
        row.residentTax
      ].map((value) => Math.round(value * 100) / 100).join(","));
    });
    return lines.join("\n");
  }

  function setControlValue(id, value){
    const el = document.getElementById(id);
    if(!el) return;
    if(el.type === "checkbox"){
      el.checked = Boolean(value);
      return;
    }
    el.value = value;
  }

  function applyDefaults(){
    Object.entries(defaults).forEach(([id, value]) => setControlValue(id, value));
    syncAllRangesFromInputs();
  }

  function applyStrategyPreset(value){
    const preset = strategyPresets[value];
    if(!preset) return;
    Object.entries(preset).forEach(([id, presetValue]) => setControlValue(id, presetValue));
    formatAllInputs();
    syncAllRangesFromInputs();
    updateDependentControls();
  }

  function markStrategyCustom(){
    const strategySelect = document.getElementById("strategyPreset");
    if(strategySelect && strategySelect.value !== "custom"){
      strategySelect.value = "custom";
    }
  }

  function syncAllRangesFromInputs(){
    document.querySelectorAll(".range[data-for]").forEach((range) => {
      syncRangeFromInput(range.dataset.for);
    });
  }

  function syncRangeFromInput(id){
    const target = document.getElementById(id);
    const range = document.querySelector(`.range[data-for="${id}"]`);
    if(!target || !range) return;
    const value = parseNumber(target.value);
    range.value = String(clamp(value, Number(range.min), Number(range.max)));
  }

  function formatAllInputs(){
    [...moneyIds, ...percentIds].forEach((id) => {
      const el = document.getElementById(id);
      if(el) el.value = formatInput(id, parseNumber(el.value));
    });
  }

  function updateDependentControls(){
    const policy = document.getElementById("divPolicy")?.value;
    const fixedInput = document.getElementById("fixedPayout");
    const fixedRange = document.querySelector('.range[data-for="fixedPayout"]');
    const disabled = policy !== "fixed";
    if(fixedInput) fixedInput.disabled = disabled;
    if(fixedRange) fixedRange.disabled = disabled;
  }

  function update(){
    const calculator = createCalculator();
    const p = readParams(document);
    const rows = calculator.runSearch(p.divPolicy, p);
    const noRows = calculator.runSearch("none", p);
    const allRows = calculator.runSearch("all", p);
    const current = bestCurrentScenario(calculator, p);
    const best = rows[0] || null;

    latestState = {rows, best, current, params:p};
    renderResult(best, current, noRows[0] || null, allRows[0] || null, p);
    renderBalance(best, current);
    renderDelta(best, current);
    renderTable(rows);
    drawChart(rows, current);
  }

  function queueUpdate(){
    if(queueUpdate.frame) cancelAnimationFrame(queueUpdate.frame);
    queueUpdate.frame = requestAnimationFrame(update);
  }

  function bindControls(){
    controlIds.forEach((id) => {
      const el = document.getElementById(id);
      if(!el) return;
      el.addEventListener("input", queueUpdate);
      el.addEventListener("change", queueUpdate);
      if(presetControlledIds.includes(id)){
        el.addEventListener("input", markStrategyCustom);
        el.addEventListener("change", markStrategyCustom);
      }
      if(moneyIds.has(id) || percentIds.has(id)){
        el.addEventListener("input", () => {
          if(moneyIds.has(id)) formatMoneyLive(el);
          syncRangeFromInput(id);
        });
        el.addEventListener("blur", () => {
          el.value = formatInput(id, parseNumber(el.value));
          syncAllRangesFromInputs();
          queueUpdate();
        });
      }
    });

    document.querySelectorAll(".range[data-for]").forEach((range) => {
      range.addEventListener("input", () => {
        const target = document.getElementById(range.dataset.for);
        if(!target) return;
        target.value = formatInput(range.dataset.for, Number(range.value));
        if(presetControlledIds.includes(range.dataset.for)) markStrategyCustom();
        queueUpdate();
      });
    });

    const strategySelect = document.getElementById("strategyPreset");
    strategySelect.addEventListener("change", () => {
      applyStrategyPreset(strategySelect.value);
      queueUpdate();
    });

    document.getElementById("resetDefaults").addEventListener("click", () => {
      applyDefaults();
      formatAllInputs();
      updateDependentControls();
      queueUpdate();
    });

    document.getElementById("copyCsv").addEventListener("click", async () => {
      const status = document.getElementById("copyStatus");
      const csv = buildCsv(latestRows.slice(0, 30));
      try{
        await navigator.clipboard.writeText(csv);
        status.textContent = "上位30件のCSVをコピーしました。";
      }catch(error){
        status.textContent = "コピーできませんでした。ダウンロードを使用してください。";
      }
    });

    document.getElementById("downloadCsv").addEventListener("click", () => {
      const csv = "\uFEFF" + buildCsv(latestRows);
      const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "officer-compensation-candidates.csv";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    });
  }

  function boot(){
    syncAllRangesFromInputs();
    formatAllInputs();
    bindControls();
    updateDependentControls();
    document.getElementById("divPolicy").addEventListener("change", updateDependentControls);
    update();
  }

  return {
    VERSION,
    boot,
    defaults,
    strategyPresets,
    presetControlledIds,
    defaultParams,
    createCalculator,
    bestByMonthlyBucket,
    monthlyBucketLabel,
    parseNumber,
    formatMoneyText,
    formatInput,
    salaryDeduction,
    salaryIncome,
    basicNational,
    basicResident,
    nationalTaxBeforeCredits,
    dividendCredit,
    bestByMonth,
    yen,
    man,
    rateLabel,
    buildCsv
  };
});
