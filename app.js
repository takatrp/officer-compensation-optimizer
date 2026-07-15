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

  const VERSION = "1.18";

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

  const prefectureRates = Object.freeze([
    {name:"北海道", healthRate:10.28},
    {name:"青森", healthRate:9.85},
    {name:"岩手", healthRate:9.51},
    {name:"宮城", healthRate:10.10},
    {name:"秋田", healthRate:10.01},
    {name:"山形", healthRate:9.75},
    {name:"福島", healthRate:9.50},
    {name:"茨城", healthRate:9.52},
    {name:"栃木", healthRate:9.82},
    {name:"群馬", healthRate:9.68},
    {name:"埼玉", healthRate:9.67},
    {name:"千葉", healthRate:9.73},
    {name:"東京", healthRate:9.85},
    {name:"神奈川", healthRate:9.92},
    {name:"新潟", healthRate:9.21},
    {name:"富山", healthRate:9.59},
    {name:"石川", healthRate:9.70},
    {name:"福井", healthRate:9.71},
    {name:"山梨", healthRate:9.55},
    {name:"長野", healthRate:9.63},
    {name:"岐阜", healthRate:9.80},
    {name:"静岡", healthRate:9.61},
    {name:"愛知", healthRate:9.93},
    {name:"三重", healthRate:9.77},
    {name:"滋賀", healthRate:9.88},
    {name:"京都", healthRate:9.89},
    {name:"大阪", healthRate:10.13},
    {name:"兵庫", healthRate:10.12},
    {name:"奈良", healthRate:9.91},
    {name:"和歌山", healthRate:10.06},
    {name:"鳥取", healthRate:9.86},
    {name:"島根", healthRate:9.94},
    {name:"岡山", healthRate:10.05},
    {name:"広島", healthRate:9.78},
    {name:"山口", healthRate:10.15},
    {name:"徳島", healthRate:10.24},
    {name:"香川", healthRate:10.02},
    {name:"愛媛", healthRate:9.98},
    {name:"高知", healthRate:10.05},
    {name:"福岡", healthRate:10.11},
    {name:"佐賀", healthRate:10.55},
    {name:"長崎", healthRate:10.06},
    {name:"熊本", healthRate:10.08},
    {name:"大分", healthRate:10.08},
    {name:"宮崎", healthRate:9.77},
    {name:"鹿児島", healthRate:10.13},
    {name:"沖縄", healthRate:9.44}
  ]);

  const moneyIds = new Set([
    "preProfit","currentMonthly","maxMonthly","step","otherIncome","targetRetained","minRetained",
    "coupleTotalMonthly","couplePrimaryMonthly","coupleSpouseMonthly",
    "couplePrimaryMaxMonthly","coupleSpouseMaxMonthly","spouseOtherIncome",
    "residentPerCapita","otherDedN","otherDedR","spouseOtherDedN","spouseOtherDedR",
    "corpThreshold","corpFixedTax"
  ]);

  const percentIds = new Set([
    "shareRate","spouseShareRate","fixedPayout","surtaxRate","residentRate",
    "healthRate","careRate","supportRate","pensionRate","childContributionRate",
    "corpLowRate","corpHighRate","divCreditNLow","divCreditNHigh","divCreditRLow","divCreditRHigh"
  ]);

  const controlIds = [
    "roleMode","preProfit","currentMonthly","maxMonthly","step","shareRate","otherIncome",
    "coupleTotalMonthly","couplePrimaryMonthly","coupleSpouseMonthly",
    "couplePrimaryMaxMonthly","coupleSpouseMaxMonthly","spouseShareRate","spouseOtherIncome",
    "objective","includeDividend","divPolicy","fixedPayout","targetRetained","minRetained","noLoss","applyDividendCredit",
    "taxYear","surtaxRate","residentRate","residentPerCapita","otherDedN","otherDedR",
    "spouseOtherDedN","spouseOtherDedR",
    "healthRate","careRate","supportRate","pensionRate","childContributionRate",
    "careApplicable","socialApplicable","spouseCareApplicable","spouseSocialApplicable",
    "corpLowRate","corpHighRate","corpThreshold",
    "corpFixedTax","divCreditNLow","divCreditNHigh","divCreditRLow","divCreditRHigh"
  ];

  const strategyPresets = Object.freeze({
    balanced: Object.freeze({
      objective:"ownerTotal",
      includeDividend:false,
      divPolicy:"optimize",
      fixedPayout:"0",
      targetRetained:"5,000,000",
      minRetained:"3,000,000",
      noLoss:true,
      applyDividendCredit:true
    }),
    retainedTarget: Object.freeze({
      objective:"retainedTarget",
      includeDividend:false,
      divPolicy:"optimize",
      fixedPayout:"0",
      targetRetained:"5,000,000",
      minRetained:"0",
      noLoss:true,
      applyDividendCredit:true
    }),
    ownerTotal: Object.freeze({
      objective:"ownerTotal",
      includeDividend:false,
      divPolicy:"optimize",
      fixedPayout:"0",
      targetRetained:"5,000,000",
      minRetained:"0",
      noLoss:true,
      applyDividendCredit:true
    }),
    personalCash: Object.freeze({
      objective:"personalCash",
      includeDividend:false,
      divPolicy:"optimize",
      fixedPayout:"0",
      targetRetained:"0",
      minRetained:"0",
      noLoss:true,
      applyDividendCredit:true
    }),
    companyReserve: Object.freeze({
      objective:"ownerTotal",
      includeDividend:false,
      divPolicy:"optimize",
      fixedPayout:"0",
      targetRetained:"5,000,000",
      minRetained:"5,000,000",
      noLoss:true,
      applyDividendCredit:true
    }),
    dividendUse: Object.freeze({
      objective:"personalCash",
      includeDividend:true,
      divPolicy:"fixed",
      fixedPayout:"50",
      targetRetained:"3,000,000",
      minRetained:"1,000,000",
      noLoss:true,
      applyDividendCredit:true
    })
  });

  const presetControlledIds = Object.freeze([
    "objective","includeDividend","divPolicy","fixedPayout","targetRetained","minRetained","noLoss","applyDividendCredit"
  ]);

  const defaults = Object.freeze({
    preProfit:"20,000,000",
    roleMode:"single",
    currentMonthly:"800,000",
    maxMonthly:"2,000,000",
    step:"50,000",
    shareRate:"100",
    otherIncome:"0",
    coupleTotalMonthly:"1,200,000",
    couplePrimaryMonthly:"800,000",
    coupleSpouseMonthly:"400,000",
    couplePrimaryMaxMonthly:"1,200,000",
    coupleSpouseMaxMonthly:"1,200,000",
    spouseShareRate:"0",
    spouseOtherIncome:"0",
    strategyPreset:"balanced",
    objective:"ownerTotal",
    includeDividend:false,
    divPolicy:"optimize",
    fixedPayout:"0",
    targetRetained:"5,000,000",
    minRetained:"3,000,000",
    noLoss:true,
    applyDividendCredit:true,
    taxYear:"r8r9",
    surtaxRate:"2.1",
    residentRate:"10.0",
    residentPerCapita:"5,000",
    otherDedN:"0",
    otherDedR:"0",
    spouseOtherDedN:"0",
    spouseOtherDedR:"0",
    prefecture:"兵庫",
    healthRate:"10.12",
    careRate:"1.62",
    supportRate:"0.23",
    pensionRate:"18.30",
    childContributionRate:"0.36",
    careApplicable:true,
    socialApplicable:true,
    spouseCareApplicable:true,
    spouseSocialApplicable:true,
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

  const reasonColors = {
    salary:"#2f7f8f",
    retained:"#2458d3",
    takeHome:"#047857",
    tax:"#6d4aff",
    social:"#b45309",
    dividend:"#b42318"
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
      if(["healthRate","careRate","supportRate","pensionRate","childContributionRate"].includes(id)){
        return formatPlain(n, 2);
      }
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

  function isCoupleModeValue(roleMode){
    return roleMode === "coupleSplit" || roleMode === "coupleGrid";
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
      roleMode:defaults.roleMode,
      currentMonthly:parseNumber(defaults.currentMonthly),
      maxMonthly:parseNumber(defaults.maxMonthly),
      step:parseNumber(defaults.step),
      share:parseNumber(defaults.shareRate) / 100,
      otherIncome:parseNumber(defaults.otherIncome),
      coupleTotalMonthly:parseNumber(defaults.coupleTotalMonthly),
      couplePrimaryMonthly:parseNumber(defaults.couplePrimaryMonthly),
      coupleSpouseMonthly:parseNumber(defaults.coupleSpouseMonthly),
      couplePrimaryMaxMonthly:parseNumber(defaults.couplePrimaryMaxMonthly),
      coupleSpouseMaxMonthly:parseNumber(defaults.coupleSpouseMaxMonthly),
      spouseShare:parseNumber(defaults.spouseShareRate) / 100,
      spouseOtherIncome:parseNumber(defaults.spouseOtherIncome),
      objective:defaults.objective,
      includeDividend:defaults.includeDividend,
      divPolicy:defaults.divPolicy,
      fixedPayout:parseNumber(defaults.fixedPayout) / 100,
      targetRetained:parseNumber(defaults.targetRetained),
      minRetained:parseNumber(defaults.minRetained),
      noLoss:defaults.noLoss,
      applyDividendCredit:defaults.applyDividendCredit,
      taxYear:defaults.taxYear,
      surtaxRate:parseNumber(defaults.surtaxRate) / 100,
      residentRate:parseNumber(defaults.residentRate) / 100,
      residentPerCapita:parseNumber(defaults.residentPerCapita),
      otherDedN:parseNumber(defaults.otherDedN),
      otherDedR:parseNumber(defaults.otherDedR),
      spouseOtherDedN:parseNumber(defaults.spouseOtherDedN),
      spouseOtherDedR:parseNumber(defaults.spouseOtherDedR),
      healthRate:parseNumber(defaults.healthRate) / 100,
      careRate:parseNumber(defaults.careRate) / 100,
      supportRate:parseNumber(defaults.supportRate) / 100,
      pensionRate:parseNumber(defaults.pensionRate) / 100,
      childContributionRate:parseNumber(defaults.childContributionRate) / 100,
      careApplicable:defaults.careApplicable,
      socialApplicable:defaults.socialApplicable,
      spouseCareApplicable:defaults.spouseCareApplicable,
      spouseSocialApplicable:defaults.spouseSocialApplicable,
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
    const roleMode = isCoupleModeValue(p.roleMode) ? p.roleMode : "single";
    const share = clamp(p.share, 0, 1);
    const spouseShare = clamp(p.spouseShare || 0, 0, 1);
    const coupleTotalMonthly = Math.max(0, p.coupleTotalMonthly || 0);
    const couplePrimaryMaxMonthly = Math.max(1000, p.couplePrimaryMaxMonthly || 0);
    const coupleSpouseMaxMonthly = Math.max(1000, p.coupleSpouseMaxMonthly || 0);
    const couplePrimaryMonthly = roleMode === "coupleSplit"
      ? clamp(p.couplePrimaryMonthly || 0, 0, coupleTotalMonthly)
      : clamp(p.couplePrimaryMonthly || 0, 0, couplePrimaryMaxMonthly);
    const includeDividend = p.includeDividend === true;
    const requestedDivPolicy = ["none","all","fixed","optimize"].includes(p.divPolicy) ? p.divPolicy : "optimize";
    return {
      ...p,
      roleMode,
      preProfit:Math.max(0, p.preProfit),
      currentMonthly:Math.max(0, p.currentMonthly),
      maxMonthly:Math.max(1000, p.maxMonthly),
      step:clamp(Math.round(p.step || 10000), 1000, 500000),
      share,
      coupleTotalMonthly,
      couplePrimaryMonthly,
      coupleSpouseMonthly:clamp(p.coupleSpouseMonthly || 0, 0, coupleSpouseMaxMonthly),
      couplePrimaryMaxMonthly,
      coupleSpouseMaxMonthly,
      spouseShare,
      spouseOtherIncome:Math.max(0, p.spouseOtherIncome || 0),
      includeDividend,
      divPolicy:includeDividend ? requestedDivPolicy : "none",
      fixedPayout:clamp(p.fixedPayout, 0, 1),
      targetRetained:Math.max(0, p.targetRetained || 0),
      minRetained:Math.max(0, p.minRetained),
      surtaxRate:Math.max(0, p.surtaxRate),
      residentRate:Math.max(0, p.residentRate),
      otherDedN:Math.max(0, p.otherDedN || 0),
      otherDedR:Math.max(0, p.otherDedR || 0),
      spouseOtherDedN:Math.max(0, p.spouseOtherDedN || 0),
      spouseOtherDedR:Math.max(0, p.spouseOtherDedR || 0),
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

  function compareRows(a, b){
    if(!a) return 1;
    if(!b) return -1;
    if(a.objective === "retainedTarget" || b.objective === "retainedTarget"){
      const gapDiff = (a.retainedGap ?? Infinity) - (b.retainedGap ?? Infinity);
      if(Math.abs(gapDiff) > 1) return gapDiff;
      const personalDiff = b.personalTakeHome - a.personalTakeHome;
      if(Math.abs(personalDiff) > 1) return personalDiff;
      return b.ownerTotal - a.ownerTotal;
    }
    return b.metric - a.metric;
  }

  function rowIsBetter(candidate, current){
    return !current || compareRows(candidate, current) < 0;
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

    function personConfig(key, label, monthly, share, otherIncome, otherDedN, otherDedR, socialApplicable, careApplicable){
      return {
        key,
        label,
        monthly:Math.max(0, monthly),
        annualSalary:Math.max(0, monthly) * 12,
        share:clamp(share, 0, 1),
        otherIncome:Math.max(0, otherIncome || 0),
        otherDedN:Math.max(0, otherDedN || 0),
        otherDedR:Math.max(0, otherDedR || 0),
        socialApplicable:Boolean(socialApplicable),
        careApplicable:Boolean(careApplicable)
      };
    }

    function primaryPerson(monthly, p){
      return personConfig(
        "primary",
        p.roleMode === "coupleSplit" ? "役員A" : "本人",
        monthly,
        p.share,
        p.otherIncome,
        p.otherDedN,
        p.otherDedR,
        p.socialApplicable,
        p.careApplicable
      );
    }

    function spousePerson(monthly, p){
      return personConfig(
        "spouse",
        "役員B",
        monthly,
        p.spouseShare,
        p.spouseOtherIncome,
        p.spouseOtherDedN,
        p.spouseOtherDedR,
        p.spouseSocialApplicable,
        p.spouseCareApplicable
      );
    }

    function personSocialInsurance(person, p){
      return socialInsurance(monthlyForPerson(person), {
        ...p,
        socialApplicable:person.socialApplicable,
        careApplicable:person.careApplicable
      });
    }

    function monthlyForPerson(person){
      return person.monthly;
    }

    function calculatePersonTax(person, ownerDividend, p){
      const si = personSocialInsurance(person, p);
      const salIncome = salaryIncome(person.annualSalary, p.taxYear);
      const totalIncome = salIncome + ownerDividend + person.otherIncome;

      const basicN = basicNational(totalIncome, p.taxYear);
      const taxableN = Math.max(0, totalIncome - si.employee - person.otherDedN - basicN);
      const baseN = nationalTaxBeforeCredits(taxableN);
      const creditNRaw = p.applyDividendCredit
        ? dividendCredit(ownerDividend, taxableN, p.divCreditNLow, p.divCreditNHigh)
        : 0;
      const creditN = Math.min(baseN, creditNRaw);
      const incomeTaxAfterCredit = Math.max(0, baseN - creditN);
      const surtax = incomeTaxAfterCredit * p.surtaxRate;
      const nationalTax = incomeTaxAfterCredit + surtax;

      const basicR = basicResident(totalIncome);
      const taxableR = Math.max(0, totalIncome - si.employee - person.otherDedR - basicR);
      const residentBefore = taxableR * p.residentRate;
      const creditRRaw = p.applyDividendCredit
        ? dividendCredit(ownerDividend, taxableR, p.divCreditRLow, p.divCreditRHigh)
        : 0;
      const creditR = Math.min(residentBefore, creditRRaw);
      const residentTax = Math.max(0, residentBefore - creditR) + (totalIncome > 0 ? p.residentPerCapita : 0);

      const personalTax = nationalTax + residentTax;
      const personalTakeHome = person.annualSalary + ownerDividend - si.employee - personalTax;

      return {
        ...person,
        ownerDividend,
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
        personalTakeHome
      };
    }

    function allocateDividends(people, totalDividend){
      const shareTotal = people.reduce((sum, person) => sum + person.share, 0);
      const householdShare = Math.min(1, shareTotal);
      const householdDividend = totalDividend * householdShare;
      return people.map((person) => ({
        person,
        ownerDividend:shareTotal > 0 ? householdDividend * person.share / shareTotal : 0
      }));
    }

    function simulatePeople(peopleInput, payoutRate, rawParams, meta){
      const p = normalizeParams(rawParams);
      const people = peopleInput.filter((person) => person.monthly > 0 || person.share > 0 || person.otherIncome > 0);
      const totalAnnualSalary = people.reduce((sum, person) => sum + person.annualSalary, 0);
      const socialRows = people.map((person) => ({person, si:personSocialInsurance(person, p)}));
      const employerSI = socialRows.reduce((sum, item) => sum + item.si.employer, 0);

      const companyTaxable = p.preProfit - totalAnnualSalary - employerSI;
      const corpTax = corporateTax(companyTaxable, p);
      const companyAfterTax = companyTaxable - corpTax;

      const effectivePayout = clamp(payoutRate, 0, 1);
      const totalDividend = companyAfterTax > 0 ? companyAfterTax * effectivePayout : 0;
      const retained = companyAfterTax - totalDividend;
      const allocated = allocateDividends(people, totalDividend);
      const calculatedPeople = allocated.map(({person, ownerDividend}) => calculatePersonTax(person, ownerDividend, p));

      const sum = (key) => calculatedPeople.reduce((total, person) => total + person[key], 0);
      const annualSalary = sum("annualSalary");
      const ownerDividend = sum("ownerDividend");
      const employeeSI = sum("employeeSI");
      const personalTax = sum("personalTax");
      const personalTakeHome = sum("personalTakeHome");
      const householdShare = Math.min(1, people.reduce((total, person) => total + person.share, 0));
      const ownerRetainedValue = retained * householdShare;
      const ownerTotal = personalTakeHome + ownerRetainedValue;
      const metric = p.objective === "personalCash" ? personalTakeHome : ownerTotal;
      const retainedGap = Math.abs(retained - p.targetRetained);

      const feasible =
        (!p.noLoss || companyAfterTax >= -1) &&
        retained >= p.minRetained - 1;

      const primary = calculatedPeople.find((person) => person.key === "primary") || calculatedPeople[0] || null;
      const spouse = calculatedPeople.find((person) => person.key === "spouse") || null;

      return {
        roleMode:p.roleMode,
        monthly:meta?.monthly ?? (primary?.monthly || 0),
        totalMonthly:meta?.totalMonthly ?? people.reduce((total, person) => total + person.monthly, 0),
        primaryMonthly:primary?.monthly || 0,
        spouseMonthly:spouse?.monthly || 0,
        annualSalary,
        primaryAnnualSalary:primary?.annualSalary || 0,
        spouseAnnualSalary:spouse?.annualSalary || 0,
        payoutRate:effectivePayout,
        totalDividend,
        ownerDividend,
        retained,
        householdShare,
        companyTaxable,
        corpTax,
        companyAfterTax,
        salIncome:sum("salIncome"),
        totalIncome:sum("totalIncome"),
        employeeSI,
        employerSI,
        healthStd:primary?.healthStd || 0,
        pensionStd:primary?.pensionStd || 0,
        taxableN:sum("taxableN"),
        taxableR:sum("taxableR"),
        nationalTax:sum("nationalTax"),
        residentTax:sum("residentTax"),
        personalTax,
        creditN:sum("creditN"),
        creditR:sum("creditR"),
        personalTakeHome,
        ownerRetainedValue,
        ownerTotal,
        metric,
        objective:p.objective,
        targetRetained:p.targetRetained,
        retainedGap,
        feasible,
        people:calculatedPeople
      };
    }

    function simulate(monthly, payoutRate, rawParams){
      const p = normalizeParams(rawParams);
      return simulatePeople([primaryPerson(monthly, p)], payoutRate, p, {
        monthly,
        totalMonthly:monthly
      });
    }

    function simulateCoupleSplit(primaryMonthly, payoutRate, rawParams){
      const p = normalizeParams({...rawParams, roleMode:"coupleSplit"});
      const totalMonthly = p.coupleTotalMonthly;
      const primaryMonthlyClamped = clamp(primaryMonthly, 0, totalMonthly);
      const spouseMonthly = Math.max(0, totalMonthly - primaryMonthlyClamped);
      return simulatePeople([
        primaryPerson(primaryMonthlyClamped, p),
        spousePerson(spouseMonthly, p)
      ], payoutRate, p, {
        monthly:primaryMonthlyClamped,
        totalMonthly
      });
    }

    function simulateCoupleGrid(primaryMonthly, spouseMonthly, payoutRate, rawParams){
      const p = normalizeParams({...rawParams, roleMode:"coupleGrid"});
      const primaryMonthlyClamped = clamp(primaryMonthly, 0, p.couplePrimaryMaxMonthly);
      const spouseMonthlyClamped = clamp(spouseMonthly, 0, p.coupleSpouseMaxMonthly);
      return simulatePeople([
        primaryPerson(primaryMonthlyClamped, p),
        spousePerson(spouseMonthlyClamped, p)
      ], payoutRate, p, {
        monthly:primaryMonthlyClamped,
        totalMonthly:primaryMonthlyClamped + spouseMonthlyClamped
      });
    }

    function payoutRatesFor(policy, p){
      if(!p.includeDividend) return [0];
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

      if(p.roleMode === "coupleSplit"){
        const totalMonthly = p.coupleTotalMonthly;
        const seen = new Set();
        const addPrimaryMonthly = (primaryMonthly) => {
          const normalizedMonthly = clamp(Math.round(primaryMonthly), 0, totalMonthly);
          if(seen.has(normalizedMonthly)) return;
          seen.add(normalizedMonthly);
          for(const rate of rates){
            const row = simulateCoupleSplit(normalizedMonthly, rate, p);
            if(row.feasible) rows.push(row);
          }
        };

        for(let monthly = 0; monthly <= totalMonthly + 0.1; monthly += p.step){
          addPrimaryMonthly(monthly);
        }
        addPrimaryMonthly(totalMonthly);

        rows.sort(compareRows);
        return rows;
      }

      if(p.roleMode === "coupleGrid"){
        const seen = new Set();
        const addPair = (primaryMonthly, spouseMonthly) => {
          const primary = clamp(Math.round(primaryMonthly), 0, p.couplePrimaryMaxMonthly);
          const spouse = clamp(Math.round(spouseMonthly), 0, p.coupleSpouseMaxMonthly);
          const key = `${primary}:${spouse}`;
          if(seen.has(key)) return;
          seen.add(key);
          for(const rate of rates){
            const row = simulateCoupleGrid(primary, spouse, rate, p);
            if(row.feasible) rows.push(row);
          }
        };

        for(let primary = 0; primary <= p.couplePrimaryMaxMonthly + 0.1; primary += p.step){
          for(let spouse = 0; spouse <= p.coupleSpouseMaxMonthly + 0.1; spouse += p.step){
            addPair(primary, spouse);
          }
        }
        addPair(p.couplePrimaryMaxMonthly, p.coupleSpouseMaxMonthly);

        rows.sort(compareRows);
        return rows;
      }

      const maxMonthly = Math.max(p.maxMonthly, p.currentMonthly);

      for(let monthly = 0; monthly <= maxMonthly + 0.1; monthly += p.step){
        for(const rate of rates){
          const row = simulate(monthly, rate, p);
          if(row.feasible) rows.push(row);
        }
      }

      rows.sort(compareRows);
      return rows;
    }

    return {socialInsurance, corporateTax, simulate, simulatePeople, simulateCoupleSplit, simulateCoupleGrid, payoutRatesFor, runSearch};
  }

  function readParams(documentRef){
    const doc = documentRef || document;
    const n = (id) => parseNumber(doc.getElementById(id).value);
    const v = (id) => doc.getElementById(id).value;
    const checked = (id) => doc.getElementById(id).checked;

    return normalizeParams({
      preProfit:n("preProfit"),
      roleMode:v("roleMode"),
      currentMonthly:n("currentMonthly"),
      maxMonthly:n("maxMonthly"),
      step:n("step"),
      share:n("shareRate") / 100,
      otherIncome:n("otherIncome"),
      coupleTotalMonthly:n("coupleTotalMonthly"),
      couplePrimaryMonthly:n("couplePrimaryMonthly"),
      coupleSpouseMonthly:n("coupleSpouseMonthly"),
      couplePrimaryMaxMonthly:n("couplePrimaryMaxMonthly"),
      coupleSpouseMaxMonthly:n("coupleSpouseMaxMonthly"),
      spouseShare:n("spouseShareRate") / 100,
      spouseOtherIncome:n("spouseOtherIncome"),
      objective:v("objective"),
      includeDividend:checked("includeDividend"),
      divPolicy:v("divPolicy"),
      fixedPayout:n("fixedPayout") / 100,
      targetRetained:n("targetRetained"),
      minRetained:n("minRetained"),
      noLoss:checked("noLoss"),
      applyDividendCredit:checked("applyDividendCredit"),
      taxYear:v("taxYear"),
      surtaxRate:n("surtaxRate") / 100,
      residentRate:n("residentRate") / 100,
      residentPerCapita:n("residentPerCapita"),
      otherDedN:n("otherDedN"),
      otherDedR:n("otherDedR"),
      spouseOtherDedN:n("spouseOtherDedN"),
      spouseOtherDedR:n("spouseOtherDedR"),
      healthRate:n("healthRate") / 100,
      careRate:n("careRate") / 100,
      supportRate:n("supportRate") / 100,
      pensionRate:n("pensionRate") / 100,
      childContributionRate:n("childContributionRate") / 100,
      careApplicable:checked("careApplicable"),
      socialApplicable:checked("socialApplicable"),
      spouseCareApplicable:checked("spouseCareApplicable"),
      spouseSocialApplicable:checked("spouseSocialApplicable"),
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
      if(rowIsBetter(row, old)) map.set(row.monthly, row);
    }
    return Array.from(map.values()).sort((a, b) => a.monthly - b.monthly);
  }

  function bestByMonthlyBucket(rows, bucketSize){
    const map = new Map();
    for(const row of rows){
      const bucket = Math.floor(row.monthly / bucketSize) * bucketSize;
      const old = map.get(bucket);
      if(rowIsBetter(row, old)){
        map.set(bucket, {...row, monthlyBucket:bucket, monthlyBucketSize:bucketSize});
      }
    }
    return Array.from(map.values()).sort(compareRows);
  }

  function monthlyBucketLabel(row){
    if(!Number.isFinite(row.monthlyBucket) || !Number.isFinite(row.monthlyBucketSize)){
      return row.roleMode === "coupleSplit" ? coupleSplitLabel(row) : yen(row.monthly);
    }
    const bucketStart = row.monthlyBucket;
    const bucketEnd = bucketStart + row.monthlyBucketSize - 1;
    const startMan = Math.floor(bucketStart / 10000).toLocaleString("ja-JP");
    const endMan = Math.floor(bucketEnd / 10000).toLocaleString("ja-JP");
    const bestMan = Math.round(row.monthly / 10000).toLocaleString("ja-JP");
    if(isCoupleRow(row)){
      const spouseMan = Math.round(row.spouseMonthly / 10000).toLocaleString("ja-JP");
      return `A月額${startMan}万〜${endMan}万（A${bestMan}万/B${spouseMan}万が最良）`;
    }
    return `月額${startMan}万〜${endMan}万（${bestMan}万が最良）`;
  }

  function isCoupleRow(row){
    return isCoupleModeValue(row?.roleMode);
  }

  function roleModeLabel(row){
    if(row?.roleMode === "coupleSplit") return "夫婦役員（合計固定）";
    if(row?.roleMode === "coupleGrid") return "夫婦役員（個別探索）";
    return "単独役員";
  }

  function coupleSplitLabel(row){
    return `A ${yen(row.primaryMonthly)} / B ${yen(row.spouseMonthly)}`;
  }

  function monthlyLabel(row){
    return isCoupleRow(row) ? coupleSplitLabel(row) : yen(row.monthly);
  }

  function totalSalaryLabel(row){
    return isCoupleRow(row) ? `合計 ${yen(row.totalMonthly)}` : yen(row.monthly);
  }

  function personalLabel(row){
    return isCoupleRow(row) ? "世帯手取り" : "個人手取り";
  }

  function ownerDividendLabel(row){
    return isCoupleRow(row) ? "世帯受取" : "本人受取";
  }

  function householdShareNote(row){
    if(!isCoupleRow(row)) return "";
    return `夫婦持分 ${(row.householdShare * 100).toLocaleString("ja-JP", {maximumFractionDigits:1})}%`;
  }

  function bestCurrentScenario(calculator, p){
    const rates = calculator.payoutRatesFor(p.divPolicy, p);
    return rates
      .map((rate) => p.roleMode === "coupleSplit"
        ? calculator.simulateCoupleSplit(p.couplePrimaryMonthly, rate, p)
        : p.roleMode === "coupleGrid"
          ? calculator.simulateCoupleGrid(p.couplePrimaryMonthly, p.coupleSpouseMonthly, rate, p)
        : calculator.simulate(p.currentMonthly, rate, p))
      .filter((row) => row.feasible)
      .sort(compareRows)[0] || null;
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

  function targetVisualData(target, actual, currentGap){
    const safeTarget = Math.max(0, Number(target) || 0);
    const actualValue = Number(actual) || 0;
    const chartActual = Math.max(0, actualValue);
    const difference = actualValue - safeTarget;
    const gap = Math.abs(difference);
    const safeCurrentGap = Number.isFinite(currentGap) ? Math.max(0, currentGap) : null;
    const amountScale = Math.max(safeTarget, chartActual, 1) * 1.08;
    const gapScale = Math.max(gap, safeCurrentGap || 0, 1);
    return {
      difference,
      gap,
      targetRatio:safeTarget > 0 ? actualValue / safeTarget : null,
      targetPosition:clamp(safeTarget / amountScale * 100, 0, 100),
      actualWidth:clamp(chartActual / amountScale * 100, 0, 100),
      currentGap:safeCurrentGap,
      currentGapWidth:safeCurrentGap === null ? null : clamp(safeCurrentGap / gapScale * 100, 0, 100),
      bestGapWidth:clamp(gap / gapScale * 100, 0, 100),
      improvement:safeCurrentGap === null ? null : safeCurrentGap - gap
    };
  }

  function renderReasonBreakdownChart(title, totalLabel, parts){
    const normalized = parts.map((part) => {
      const value = Number(part.value) || 0;
      return {...part, value, chartValue:Math.max(0, value)};
    });
    const total = normalized.reduce((sum, part) => sum + part.chartValue, 0) || 1;
    const ariaText = normalized.map((part) => `${part.label} ${man(part.value)}`).join("、");
    const segments = normalized.map((part) => {
      const width = part.chartValue / total * 100;
      return `<span class="reason-breakdown-segment" aria-hidden="true" style="width:${width}%;background:${part.color}"></span>`;
    }).join("");
    const legend = normalized.map((part) => `
      <div class="reason-breakdown-item${part.value < 0 ? " negative" : ""}">
        <span><i style="background:${part.color}" aria-hidden="true"></i>${part.label}</span>
        <strong>${man(part.value)}</strong>
      </div>
    `).join("");

    return `
      <section class="reason-visual-section">
        <div class="reason-visual-head"><span>${title}</span><strong>${totalLabel}</strong></div>
        <div class="reason-breakdown-track" role="img" aria-label="${title}。${ariaText}">${segments}</div>
        <div class="reason-breakdown-legend">${legend}</div>
      </section>
    `;
  }

  function renderTargetReasonVisual(best, current, p){
    const visual = targetVisualData(p.targetRetained, best.retained, current?.retainedGap);
    const differenceLabel = visual.difference < 0
      ? `不足 ${man(visual.gap)}`
      : visual.difference > 0
        ? `超過 ${man(visual.gap)}`
        : "目標と一致";
    const ratioLabel = visual.targetRatio === null
      ? "目標比 -"
      : `目標比 ${(visual.targetRatio * 100).toLocaleString("ja-JP", {maximumFractionDigits:2})}%`;
    const gapComparison = visual.currentGap === null ? "" : `
      <div class="reason-gap-chart" aria-label="比較月額の目標差額 ${man(visual.currentGap)}、おすすめの目標差額 ${man(visual.gap)}">
        <p class="reason-chart-caption">比較月額からどれだけ目標へ近づいたか</p>
        <div class="reason-gap-row">
          <span>比較月額</span>
          <i><b class="current" style="width:${visual.currentGapWidth}%"></b></i>
          <strong>${man(visual.currentGap)}</strong>
        </div>
        <div class="reason-gap-row">
          <span>おすすめ</span>
          <i><b class="best" style="width:${visual.bestGapWidth}%"></b></i>
          <strong>${man(visual.gap)}</strong>
        </div>
        <p class="reason-gap-result">目標差額を ${man(Math.max(0, visual.improvement))} 縮小</p>
      </div>
    `;

    return `
      <section class="reason-visual-section reason-target-section">
        <div class="reason-visual-head"><span>目標との距離</span><strong>${differenceLabel}</strong></div>
        <div class="reason-target-values">
          <span>おすすめ <strong>${man(best.retained)}</strong></span>
          <span>目標 <strong>${man(p.targetRetained)}</strong></span>
        </div>
        <div class="reason-target-track" role="img" aria-label="目標法人留保 ${man(p.targetRetained)} に対し、おすすめ法人留保 ${man(best.retained)}、${differenceLabel}">
          <span class="reason-target-fill" style="width:${visual.actualWidth}%"></span>
          <i class="reason-target-marker" style="left:${visual.targetPosition}%"><span>目標</span></i>
        </div>
        <div class="reason-target-summary"><strong>${differenceLabel}</strong><span>${ratioLabel}</span></div>
        ${gapComparison}
      </section>
    `;
  }

  function renderRecommendationReason(best, current, p){
    const personalName = personalLabel(best);
    const retainedName = isCoupleRow(best) ? "夫婦持分相当の法人留保" : "持分相当の法人留保";
    const retainedTargetMode = p.objective === "retainedTarget";
    const metricLine = retainedTargetMode
      ? `目標 ${man(p.targetRetained)} に対して法人留保 ${man(best.retained)}（差額 ${signedMan(best.retained - p.targetRetained)}）`
      : p.objective === "ownerTotal"
        ? `${personalName} ${man(best.personalTakeHome)} ＋ ${retainedName} ${man(best.ownerRetainedValue)} ＝ ${man(best.metric)}`
        : `${personalName} ${man(best.personalTakeHome)} ＝ ${man(best.metric)}`;
    const comparisonLine = current
      ? retainedTargetMode
        ? `比較案比：目標差額 ${man(current.retainedGap)} → ${man(best.retainedGap)}（改善 ${signedMan(current.retainedGap - best.retainedGap)}）、${personalName} ${signedMan(best.personalTakeHome - current.personalTakeHome)}`
        : `比較案比：評価指標 ${signedMan(best.metric - current.metric)}、${personalName} ${signedMan(best.personalTakeHome - current.personalTakeHome)}、法人留保 ${signedMan(best.retained - current.retained)}`
      : "比較月額は現在の条件では候補外です。";
    const personalEquation = (person) => `給与 ${man(person.annualSalary)}${p.includeDividend ? ` ＋ 配当 ${man(person.ownerDividend)}` : ""} − 個人税 ${man(person.personalTax)} − 本人社保 ${man(person.employeeSI)} ＝ ${man(person.personalTakeHome)}`;
    const personalDetail = isCoupleRow(best)
      ? best.people.map((person) => `${person.label}: ${personalEquation(person)}`).join("<br>")
      : personalEquation(best);
    const objectiveVisual = retainedTargetMode
      ? renderTargetReasonVisual(best, current, p)
      : p.objective === "ownerTotal"
        ? renderReasonBreakdownChart("評価指標の内訳", `合計 ${man(best.metric)}`, [
            {label:personalName, value:best.personalTakeHome, color:reasonColors.takeHome},
            {label:retainedName, value:best.ownerRetainedValue, color:reasonColors.retained}
          ])
        : "";
    const companyParts = [
      {label:"役員報酬", value:best.annualSalary, color:reasonColors.salary},
      {label:"会社社保", value:best.employerSI, color:reasonColors.social},
      {label:"法人税等", value:best.corpTax, color:reasonColors.tax},
      {label:"法人留保", value:best.retained, color:reasonColors.retained}
    ];
    if(p.includeDividend){
      companyParts.splice(3, 0, {label:"配当", value:best.totalDividend, color:reasonColors.dividend});
    }
    const companyVisual = renderReasonBreakdownChart("会社利益の行き先", `前利益 ${man(p.preProfit)}`, companyParts);
    const personalInflow = best.annualSalary + (p.includeDividend ? best.ownerDividend : 0);
    const personalVisual = renderReasonBreakdownChart(`${personalName}の行き先`, `${p.includeDividend ? "給与・配当" : "給与"} ${man(personalInflow)}`, [
      {label:personalName, value:best.personalTakeHome, color:reasonColors.takeHome},
      {label:"個人税", value:best.personalTax, color:reasonColors.tax},
      {label:"本人社保", value:best.employeeSI, color:reasonColors.social}
    ]);

    return `
      <div class="recommendation-reason">
        <p class="reason-title">なぜこの金額か</p>
        <p class="reason-lead">${retainedTargetMode ? "法人に残したい目標留保に最も近く、同程度なら手取りが大きい候補です。" : "探索条件を満たす候補のうち、評価指標が最も高い月額です。"}</p>
        <div class="reason-visuals">
          ${objectiveVisual}
          ${companyVisual}
          ${personalVisual}
        </div>
        <details class="reason-formulas">
          <summary>計算式の詳細</summary>
          <dl class="reason-list">
            <div>
              <dt>${retainedTargetMode ? "目標法人留保" : "評価指標"}</dt>
              <dd>${metricLine}</dd>
            </div>
            <div>
              <dt>${personalName}</dt>
              <dd>${personalDetail}</dd>
            </div>
            <div>
              <dt>法人側</dt>
              <dd>前利益 ${man(p.preProfit)} − 年間役員報酬 ${man(best.annualSalary)} − 会社社保 ${man(best.employerSI)} − 法人税等 ${man(best.corpTax)} ＝ 税引後利益 ${man(best.companyAfterTax)}<br>
              ${p.includeDividend ? `税引後利益 ${man(best.companyAfterTax)} − 配当総額 ${man(best.totalDividend)} ＝ 法人留保 ${man(best.retained)}` : `税引後利益 ${man(best.companyAfterTax)} ＝ 法人留保 ${man(best.retained)}`}</dd>
            </div>
            <div>
              <dt>比較</dt>
              <dd>${comparisonLine}</dd>
            </div>
          </dl>
        </details>
      </div>
    `;
  }

  function scenarioSummary(row, p){
    if(!row) return "候補なし";
    if(p.objective === "retainedTarget"){
      return `${monthlyLabel(row)}、法人留保 ${man(row.retained)}、目標差額 ${man(row.retainedGap)}`;
    }
    return `${monthlyLabel(row)}、評価指標 ${man(row.metric)}`;
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
    const coupleMode = isCoupleRow(best);
    const personalName = personalLabel(best);
    const retainedTargetMode = p.objective === "retainedTarget";
    const objectiveText = retainedTargetMode
      ? `法人留保 ${man(p.targetRetained)} から逆算`
      : p.objective === "ownerTotal"
      ? `${personalName}＋${coupleMode ? "夫婦" : ""}持分相当の法人留保`
      : (coupleMode ? "世帯キャッシュ" : "個人キャッシュ");
    let messageClass = "";
    let message = "税・社会保険の概算だけで見ると、この条件では上記の月額が最も有利です。";

    if(retainedTargetMode){
      message = `法人留保目標 ${man(p.targetRetained)} に対して、この候補の法人留保は ${man(best.retained)}（差額 ${signedMan(best.retained - p.targetRetained)}）です。`;
    }else if(p.objective === "ownerTotal" && best.payoutRate > 0.001 && best.householdShare >= .999){
      messageClass = "warn";
      message = coupleMode
        ? "夫婦で100%持分の総手残りを重視する場合、配当は法人留保を世帯へ移すたびに追加課税を受けます。個人資金化を重視する場面か確認してください。"
        : "100%オーナーで総手残りを重視する場合、配当は法人留保を個人へ移すたびに追加課税を受けます。個人資金化を重視する場面か確認してください。";
    }else if(p.objective === "personalCash" && best.totalDividend > 0){
      messageClass = "warn";
      message = `${coupleMode ? "世帯" : "個人"}キャッシュ最大では法人留保より${personalName}を優先するため、配当が多く出やすくなります。運転資金と金融機関評価を別途確認してください。`;
    }
    const shareNote = p.includeDividend && best.totalDividend > 0 && best.householdShare < 0.999
      ? `<div class="message">${coupleMode ? "夫婦持分" : "持株割合"} ${(best.householdShare * 100).toLocaleString("ja-JP", {maximumFractionDigits:1})}% のため、配当総額 ${man(best.totalDividend)} のうち${ownerDividendLabel(best)}は ${man(best.ownerDividend)} です。残りは他の株主に配当されます。</div>`
      : "";
    const shareOverNote = coupleMode && p.share + p.spouseShare > 1.000001
      ? '<div class="message warn">役員A/Bの持株割合合計が100%を超えています。計算上は世帯持分100%を上限に補正しています。入力値をご確認ください。</div>'
      : "";
    const resultLabel = coupleMode ? "おすすめ月額配分" : "おすすめ月額役員報酬";
    const resultAmount = coupleMode
      ? `<span>A ${yen(best.primaryMonthly)}</span><span>B ${yen(best.spouseMonthly)}</span>`
      : yen(best.monthly);
    const resultAmountClass = coupleMode ? "amount split-amount" : "amount";
    const payoutText = p.includeDividend ? `配当率 ${pct(best.payoutRate)}` : "役員報酬のみ";
    const resultSub = coupleMode
      ? `合計月額 ${yen(best.totalMonthly)} ／ 年間合計 ${yen(best.annualSalary)} ／ ${payoutText} ／ 評価基準 ${objectiveText}`
      : `年間 ${yen(best.annualSalary)} ／ ${payoutText} ／ 評価基準 ${objectiveText}`;
    const stdText = coupleMode
      ? `A 健保 ${man(best.people.find((person) => person.key === "primary")?.healthStd || 0)} / B 健保 ${man(best.people.find((person) => person.key === "spouse")?.healthStd || 0)}`
      : `厚年 ${man(best.pensionStd)}`;

    box.innerHTML = `
      <div class="hero-result">
        <div class="recommendation-main">
          <p class="label">${resultLabel}</p>
          <p class="${resultAmountClass}">${resultAmount}</p>
          <p class="sub">${resultSub}</p>
        </div>
        ${renderRecommendationReason(best, current, p)}
      </div>

      <div class="metric-grid">
        <div class="metric-card"><p class="k">${personalName}</p><p class="v">${man(best.personalTakeHome)}</p><p class="s">${p.includeDividend ? "給与＋配当" : "給与"}−個人税−本人社保</p></div>
        <div class="metric-card"><p class="k">法人留保</p><p class="v">${man(best.retained)}</p><p class="s">${p.includeDividend ? "法人税後利益−配当総額" : "法人税後利益"}</p></div>
        <div class="metric-card"><p class="k">税・社保合計</p><p class="v">${man(best.corpTax + best.personalTax + best.employeeSI + best.employerSI)}</p><p class="s">法人税、個人税、本人・会社社保</p></div>
        ${p.includeDividend ? `<div class="metric-card"><p class="k">配当総額</p><p class="v">${man(best.totalDividend)}</p><p class="s">${ownerDividendLabel(best)} ${man(best.ownerDividend)} ${householdShareNote(best)}</p></div>` : ""}
        <div class="metric-card"><p class="k">法人税等</p><p class="v">${man(best.corpTax)}</p><p class="s">課税所得 ${man(best.companyTaxable)}</p></div>
        <div class="metric-card"><p class="k">標準報酬月額</p><p class="v">${man(best.healthStd)}</p><p class="s">${stdText}</p></div>
      </div>

      <div class="message ${messageClass}">${message}</div>
      ${p.includeDividend
        ? `<div class="message">配当なし案：${scenarioSummary(noDivBest, p)}。 全額配当案：${scenarioSummary(allDivBest, p)}。</div>`
        : '<div class="message">役員報酬のみで探索しています。</div>'}
      ${shareOverNote}
      ${shareNote}
    `;
  }

  function balanceParts(row){
    return [
      {key:"personal", label:personalLabel(row), value:Math.max(0, row.personalTakeHome), color:balanceColors.personal},
      {key:"retained", label:"法人留保", value:Math.max(0, row.retained), color:balanceColors.retained},
      {key:"personalLoad", label:"個人税・本人社保", value:Math.max(0, row.personalTax + row.employeeSI), color:balanceColors.personalLoad},
      {key:"companyLoad", label:"法人税・会社社保", value:Math.max(0, row.corpTax + row.employerSI), color:balanceColors.companyLoad}
    ];
  }

  function renderBalanceRow(title, row, p){
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
        <div class="balance-head"><span>${title}</span><span>${monthlyLabel(row)} ／ ${p.includeDividend ? `配当率 ${pct(row.payoutRate)}` : "役員報酬のみ"}</span></div>
        <div class="balance-track">${segments}</div>
        <div class="balance-legend">${legend}</div>
      </div>
    `;
  }

  function renderBalance(best, current, p){
    const box = document.getElementById("balanceBox");
    box.innerHTML = `
      <div class="panel-head">
        <div>
          <p class="section-kicker">バランス</p>
          <h2>利益の配分イメージ</h2>
        </div>
      </div>
      ${renderBalanceRow("おすすめ案", best, p)}
      ${renderBalanceRow("比較月額", current, p)}
    `;
  }

  function renderDelta(best, current){
    const box = document.getElementById("deltaBox");
    if(!best || !current){
      box.innerHTML = "";
      return;
    }

    const retainedTargetMode = best.objective === "retainedTarget";
    const deltas = [
      retainedTargetMode
        ? {label:"目標差額の改善", value:current.retainedGap - best.retainedGap}
        : {label:"評価指標の差", value:best.metric - current.metric},
      {label:`${personalLabel(best)}の差`, value:best.personalTakeHome - current.personalTakeHome},
      {label:"法人留保の差", value:best.retained - current.retained}
    ];

    box.innerHTML = deltas.map((item) => `
      <div class="delta-card ${classifyDelta(item.value)}">
        <p class="k">${item.label}</p>
        <p class="v">${item.value >= 0 ? "+" : ""}${man(item.value)}</p>
      </div>
    `).join("");
  }

  function renderTable(rows, p){
    latestRows = rows;
    const box = document.getElementById("tableBox");
    if(!rows.length){
      box.innerHTML = '<div class="message bad">上位候補を表示できません。</div>';
      return;
    }

    const top = bestByMonthlyBucket(rows, 100000).slice(0, 12);
    const coupleMode = isCoupleRow(rows[0]);
    const retainedTargetMode = rows[0]?.objective === "retainedTarget";
    box.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>順位</th>
            <th>${coupleMode ? "配分" : "月額"}</th>
            ${p.includeDividend ? "<th>配当率</th>" : ""}
            <th>${coupleMode ? "世帯手取り" : "個人手取り"}</th>
            <th>法人留保</th>
            ${retainedTargetMode ? "<th>目標差額</th>" : ""}
            <th>総手残り</th>
            <th>税・社保</th>
          </tr>
        </thead>
        <tbody>
          ${top.map((row, index) => `
            <tr>
              <td>${index + 1}</td>
              <td>${monthlyBucketLabel(row)}</td>
              ${p.includeDividend ? `<td>${pct(row.payoutRate)}</td>` : ""}
              <td>${man(row.personalTakeHome)}</td>
              <td>${man(row.retained)}</td>
              ${retainedTargetMode ? `<td>${man(row.retainedGap)}</td>` : ""}
              <td>${man(row.ownerTotal)}</td>
              <td>${man(row.corpTax + row.personalTax + row.employeeSI + row.employerSI)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function drawChart(rows, current, p){
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
    if(p.targetRetained > 0) values.push(p.targetRetained);
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

    if(p.targetRetained > 0){
      const targetY = yScale(p.targetRetained);
      ctx.strokeStyle = "#b45309";
      ctx.lineWidth = 2;
      ctx.setLineDash([6, 6]);
      ctx.beginPath();
      ctx.moveTo(pad.l, targetY);
      ctx.lineTo(pad.l + w, targetY);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = "#b45309";
      ctx.font = "12px sans-serif";
      ctx.fillText("目標留保", pad.l + 8, targetY - 8);
    }

    const best = rows[0];
    const bestX = xScale(best.monthly);
    const bestY = yScale(best.objective === "retainedTarget" ? best.retained : best.metric);
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
    ctx.fillText(isCoupleRow(rows[0]) ? "役員A月額" : "月額役員報酬", pad.l + w - 82, canvas.height - 18);
    ctx.fillText("金額", pad.l, 18);
  }

  function buildCsv(rows){
    const headers = [
      "順位","役員構成","月額役員報酬","役員A月額","役員B月額","年間役員報酬","配当率","配当総額","本人・世帯配当",
      "個人手取り","法人留保","目標法人留保","目標差額","オーナー総手残り","法人税等","法人課税所得",
      "本人社保","会社社保","個人税","所得税等","住民税"
    ];
    const lines = [headers.join(",")];
    const csvCell = (value) => {
      if(typeof value === "number") return String(Math.round(value * 100) / 100);
      return `"${String(value ?? "").replace(/"/g, '""')}"`;
    };
    rows.forEach((row, index) => {
      lines.push([
        index + 1,
        roleModeLabel(row),
        row.monthly,
        row.primaryMonthly,
        row.spouseMonthly,
        row.annualSalary,
        row.payoutRate,
        row.totalDividend,
        row.ownerDividend,
        row.personalTakeHome,
        row.retained,
        row.targetRetained,
        row.retainedGap,
        row.ownerTotal,
        row.corpTax,
        row.companyTaxable,
        row.employeeSI,
        row.employerSI,
        row.personalTax,
        row.nationalTax,
        row.residentTax
      ].map(csvCell).join(","));
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

  function getPrefectureRate(name){
    return prefectureRates.find((item) => item.name === name) || null;
  }

  function applyPrefectureRate(name){
    if(name === "custom") return;
    const rate = getPrefectureRate(name);
    const healthInput = document.getElementById("healthRate");
    if(!rate || !healthInput) return;
    healthInput.value = formatInput("healthRate", rate.healthRate);
  }

  function markPrefectureCustom(){
    const prefectureSelect = document.getElementById("prefecture");
    if(prefectureSelect && prefectureSelect.value !== "custom"){
      prefectureSelect.value = "custom";
    }
  }

  function applyDefaults(){
    Object.entries(defaults).forEach(([id, value]) => setControlValue(id, value));
  }

  function applyStrategyPreset(value){
    const preset = strategyPresets[value];
    if(!preset) return;
    Object.entries(preset).forEach(([id, presetValue]) => setControlValue(id, presetValue));
    formatAllInputs();
    updateDependentControls();
  }

  function markStrategyCustom(){
    const strategySelect = document.getElementById("strategyPreset");
    if(strategySelect && strategySelect.value !== "custom"){
      strategySelect.value = "custom";
    }
  }

  function formatAllInputs(){
    [...moneyIds, ...percentIds].forEach((id) => {
      const el = document.getElementById(id);
      if(el) el.value = formatInput(id, parseNumber(el.value));
    });
  }

  function updateRoleModeControls(){
    const roleMode = document.getElementById("roleMode")?.value || "single";
    const coupleMode = isCoupleModeValue(roleMode);
    const splitMode = roleMode === "coupleSplit";
    const gridMode = roleMode === "coupleGrid";
    document.querySelectorAll(".couple-only").forEach((el) => {
      el.hidden = !coupleMode;
    });
    document.querySelectorAll(".couple-split-only").forEach((el) => {
      el.hidden = !splitMode;
    });
    document.querySelectorAll(".couple-grid-only").forEach((el) => {
      el.hidden = !gridMode;
    });
    document.querySelectorAll(".single-only").forEach((el) => {
      el.hidden = coupleMode;
    });

    const totalInput = document.getElementById("coupleTotalMonthly");
    const primaryInput = document.getElementById("couplePrimaryMonthly");
    const spouseInput = document.getElementById("coupleSpouseMonthly");
    const primaryMaxInput = document.getElementById("couplePrimaryMaxMonthly");
    const spouseMaxInput = document.getElementById("coupleSpouseMaxMonthly");
    const stepLabel = document.getElementById("stepLabel");
    const stepNote = document.getElementById("stepNote");
    if(stepLabel){
      stepLabel.textContent = splitMode
        ? "探索刻み（役員A月額の刻み幅）"
        : gridMode
          ? "探索刻み（A/B月額の刻み幅）"
          : "探索刻み";
    }
    if(stepNote){
      stepNote.textContent = gridMode
        ? "A/Bそれぞれで探索するため、候補数が増えます。まず5万〜10万円刻みがおすすめです。"
        : "";
      stepNote.hidden = !gridMode;
    }
    if(totalInput && primaryInput){
      const total = Math.max(0, parseNumber(totalInput.value));
      const primaryLimit = splitMode
        ? total
        : Math.max(1000, parseNumber(primaryMaxInput?.value || 0));
      if(parseNumber(primaryInput.value) > primaryLimit){
        primaryInput.value = formatInput("couplePrimaryMonthly", primaryLimit);
      }
    }
    if(spouseInput && spouseMaxInput){
      const spouseLimit = Math.max(1000, parseNumber(spouseMaxInput.value));
      if(parseNumber(spouseInput.value) > spouseLimit){
        spouseInput.value = formatInput("coupleSpouseMonthly", spouseLimit);
      }
    }
  }

  function updateDependentControls(){
    const includeDividend = document.getElementById("includeDividend")?.checked === true;
    const policy = document.getElementById("divPolicy")?.value;
    const objective = document.getElementById("objective")?.value;
    const fixedInput = document.getElementById("fixedPayout");
    const targetField = document.querySelector(".target-retained-field");
    const disabled = !includeDividend || policy !== "fixed";
    if(fixedInput) fixedInput.disabled = disabled;
    document.querySelectorAll(".dividend-option-field").forEach((el) => {
      el.hidden = !includeDividend;
    });
    if(targetField) targetField.hidden = objective !== "retainedTarget";
    updateRoleModeControls();
  }

  function update(){
    const calculator = createCalculator();
    const p = readParams(document);
    const rows = calculator.runSearch(p.divPolicy, p);
    const noRows = calculator.runSearch("none", p);
    const allRows = p.includeDividend ? calculator.runSearch("all", p) : [];
    const current = bestCurrentScenario(calculator, p);
    const best = rows[0] || null;

    latestState = {rows, best, current, params:p};
    renderResult(best, current, noRows[0] || null, allRows[0] || null, p);
    renderBalance(best, current, p);
    renderDelta(best, current);
    renderTable(rows, p);
    drawChart(rows, current, p);
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
      if(id === "healthRate"){
        el.addEventListener("input", markPrefectureCustom);
        el.addEventListener("change", markPrefectureCustom);
      }
      if(["roleMode","coupleTotalMonthly","couplePrimaryMonthly","coupleSpouseMonthly","couplePrimaryMaxMonthly","coupleSpouseMaxMonthly"].includes(id)){
        el.addEventListener("input", updateRoleModeControls);
        el.addEventListener("change", updateRoleModeControls);
      }
      if(id === "includeDividend"){
        el.addEventListener("input", updateDependentControls);
        el.addEventListener("change", updateDependentControls);
      }
      if(moneyIds.has(id) || percentIds.has(id)){
        el.addEventListener("input", () => {
          if(moneyIds.has(id)) formatMoneyLive(el);
        });
        el.addEventListener("blur", () => {
          el.value = formatInput(id, parseNumber(el.value));
          queueUpdate();
        });
      }
    });

    const prefectureSelect = document.getElementById("prefecture");
    if(prefectureSelect){
      prefectureSelect.addEventListener("change", () => {
        applyPrefectureRate(prefectureSelect.value);
        queueUpdate();
      });
    }

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

    document.getElementById("printPage").addEventListener("click", () => {
      window.print();
    });
  }

  function bindManualDialog(){
    const dialog = document.getElementById("manualDialog");
    const openButton = document.getElementById("openManual");
    const closeButton = document.getElementById("closeManual");
    if(!dialog || !openButton || !closeButton) return;

    const close = () => {
      if(typeof dialog.close === "function") dialog.close();
      else dialog.removeAttribute("open");
    };

    openButton.addEventListener("click", () => {
      if(typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute("open", "");
    });
    closeButton.addEventListener("click", close);
    dialog.addEventListener("click", (event) => {
      if(event.target === dialog) close();
    });
  }

  function boot(){
    formatAllInputs();
    bindControls();
    bindManualDialog();
    updateDependentControls();
    document.getElementById("divPolicy").addEventListener("change", updateDependentControls);
    document.getElementById("objective").addEventListener("change", updateDependentControls);
    update();
  }

  return {
    VERSION,
    boot,
    defaults,
    prefectureRates,
    getPrefectureRate,
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
    buildCsv,
    targetVisualData
  };
});
