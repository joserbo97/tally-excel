// Hard tier: tasks that need planning, large data, or careful reasoning — where smaller local models
// tend to fall behind frontier models. Same task shape as tasks.mjs.
import { check, near, mentions, mentionsNumber, findLabel, findColumn, rowNumbersRight, dynamicCheck, noNewErrors, unchanged } from "./checks.mjs";

function rng(seed) {
  return () => {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const round2 = (x) => Math.round(x * 100) / 100;
const pmt = (P, annual, years) => { const r = annual / 12, n = years * 12; return (P * r) / (1 - (1 + r) ** -n); };

// ---------- data ----------
const CATS = ["Travel", "Meals", "Software", "Office"];
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
export const TX200 = (() => {
  const r = rng(99);
  const days = [31, 28, 31, 30, 31, 30];
  return Array.from({ length: 200 }, () => {
    const m = Math.floor(r() * 6);
    // include month-boundary dates (1st and last day) on purpose
    const d = [1, days[m], 1 + Math.floor(r() * days[m])][Math.floor(r() * 3)];
    return [`2026-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`, CATS[Math.floor(r() * 4)], round2(10 + r() * 590)];
  });
})();
const monthCat = (m, c) => round2(TX200.filter((t) => Number(t[0].slice(5, 7)) === m + 1 && t[1] === c).reduce((a, t) => a + t[2], 0));

const LEDGER = (() => {
  const r = rng(5);
  return Array.from({ length: 60 }, (_, i) => [`T${String(1001 + i)}`, `2026-04-${String(1 + (i % 30)).padStart(2, "0")}`, round2(50 + r() * 4950)]);
})();
// Bank = ledger with 4 planted discrepancies.
const BANK = (() => {
  const b = LEDGER.map((x) => [...x]);
  b[11][2] = round2(b[11][2] + 90); // T1012 amount mismatch
  b.splice(27, 1); // T1028 missing from bank
  b.push(["T1061", "2026-04-30", 312.4]); // in bank, not in ledger
  b.splice(40, 0, [...b[39]]); // duplicate of b[39] in bank
  return b;
})();
const DUP_ID = LEDGER.filter((x) => x[0] !== "T1028")[39][0];

const PRICES = { Coffee: 12.5, Tea: 8, Cocoa: 10.25, Mate: 6.75, Chai: 9.5, Matcha: 18, Rooibos: 7.25, "Yerba Blend": 11 };
const SALES3K = (() => {
  const r = rng(2026);
  const prods = Object.keys(PRICES);
  return Array.from({ length: 3000 }, () => {
    const m = 1 + Math.floor(r() * 3);
    const p = prods[Math.floor(r() * prods.length)];
    return [`2026-${String(m).padStart(2, "0")}-${String(1 + Math.floor(r() * 28)).padStart(2, "0")}`, p, 1 + Math.floor(r() * 20), PRICES[p]];
  });
})();
const marchTop3 = (() => {
  const t = {};
  for (const [d, p, q, pr] of SALES3K) if (d.slice(5, 7) === "03") t[p] = (t[p] || 0) + q * pr;
  return Object.entries(t).sort((a, b) => b[1] - a[1]).slice(0, 3);
})();

// Messy Bolivian-style export: mixed date formats and "Bs 1.234,50" amounts (decimal comma).
const MESSY = [
  ["15/03/2026", "Bs 1.234,50"], ["2026-3-16", "Bs 980,00"], ["March 17, 2026", "Bs 12.500,75"],
  ["18/03/2026", "Bs 45,9"], ["2026-03-19", "Bs 3.000"], ["20 Mar 2026", "Bs 0,99"],
];
const MESSY_NUM = [1234.5, 980, 12500.75, 45.9, 3000, 0.99];
const serial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;

// Column whose header matches re and has the most numeric cells below it.
function bestColumn(api, sheet, re) {
  let best = null;
  for (const lab of findLabel(api, re, sheet)) {
    const below = api.cells(sheet).filter((c) => c.c === lab.c && c.r > lab.r && typeof c.value === "number").length;
    if (!best || below > best.n) best = { sheet, c: lab.c, headerRow: lab.r, n: below };
  }
  return best;
}

// ---------- tasks ----------
export const HARD_TASKS = [
  {
    id: "hard-amortization-schedule", category: "hard", level: 3,
    workbook: { sheets: { Loan: [] } },
    prompt: "Build a full monthly amortization schedule for a 180,000 loan at 7.2% annual interest over 15 years: inputs in their own cells, then one row per month with payment, interest, principal and remaining balance. Also show total interest paid.",
    check: (api) => {
      const want = pmt(180000, 0.072, 15);
      const cells = api.cells("Loan");
      const balances = bestColumn(api, "Loan", /balance/i);
      const balCells = balances ? cells.filter((c) => c.c === balances.c && c.r > balances.headerRow && typeof c.value === "number") : [];
      const last = balCells[balCells.length - 1];
      // The schedule's Interest header sits on the same row as its Balance header (not "Total interest" up in the inputs).
      const intLab = balances && findLabel(api, /interest/i, "Loan").find((l) => l.r === balances.headerRow);
      const interestCol = intLab && { sheet: "Loan", c: intLab.c, headerRow: intLab.r };
      const firstInterest = interestCol && cells.find((c) => c.c === interestCol.c && c.r > interestCol.headerRow && typeof c.value === "number" && c.formula);
      const totalInt = cells.find((c) => c.formula && near(c.value, want * 180 - 180000, 0.001));
      const rate = cells.find((c) => !c.formula && (c.value === 0.072 || c.value === 7.2));
      return [
        check("180 monthly rows", balCells.length >= 180 && balCells.length <= 182, `${balCells.length} balance rows`),
        check("final balance ≈ 0", last && Math.abs(last.value) < 1, last?.value),
        check("month 1 interest = 1,080", firstInterest && near(firstInterest.value, 1080), firstInterest?.value),
        check(`total interest ≈ ${Math.round(want * 180 - 180000).toLocaleString("en-US")}`, totalInt, totalInt?.value),
        rate && firstInterest ? dynamicCheck("schedule is live (rate → 6%)", api, `Loan!${rate.ref}`, rate.value === 7.2 ? 6 : 0.06, () => api.value(`Loan!${firstInterest.ref}`), () => 900) : check("schedule is live", false, "cells not found"),
        rate && last ? dynamicCheck("still pays off at 6%", api, `Loan!${rate.ref}`, rate.value === 7.2 ? 6 : 0.06, () => Math.round(api.value(`Loan!${last.ref}`)), () => 0) : check("still pays off", false, "cells not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "hard-month-category-grid", category: "hard", level: 3,
    workbook: { sheets: { Transactions: [["Date", "Category", "Amount"], ...TX200.map((t) => [...t])] } },
    prompt: "Create a sheet called Monthly with a grid of total amount by month (Jan–Jun 2026, as rows) and category (as columns), plus a total column and a total row. Use formulas based on the Transactions sheet dates.",
    check: (api) => {
      const out = [check("Monthly sheet exists", api.hasSheet("Monthly"))];
      if (!api.hasSheet("Monthly")) return out;
      let ok = 0, any = null;
      for (let m = 0; m < 6; m++) {
        // Month labels may be text ("Jan") or real dates (a serial for the 1st, shown as e.g. yyyy-mm).
        const dateCell = api.cells("Monthly").find((c) => c.c === 0 && typeof c.value === "number" && near(c.value, serial(2026, m + 1, 1), 0.0001));
        const lab = findLabel(api, new RegExp(`^(${MONTHS[m]}|${["January", "February", "March", "April", "May", "June"][m]}|2026-0${m + 1}|0?${m + 1}/2026)`, "i"), "Monthly")[0]
          || (dateCell && { ...dateCell, sheet: "Monthly" });
        const nums = lab ? rowNumbersRight(api, lab).map((c) => c) : [];
        const hits = CATS.filter((c) => nums.some((n) => near(n.value, monthCat(m, c), 0.0001))).length;
        if (hits === 4) ok++;
        if (m === 0) any = nums.find((n) => near(n.value, monthCat(0, "Travel"), 0.0001));
      }
      out.push(check("all 24 month×category totals exact (incl. month-boundary dates)", ok === 6, `${ok}/6 months fully correct`));
      const grand = api.cells("Monthly").find((c) => c.formula && near(c.value, round2(TX200.reduce((a, t) => a + t[2], 0)), 0.0001));
      out.push(check("grand total present", grand));
      out.push(check("grid uses formulas", any?.formula));
      const firstJanTravel = TX200.findIndex((t) => t[0].startsWith("2026-01") && t[1] === "Travel");
      out.push(any ? dynamicCheck("grid is live", api, `Transactions!C${firstJanTravel + 2}`, TX200[firstJanTravel][2] + 1000, () => api.value(`Monthly!${any.ref}`), () => monthCat(0, "Travel") + 1000) : check("grid is live", false, "Jan/Travel cell not found"));
      out.push(noNewErrors(api));
      return out;
    },
  },
  {
    id: "hard-reconciliation", category: "hard", level: 3,
    workbook: {
      active: "Ledger",
      sheets: { Ledger: [["Ref", "Date", "Amount"], ...LEDGER.map((x) => [...x])], Bank: [["Ref", "Date", "Amount"], ...BANK.map((x) => [...x])] },
    },
    prompt: "Reconcile the Ledger sheet against the Bank sheet (match on Ref). In a new sheet called Reconciliation, list every discrepancy: refs missing from either side, amount mismatches (with both amounts), and duplicates. Then tell me how many discrepancies you found.",
    check: (api, { answer }) => {
      const out = [check("Reconciliation sheet exists", api.hasSheet("Reconciliation"))];
      // Helper columns that list (nearly) every ref — e.g. a UNIQUE(VSTACK(...)) key column next to COUNTIF flags —
      // are working space, not claimed discrepancies. Score only the other columns.
      const cells = api.hasSheet("Reconciliation") ? api.cells("Reconciliation") : [];
      const refsByCol = new Map();
      for (const c of cells) if (/^T1\d{3}$/.test(String(c.value))) refsByCol.set(c.c, (refsByCol.get(c.c) || new Set()).add(c.value));
      const helperCols = new Set([...refsByCol].filter(([, s]) => s.size > 30).map(([col]) => col));
      const text = cells.filter((c) => !helperCols.has(c.c)).map((c) => String(c.value)).join(" | ");
      out.push(check("finds T1012 amount mismatch", /T1012/.test(text)));
      out.push(check("finds T1028 missing from Bank", /T1028/.test(text)));
      out.push(check("finds T1061 missing from Ledger", /T1061/.test(text)));
      out.push(check(`finds duplicate ${DUP_ID}`, new RegExp(DUP_ID).test(text)));
      const extra = (text.match(/T1\d{3}/g) || []).filter((r) => ![DUP_ID, "T1012", "T1028", "T1061"].includes(r));
      out.push(check("no false discrepancies", new Set(extra).size === 0, [...new Set(extra)].slice(0, 6).join(",")));
      out.push(check("reports 4 discrepancies", mentionsNumber(answer, 4) || /\bfour\b/i.test(answer), answer.slice(0, 200)));
      return out;
    },
  },
  {
    id: "hard-npv-irr", category: "hard", level: 3,
    workbook: {
      sheets: {
        Projects: [["Year", "Project A", "Project B"], [0, -100000, -100000], [1, 20000, 60000], [2, 30000, 40000], [3, 40000, 20000], [4, 50000, 10000], [5, 60000, 5000]],
      },
    },
    prompt: "Using a 10% discount rate in its own input cell, calculate the NPV and IRR of both projects with formulas, and tell me which project to choose and why.",
    check: (api, { answer }) => {
      const cf = { A: [-100000, 20000, 30000, 40000, 50000, 60000], B: [-100000, 60000, 40000, 20000, 10000, 5000] };
      const npv = (r, c) => c.reduce((a, v, t) => a + v / (1 + r) ** t, 0);
      const irr = (c) => { let lo = -0.9, hi = 1; for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; npv(m, c) > 0 ? (lo = m) : (hi = m); } return (lo + hi) / 2; };
      const cells = api.cells("Projects");
      const nA = cells.find((c) => c.formula && near(c.value, npv(0.1, cf.A), 0.001));
      const nB = cells.find((c) => c.formula && near(c.value, npv(0.1, cf.B), 0.001));
      const iA = cells.find((c) => c.formula && near(c.value, irr(cf.A), 0.002));
      const iB = cells.find((c) => c.formula && near(c.value, irr(cf.B), 0.002));
      const rate = cells.find((c) => !c.formula && (c.value === 0.1 || c.value === 10) && c.c !== 0);
      return [
        check(`NPV A ≈ ${Math.round(npv(0.1, cf.A)).toLocaleString("en-US")} (Year 0 not discounted)`, nA, nA?.value),
        check(`NPV B ≈ ${Math.round(npv(0.1, cf.B)).toLocaleString("en-US")}`, nB, nB?.value),
        check("IRR A and B correct", iA && iB, `${iA?.value} / ${iB?.value}`),
        check("recommends Project A (higher NPV)", /project a\b|\bA\b[^.]{0,60}(choose|recommend|higher npv|better)|(choose|recommend|go with)[^.]{0,30}\bA\b/i.test(answer) && !/(choose|recommend|go with)[^.]{0,30}project b\b/i.test(answer), answer.slice(0, 250)),
        rate && nA ? dynamicCheck("NPV is live (rate → 15%)", api, `Projects!${rate.ref}`, rate.value === 10 ? 15 : 0.15, () => api.value(`Projects!${nA.ref}`), () => npv(0.15, cf.A)) : check("NPV is live", false, "rate input / NPV cell not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "hard-scenario-switch", category: "hard", level: 3,
    workbook: { sheets: { Model: [] } },
    prompt: "Build a scenario model: three scenarios (Worst, Base, Best) with units sold 8,000 / 10,000 / 13,000, price 22 / 25 / 27 and variable cost per unit 15 / 14 / 13; fixed costs 60,000 in all. Add a selector cell where I type 1, 2 or 3 to choose the scenario, and a P&L (revenue, variable costs, fixed costs, profit) that follows the selected scenario.",
    check: (api) => {
      const profit = (u, p, v) => u * (p - v) - 60000;
      const P = [profit(8000, 22, 15), profit(10000, 25, 14), profit(13000, 27, 13)];
      const cells = api.cells("Model");
      // The selector is the plain 1/2/3 input that, when set to s, makes some formula cell equal P[s-1].
      let sel = null, prof = null;
      for (const cand of cells.filter((c) => !c.formula && [1, 2, 3].includes(c.value))) {
        const orig = cand.value;
        const found = [1, 2, 3].map((s) => {
          api.setValue(`Model!${cand.ref}`, s);
          return new Set(api.cells("Model").filter((c) => c.formula && near(c.value, P[s - 1])).map((c) => c.ref));
        });
        api.setValue(`Model!${cand.ref}`, orig);
        const common = [...found[0]].filter((r) => found[1].has(r) && found[2].has(r));
        if (common.length) { sel = cand; prof = cells.find((c) => c.ref === common[0]); break; }
      }
      const out = [
        check("selector cell drives the model", sel, sel?.ref || "no 1/2/3 input switches profit between the three scenarios"),
        check(`profit for the selected scenario (${sel ? ["Worst", "Base", "Best"][sel.value - 1] : "?"})`, prof && near(prof.value, P[sel.value - 1]), prof?.value),
      ];
      if (sel && prof) {
        for (const s of [1, 2, 3]) out.push(dynamicCheck(`selector = ${s} → profit ${P[s - 1].toLocaleString("en-US")}`, api, `Model!${sel.ref}`, s, () => api.value(`Model!${prof.ref}`), () => P[s - 1]));
      }
      out.push(noNewErrors(api));
      return out;
    },
  },
  {
    id: "hard-large-data-question", category: "hard", level: 3,
    workbook: { sheets: { Sales: [["Date", "Product", "Qty", "Unit Price"], ...SALES3K.map((r) => [...r])] } },
    prompt: "Which 3 products had the highest revenue (Qty × Unit Price) in March 2026, and what was each one's March revenue? The sheet has 3,000 rows. You may add helper formulas or a helper sheet if useful.",
    check: (api, { answer }) => [
      check(`top 3 = ${marchTop3.map((x) => x[0]).join(", ")} in order`, (() => {
        const idx = marchTop3.map(([p]) => answer.search(new RegExp(`\\b${p}\\b`, "i")));
        return idx.every((i) => i >= 0) && idx[0] < idx[1] && idx[1] < idx[2];
      })(), answer.slice(0, 300)),
      ...marchTop3.map(([p, v]) => check(`${p} March revenue ≈ ${Math.round(v).toLocaleString("en-US")}`, mentionsNumber(answer, v, 0.002))),
    ],
  },
  {
    id: "hard-messy-bolivian-data", category: "hard", level: 3,
    workbook: { sheets: { Export: [["Fecha", "Monto"], ...MESSY.map((r) => [`'${r[0]}`, `'${r[1]}`])] } },
    prompt: "This export from a Bolivian bank has dates in mixed formats (day/month/year when written with slashes) and amounts like \"Bs 1.234,50\" (dot = thousands, comma = decimals). Convert the Fecha column to real Excel dates and the Monto column to real numbers, then add a total of the amounts.",
    check: (api) => {
      const dates = [2, 3, 4, 5, 6, 7].map((r) => api.value(`Export!A${r}`));
      const nums = api.cells("Export").filter((c) => c.c === 1 && c.r >= 1 && c.r <= 6).map((c) => c.value);
      const want = [15, 16, 17, 18, 19, 20].map((d) => serial(2026, 3, d));
      const total = api.cells("Export").find((c) => c.formula && near(c.value, MESSY_NUM.reduce((a, b) => a + b, 0)));
      return [
        check("all 6 dates are real dates for 15–20 March 2026", dates.every((d, i) => d === want[i]), dates.join(",")),
        check("amounts parsed with decimal comma (1234.5, 980, 12500.75, 45.9, 3000, 0.99)", nums.length === 6 && nums.every((n, i) => near(n, MESSY_NUM[i], 0.0001)), nums.join(",")),
        check("total = 17,762.14 by formula", total, total?.value),
      ];
    },
  },
];
