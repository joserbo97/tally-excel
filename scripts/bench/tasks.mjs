// Benchmark tasks: realistic Excel requests with automatic scoring.
// Each task: { id, category, level (1-3), workbook, attachments?, prompt, check(api, ctx) -> checks[] }
// ctx = { answer, before (snapshot), baselineErrors, fixtures }
import { check, near, mentions, mentionsNumber, findLabel, findColumn, columnCells, valueAt, rowNumbersRight, dynamicCheck, noNewErrors, unchanged } from "./checks.mjs";

// ---------- deterministic data ----------
function rng(seed) {
  return () => {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const round2 = (x) => Math.round(x * 100) / 100;

const REGIONS = [
  ["North", 120000, 135000, 150000, 160000],
  ["South", 98000, 112000, 125000, 138000],
  ["East", 156000, 149000, 172000, 185000],
  ["West", 87000, 95000, 108000, 121000],
  ["Central", 140000, 138000, 145000, 150000],
  ["Andes", 60000, 75000, 90000, 130000],
];
const salesSheet = () => [["Region", "Q1", "Q2", "Q3", "Q4"], ...REGIONS.map((r) => [...r])];
const regionTotal = (name) => REGIONS.find((r) => r[0] === name).slice(1).reduce((a, b) => a + b, 0);

const PRODUCTS = [["Coffee", 40000, 46000], ["Tea", 25000, 31000], ["Cocoa", 18000, 21600], ["Mate", 9000, 12600], ["Chai", 12000, 13200]];
const productsSheet = () => [["Product", "2024", "2025"], ...PRODUCTS.map((p) => [...p])];

const CATS = ["Travel", "Meals", "Software", "Office"];
const TX = (() => {
  const r = rng(42);
  return Array.from({ length: 40 }, (_, i) => {
    const d = new Date(Date.UTC(2026, 0, 1 + Math.floor(r() * 90)));
    return [d.toISOString().slice(0, 10), CATS[Math.floor(r() * CATS.length)], round2(20 + r() * 980)];
  });
})();
const txSheet = () => [["Date", "Category", "Amount"], ...TX.map((t) => [...t])];

const REPS = [["Ana", 42000], ["Bruno", 50000], ["Carla", 76500], ["Diego", 100000], ["Elena", 100001], ["Fabio", 131250], ["Gaby", 18900], ["Hugo", 64000]];
const commission = (rev) => rev * (rev < 50000 ? 0.02 : rev <= 100000 ? 0.03 : 0.05);

// ---------- tasks ----------
export const TASKS = [
  // ===== analysis (answer only, sheet must stay unchanged) =====
  {
    id: "analysis-top-region", category: "analysis", level: 1,
    workbook: { sheets: { Sales: salesSheet() } },
    prompt: "Which region had the highest total sales for the full year, and what was that total? Don't change the sheet.",
    check: (api, { answer, before }) => [
      check("names East", mentions(answer, /\bEast\b/), answer.slice(0, 200)),
      check("states total 662,000", mentionsNumber(answer, 662000)),
      unchanged(api, before),
    ],
  },
  {
    id: "analysis-category-average", category: "analysis", level: 2,
    workbook: { sheets: { Transactions: txSheet() } },
    prompt: "What is the average transaction amount for the Travel category? Round to 2 decimals. Don't change the sheet.",
    check: (api, { answer, before }) => {
      const t = TX.filter((x) => x[1] === "Travel").map((x) => x[2]);
      const avg = round2(t.reduce((a, b) => a + b, 0) / t.length);
      return [check(`states average ${avg}`, mentionsNumber(answer, avg, 0.0005), answer.slice(0, 200)), unchanged(api, before)];
    },
  },
  {
    id: "analysis-find-errors", category: "analysis", level: 2,
    workbook: {
      sheets: {
        Margins: [
          ["Product", "Revenue", "Cost", "Margin", "List price"],
          ["Coffee", 1000, 600, "=(B2-C2)/B2", "=VLOOKUP(A2,Prices!A:B,2,FALSE)"],
          ["Tea", 500, 450, "=(B3-C3)/B3", "=VLOOKUP(A3,Prices!A:B,2,FALSE)"],
          ["Cocoa", 0, 0, "=(B4-C4)/B4", "=VLOOKUP(A4,Prices!A:B,2,FALSE)"],
          ["Mate", 800, 200, "=(B5-C5)/B5", "=VLOOKUP(A5,Prices!A:B,2,FALSE)"],
          ["Chai", 300, 120, "=(B6-C6)/B6", "=VLOOKUP(A6,Prices!A:B,2,FALSE)"],
        ],
        Prices: [["Product", "Price"], ["Coffee", 12.5], ["Tea", 8], ["Cocoa", 10.25], ["Mate", 6.75]],
      },
    },
    prompt: "Check the Margins sheet for errors. Which cells have errors and what causes each one? Don't fix anything yet.",
    check: (api, { answer, before }) => [
      check("finds D4 (#DIV/0!)", mentions(answer, /\bD4\b/), answer.slice(0, 300)),
      check("explains division by zero revenue", mentions(answer, /zero|divid|\b0\b revenue|revenue (is|of) 0/i)),
      check("finds E6 (#N/A)", mentions(answer, /\bE6\b/)),
      check("explains Chai missing from Prices", mentions(answer, /chai/i) && mentions(answer, /not (found|in|listed)|missing|doesn.t exist|no (match|entry)/i)),
      unchanged(api, before),
    ],
  },
  {
    id: "analysis-fastest-growth", category: "analysis", level: 1,
    workbook: { sheets: { Products: productsSheet() } },
    prompt: "Which product grew fastest from 2024 to 2025 in percentage terms, and by what percentage? Don't change the sheet.",
    check: (api, { answer, before }) => [
      check("names Mate", mentions(answer, /\bMate\b/), answer.slice(0, 200)),
      check("states 40%", mentionsNumber(answer, 40, 0.001) || mentionsNumber(answer, 0.4, 0.001)),
      unchanged(api, before),
    ],
  },

  // ===== formulas =====
  {
    id: "formulas-totals", category: "formulas", level: 1,
    workbook: { sheets: { Sales: salesSheet() } },
    prompt: "Add a Total column for each region and a Total row for each quarter, using formulas.",
    check: (api) => {
      const col = findColumn(api, /^total/i, "Sales");
      const north = col && valueAt(api, col, /^North$/);
      const q1 = findColumn(api, /^Q1$/, "Sales");
      const totRow = findLabel(api, /^(grand )?total$/i, "Sales").find((l) => l.r > 0);
      const q1Total = q1 && totRow && api.cells("Sales").find((x) => x.r === totRow.r && x.c === q1.c);
      return [
        check("Total column: North = 565,000", north && near(north.value, regionTotal("North")), north?.value),
        check("Total column uses formulas", north?.formula),
        check("Total row: Q1 = 661,000", q1Total && near(q1Total.value, 661000), q1Total?.value),
        check("Total row uses formulas", q1Total?.formula),
        north && q1Total ? dynamicCheck("totals are live (change North Q1)", api, "Sales!B2", 0, () => api.cells("Sales").find((x) => x.r === north.r && x.c === col.c).value, () => regionTotal("North") - 120000) : check("totals are live", false, "totals not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "formulas-growth-percent", category: "formulas", level: 1,
    workbook: { sheets: { Products: productsSheet() } },
    prompt: "Add a Growth % column showing the change from 2024 to 2025 for each product, formatted as a percentage.",
    check: (api) => {
      const col = findColumn(api, /growth|change|%/i, "Products");
      const mate = col && valueAt(api, col, /^Mate$/);
      return [
        check("Mate growth = 40% (0.4)", mate && near(mate.value, 0.4), mate?.value),
        check("growth uses formulas", mate?.formula),
        check("formatted as percentage", mate && /%/.test(api.style(`Products!${mate.ref}`).numberFormat || ""), mate && api.style(`Products!${mate.ref}`).numberFormat),
        mate ? dynamicCheck("growth is live (change Mate 2025)", api, "Products!C5", 18000, () => api.value(`Products!${mate.ref}`), () => 1) : check("growth is live", false, "growth column not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "formulas-lookup", category: "formulas", level: 2,
    workbook: {
      active: "Orders",
      sheets: {
        Products: [["Code", "Name", "Price"], ["P01", "Coffee", 12.5], ["P02", "Tea", 8], ["P03", "Cocoa", 10.25], ["P04", "Mate", 6.75]],
        Orders: [["Order", "Code", "Qty"], [1001, "P02", 10], [1002, "P01", 4], [1003, "P04", 12], [1004, "P02", 3], [1005, "P03", 7], [1006, "P01", 20], [1007, "P04", 1], [1008, "P03", 5]],
      },
    },
    prompt: "In the Orders sheet, add the unit price for each order from the Products sheet, and a Line Total (Qty × Price). Use lookup formulas so prices stay in sync.",
    check: (api) => {
      const price = findColumn(api, /price/i, "Orders");
      const total = findColumn(api, /total|amount/i, "Orders");
      const p1001 = price && valueAt(api, price, /^1001$/) || (price && api.cells("Orders").find((x) => x.r === 1 && x.c === price.c));
      const t1006 = total && api.cells("Orders").find((x) => x.r === 6 && x.c === total.c);
      return [
        check("price for order 1001 (P02) = 8", p1001 && near(p1001.value, 8), p1001?.value),
        check("line total for 1006 = 250 (20 × 12.5)", t1006 && near(t1006.value, 250), t1006?.value),
        check("price formulas reference Products", p1001?.formula && /Products/i.test(p1001.formula), p1001?.formula),
        p1001 && t1006 ? dynamicCheck("prices are live (change P02 price)", api, "Products!C3", 9, () => api.cells("Orders").find((x) => x.r === 1 && x.c === total.c)?.value, () => 90) : check("prices are live", false, "columns not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "formulas-sumifs-summary", category: "formulas", level: 2,
    workbook: { sheets: { Transactions: txSheet() } },
    prompt: "Create a sheet called Summary with the total amount and the number of transactions for each category, using formulas that reference the Transactions sheet.",
    check: (api) => {
      const out = [check("Summary sheet exists", api.hasSheet("Summary"))];
      if (!api.hasSheet("Summary")) return out;
      const totals = Object.fromEntries(CATS.map((c) => [c, round2(TX.filter((t) => t[1] === c).reduce((a, t) => a + t[2], 0))]));
      const counts = Object.fromEntries(CATS.map((c) => [c, TX.filter((t) => t[1] === c).length]));
      let okT = 0, okC = 0, live = null;
      for (const c of CATS) {
        const lab = findLabel(api, new RegExp(`^${c}$`, "i"), "Summary")[0];
        const nums = lab ? rowNumbersRight(api, lab) : [];
        if (nums.some((n) => near(n.value, totals[c]))) okT++;
        if (nums.some((n) => n.value === counts[c])) okC++;
        if (c === TX[0][1]) live = nums.find((n) => near(n.value, totals[c]));
      }
      out.push(check("category totals correct (4/4)", okT === 4, `${okT}/4`));
      out.push(check("category counts correct (4/4)", okC === 4, `${okC}/4`));
      out.push(check("summary uses formulas", live?.formula, live?.formula));
      out.push(live ? dynamicCheck("summary is live (change first amount)", api, "Transactions!C2", TX[0][2] + 1000, () => api.value(`Summary!${live.ref}`), () => totals[TX[0][1]] + 1000) : check("summary is live", false, "cell not found"));
      out.push(noNewErrors(api));
      return out;
    },
  },
  {
    id: "formulas-commission-tiers", category: "formulas", level: 2,
    workbook: { sheets: { Reps: [["Rep", "Revenue"], ...REPS.map((r) => [...r])] } },
    prompt: "Add a Commission column: 2% of revenue below 50,000; 3% from 50,000 up to and including 100,000; 5% above 100,000. The rate applies to the whole revenue.",
    check: (api) => {
      const col = findColumn(api, /commission/i, "Reps");
      const got = (name) => col && valueAt(api, col, new RegExp(`^${name}$`));
      const bad = REPS.filter(([n, rev]) => !(got(n) && near(got(n).value, commission(rev))));
      const bruno = got("Bruno");
      return [
        check("all 8 commissions correct (incl. 50,000 and 100,000 boundaries)", col && !bad.length, bad.map(([n, r]) => `${n}: got ${got(n)?.value}, want ${commission(r)}`).join("; ")),
        check("commission uses formulas", bruno?.formula),
        bruno ? dynamicCheck("commission is live (Bruno → 120,000)", api, "Reps!B3", 120000, () => api.value(`Reps!${bruno.ref}`), () => 6000) : check("commission is live", false, "column not found"),
        noNewErrors(api),
      ];
    },
  },

  // ===== fix =====
  {
    id: "fix-div0", category: "fix", level: 1,
    workbook: {
      sheets: { Margins: [["Product", "Revenue", "Cost", "Margin"], ["Coffee", 1000, 600, "=(B2-C2)/B2"], ["Tea", 500, 450, "=(B3-C3)/B3"], ["Cocoa", 0, 0, "=(B4-C4)/B4"], ["Mate", 800, 200, "=(B5-C5)/B5"]] },
      formats: { Margins: { "D2:D5": "0.0%" } },
    },
    prompt: "Fix the #DIV/0! error in the Margin column so products with zero revenue show 0 instead of an error. Keep the other margins correct.",
    check: (api) => [
      check("no errors left", !api.errors().length, api.errors().join("; ")),
      check("Cocoa margin shows 0", [0, ""].includes(api.value("Margins!D4")), api.value("Margins!D4")),
      check("other margins unchanged (Mate 75%)", near(api.value("Margins!D5"), 0.75) && near(api.value("Margins!D2"), 0.4)),
      check("margins still formulas", api.formula("Margins!D5")),
      dynamicCheck("fixed formula still live (Cocoa revenue 100)", api, "Margins!B4", 100, () => api.value("Margins!D4"), () => 1),
    ],
  },
  {
    id: "fix-sum-range", category: "fix", level: 1,
    workbook: {
      sheets: {
        Expenses: [["Item", "Amount"], ["Rent", 2500], ["Payroll", 18000], ["Software", 640], ["Travel", 1320], ["Meals", 410], ["Utilities", 380], ["Insurance", 900], ["Marketing", 2200], ["Training", 750], ["Equipment", 3100], ["Total", "=SUM(B2:B10)"]],
      },
    },
    prompt: "The total at the bottom of the Expenses sheet looks too low. Find and fix the problem.",
    check: (api) => [
      check("total = 30,200 (includes Equipment)", near(api.value("Expenses!B12"), 30200), api.value("Expenses!B12")),
      check("total is a formula", api.formula("Expenses!B12")),
      dynamicCheck("total is live (Equipment → 0)", api, "Expenses!B11", 0, () => api.value("Expenses!B12"), () => 27100),
    ],
  },
  {
    id: "fix-hardcoded-totals", category: "fix", level: 2,
    workbook: {
      sheets: { Budget: [["Dept", "Jan", "Feb", "Mar", "Total"], ["Sales", 12000, 13500, 12800, 38300], ["Ops", 8000, 8200, 8100, 24300], ["IT", 5000, 5600, 5100, 15800], ["HR", 3000, 3000, 3200, 9200]] },
    },
    prompt: "Make the Total column use live formulas instead of typed numbers, and tell me which of the old totals (if any) were wrong.",
    check: (api, { answer }) => [
      check("totals correct (IT = 15,700)", near(api.value("Budget!E4"), 15700) && near(api.value("Budget!E2"), 38300)),
      check("totals are formulas", ["E2", "E3", "E4", "E5"].every((r) => api.formula(`Budget!${r}`))),
      check("reports IT's old total was wrong", mentions(answer, /\bIT\b/) && (mentionsNumber(answer, 15800) || mentionsNumber(answer, 15700)), answer.slice(0, 250)),
      check("doesn't blame correct rows", !/\b(Ops|HR)\b[^.\n]{0,40}(wrong|incorrect|off)/i.test(answer)),
      dynamicCheck("totals are live (HR Mar → 4,000)", api, "Budget!D5", 4000, () => api.value("Budget!E5"), () => 10000),
    ],
  },

  // ===== cleanup =====
  {
    id: "clean-split-names", category: "cleanup", level: 1,
    workbook: { sheets: { Contacts: [["Full Name", "Email"], ["Ana Quispe", "ana@x.bo"], ["Carlos Mendoza", "carlos@x.bo"], ["Lucía Vargas", "lucia@x.bo"], ["Jorge Mamani", "jorge@x.bo"], ["Sofía Rojas", "sofia@x.bo"], ["Mateo Choque", "mateo@x.bo"]] } },
    prompt: "Split the Full Name column into separate First Name and Last Name columns.",
    check: (api) => {
      // Pairs are matched row by row, so it doesn't matter whether the original Full Name column was kept.
      const first = findColumn(api, /first/i, "Contacts");
      const last = findColumn(api, /last|surname|apellido/i, "Contacts");
      const pairs = first && last
        ? new Set(api.cells("Contacts").filter((c) => c.c === first.c && c.r > first.headerRow).map((c) => `${c.value}|${api.cells("Contacts").find((x) => x.r === c.r && x.c === last.c)?.value}`))
        : new Set();
      const want = ["Ana Quispe", "Carlos Mendoza", "Lucía Vargas", "Jorge Mamani", "Sofía Rojas", "Mateo Choque"].map((n) => n.replace(" ", "|"));
      return [
        check("Lucía / Vargas split correctly (accents kept)", pairs.has("Lucía|Vargas"), [...pairs].slice(0, 3).join(", ")),
        check("all 6 names split", want.every((w) => pairs.has(w)), `${want.filter((w) => pairs.has(w)).length}/6`),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "clean-text-numbers", category: "cleanup", level: 2,
    workbook: { sheets: { Payments: [["Ref", "Amount"], ["INV-01", "'1,234.50"], ["INV-02", "' 99 "], ["INV-03", "'2,000"], ["INV-04", "'15.75"], ["INV-05", "'300"], ["INV-06", "'48.25 "]] } },
    prompt: "The Amount column is stored as text, so it can't be totalled. Convert the amounts to real numbers and add a total below them.",
    check: (api) => {
      const amounts = api.cells("Payments").filter((c) => c.c === 1 && c.r >= 1 && c.r <= 6);
      const tot = findLabel(api, /total/i, "Payments").map((l) => rowNumbersRight(api, l)[0]).find(Boolean) || api.cells("Payments").find((c) => c.r === 7 && typeof c.value === "number");
      return [
        check("all 6 amounts are numbers", amounts.length === 6 && amounts.every((c) => typeof c.value === "number"), amounts.map((c) => JSON.stringify(c.value)).join(",")),
        check("values correct (1,234.50 and 99)", near(api.value("Payments!B2"), 1234.5) && near(api.value("Payments!B3"), 99)),
        check("total = 3,697.50", tot && near(tot.value, 3697.5), tot?.value),
        check("total is a formula", tot?.formula),
      ];
    },
  },
  {
    id: "clean-dedupe", category: "cleanup", level: 2,
    workbook: {
      sheets: {
        Customers: (() => {
          const base = Array.from({ length: 25 }, (_, i) => [`C${String(i + 1).padStart(3, "0")}`, `Customer ${i + 1}`, ["La Paz", "Santa Cruz", "Cochabamba"][i % 3]]);
          const rows = [...base];
          for (const i of [3, 7, 12, 18, 21]) rows.splice(i + 4, 0, [...base[i]]);
          return [["ID", "Name", "City"], ...rows];
        })(),
      },
    },
    prompt: "Remove the duplicate rows from the Customers sheet, keeping the first occurrence of each.",
    check: (api) => {
      const ids = api.cells("Customers").filter((c) => c.c === 0 && c.r > 0).map((c) => c.value);
      const uniq = new Set(ids);
      return [
        check("25 rows remain", ids.length === 25, `${ids.length} rows`),
        check("no duplicate IDs", uniq.size === ids.length),
        check("all 25 customers kept", uniq.size === 25),
        check("header intact", api.value("Customers!A1") === "ID"),
      ];
    },
  },
  {
    id: "clean-sort", category: "cleanup", level: 1,
    workbook: { sheets: { Reps: [["Rep", "Revenue"], ...REPS.map((r) => [...r])] } },
    prompt: "Sort the reps by revenue, highest first.",
    check: (api) => {
      const rows = api.cells("Reps").filter((c) => c.r > 0);
      const names = rows.filter((c) => c.c === 0).map((c) => c.value);
      const revs = rows.filter((c) => c.c === 1).map((c) => c.value);
      const sorted = [...REPS].sort((a, b) => b[1] - a[1]);
      return [
        check("order is Fabio, Elena, Diego, …, Gaby", JSON.stringify(names) === JSON.stringify(sorted.map((r) => r[0])), names.join(", ")),
        check("each rep keeps their revenue", names.every((n, i) => REPS.find((r) => r[0] === n)?.[1] === revs[i])),
      ];
    },
  },
  {
    id: "clean-standardize", category: "cleanup", level: 2,
    workbook: {
      sheets: { Expenses: [["Date", "Category", "Amount"], ["2026-02-01", "travel", 120], ["2026-02-02", "Travel ", 80], ["2026-02-03", "TRAVEL", 60], ["2026-02-04", "Meals", 25], ["2026-02-05", " meals", 30], ["2026-02-06", "software", 99], ["2026-02-07", "Software", 49], ["2026-02-08", "MEALS ", 18]] },
    },
    prompt: "Standardize the Category column so each category is spelled exactly the same way (Travel, Meals, Software) with no extra spaces.",
    check: (api) => {
      const cats = api.cells("Expenses").filter((c) => c.c === 1 && c.r > 0).map((c) => c.value);
      const amts = api.cells("Expenses").filter((c) => c.c === 2 && c.r > 0).map((c) => c.value);
      return [
        check("only Travel / Meals / Software remain", cats.length === 8 && cats.every((c) => ["Travel", "Meals", "Software"].includes(c)), JSON.stringify(cats)),
        check("rows and amounts unchanged", JSON.stringify(amts) === "[120,80,60,25,30,99,49,18]"),
      ];
    },
  },

  // ===== modeling =====
  {
    id: "model-loan", category: "modeling", level: 2,
    workbook: { sheets: { Loan: [] } },
    prompt: "Build a loan calculator on this sheet: loan amount 250,000, annual interest rate 6.5%, term 30 years. Calculate the monthly payment, total paid and total interest, with the inputs in their own cells.",
    check: (api) => {
      const pmt = (P, a, y) => { const r = a / 12, n = y * 12; return (P * r) / (1 - (1 + r) ** -n); };
      const want = pmt(250000, 0.065, 30);
      const pay = api.cells("Loan").find((c) => c.formula && near(Math.abs(c.value), want, 0.0005));
      const interest = api.cells("Loan").find((c) => c.formula && near(Math.abs(c.value), want * 360 - 250000, 0.001));
      const rate = api.cells("Loan").find((c) => !c.formula && (c.value === 0.065 || c.value === 6.5));
      return [
        check(`monthly payment ≈ ${want.toFixed(2)}`, pay, pay ? `${pay.ref}=${pay.value}` : "not found"),
        check(`total interest ≈ ${Math.round(want * 360 - 250000).toLocaleString("en-US")}`, interest, interest?.value),
        check("rate is an input cell", rate),
        pay && rate ? dynamicCheck("payment is live (rate → 5%)", api, `Loan!${rate.ref}`, rate.value === 6.5 ? 5 : 0.05, () => Math.abs(api.value(`Loan!${pay.ref}`)), () => pmt(250000, 0.05, 30)) : check("payment is live", false, "cells not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "model-projection", category: "modeling", level: 3,
    workbook: { sheets: { Plan: [] } },
    prompt: "Create a 3-year projection on this sheet: Year 1 revenue is 1,200,000 and it grows 8% per year; costs are 65% of revenue. Show revenue, costs and profit for Years 1–3, with the growth rate and cost ratio in their own input cells.",
    check: (api) => {
      const y3 = api.cells("Plan").find((c) => c.formula && near(c.value, 1399680));
      const p3 = api.cells("Plan").find((c) => c.formula && near(c.value, 1399680 * 0.35));
      const g = api.cells("Plan").find((c) => !c.formula && (c.value === 0.08 || c.value === 8));
      const cr = api.cells("Plan").find((c) => !c.formula && (c.value === 0.65 || c.value === 65));
      return [
        check("Year 3 revenue = 1,399,680", y3, y3?.ref),
        check("Year 3 profit = 489,888", p3, p3?.ref),
        check("growth and cost ratio are input cells", g && cr),
        y3 && g ? dynamicCheck("projection is live (growth → 10%)", api, `Plan!${g.ref}`, g.value === 8 ? 10 : 0.1, () => api.value(`Plan!${y3.ref}`), () => 1452000) : check("projection is live", false, "cells not found"),
        p3 && cr ? dynamicCheck("profit follows cost ratio (→ 60%)", api, `Plan!${cr.ref}`, cr.value === 65 ? 60 : 0.6, () => api.value(`Plan!${p3.ref}`), () => 1399680 * 0.4) : check("profit follows cost ratio", false, "cells not found"),
        noNewErrors(api),
      ];
    },
  },
  {
    id: "model-breakeven", category: "modeling", level: 2,
    workbook: { sheets: { BreakEven: [] } },
    prompt: "Build a break-even analysis: fixed costs 50,000, price per unit 25, variable cost per unit 15. Calculate break-even units and break-even revenue, with the inputs in separate cells.",
    check: (api) => {
      const units = api.cells("BreakEven").find((c) => c.formula && near(c.value, 5000));
      const rev = api.cells("BreakEven").find((c) => c.formula && near(c.value, 125000));
      const price = api.cells("BreakEven").find((c) => !c.formula && c.value === 25);
      return [
        check("break-even units = 5,000", units, units?.ref),
        check("break-even revenue = 125,000", rev, rev?.ref),
        units && price ? dynamicCheck("live (price → 30)", api, `BreakEven!${price.ref}`, 30, () => api.value(`BreakEven!${units.ref}`), () => 50000 / 15) : check("live", false, "cells not found"),
        noNewErrors(api),
      ];
    },
  },

  // ===== visual =====
  {
    id: "visual-table-chart", category: "visual", level: 1,
    workbook: { sheets: { Sales: salesSheet() } },
    prompt: "Turn the sales data into an Excel table and add a clustered column chart comparing the quarters by region.",
    check: (api) => {
      const t = api.tables()[0];
      const c = api.charts()[0];
      return [
        check("table covers A1:E7", t && /A1:E7$/.test(t.address), t?.address),
        check("column chart created", c && /column/i.test(c.type), c?.type),
        check("chart uses the sales data", c && /A1:E7$|A1:E7/.test(c.source), c?.source),
      ];
    },
  },
  {
    id: "visual-format", category: "visual", level: 2,
    workbook: { sheets: { Sales: salesSheet() } },
    prompt: "Format this sheet professionally: bold header row with a fill color, thousands separators on the sales numbers, and freeze the header row.",
    check: (api) => [
      check("header bold", api.style("Sales!A1").bold && api.style("Sales!E1").bold),
      check("header has a fill", api.style("Sales!C1").fill),
      check("numbers use thousands separators", /#,##0/.test(api.style("Sales!C4").numberFormat || ""), api.style("Sales!C4").numberFormat),
      check("header row frozen", api.frozen("Sales").rows >= 1, JSON.stringify(api.frozen("Sales"))),
    ],
  },

  // ===== files (attachments) =====
  {
    id: "file-csv-import", category: "files", level: 2,
    workbook: { sheets: { Sheet1: [] } },
    attachments: ["transactions.csv"],
    prompt: "Import the attached CSV into a new sheet called Data, then add a total of the Amount column below the data.",
    check: (api, { fixtures }) => {
      const out = [check("Data sheet exists", api.hasSheet("Data"))];
      if (!api.hasSheet("Data")) return out;
      const col = findColumn(api, /^amount$/i, "Data");
      const amounts = col ? columnCells(api, col, col.headerRow + 1, col.headerRow + 50).filter((c) => typeof c.value === "number") : [];
      const want = fixtures.csvTotal;
      const tot = col && api.cells("Data").find((c) => c.c === col.c && c.formula && near(c.value, want));
      out.push(check("all 50 rows imported", amounts.length >= 50, `${amounts.length} numeric amounts`));
      out.push(check(`total = ${want}`, tot, tot?.ref));
      out.push(check("vendor text with leading +/- kept as text", api.find("Data", "+Plus Market").length === 1));
      return out;
    },
  },
  {
    id: "file-pdf-invoice", category: "files", level: 2,
    workbook: { sheets: { Sheet1: [] } },
    attachments: ["invoice.pdf"],
    prompt: "Put the line items from the attached invoice PDF into a new sheet called Invoice with columns Item, Qty, Unit Price and Amount, and add a total with a formula.",
    check: (api) => {
      const out = [check("Invoice sheet exists", api.hasSheet("Invoice"))];
      if (!api.hasSheet("Invoice")) return out;
      const tot = api.cells("Invoice").find((c) => c.formula && near(c.value, 389));
      const beans = findLabel(api, /arabica/i, "Invoice")[0];
      const beanNums = beans ? rowNumbersRight(api, beans).map((c) => c.value) : [];
      out.push(check("Arabica row: 12 × 18.50 = 222", [12, 18.5, 222].every((n) => beanNums.some((v) => near(v, n))), beanNums.join(",")));
      out.push(check("total = 389.00 by formula", tot, tot?.ref));
      return out;
    },
  },
  {
    id: "file-pptx-chart", category: "files", level: 2,
    workbook: { sheets: { Sheet1: [] } },
    attachments: ["budget-deck.pptx"],
    prompt: "Put the revenue-by-source chart data from slide 5 of the attached presentation into a new sheet called Revenue, with a total row.",
    check: (api) => {
      const out = [check("Revenue sheet exists", api.hasSheet("Revenue"))];
      if (!api.hasSheet("Revenue")) return out;
      const vals = [35742525, 6127000, 2960699.46, 18000, 891000, 1140400, 1223023];
      const nums = api.cells("Revenue").map((c) => c.value).filter((v) => typeof v === "number");
      const tot = api.cells("Revenue").find((c) => c.formula && near(c.value, 48102647.46, 0.0001));
      out.push(check("all 7 sources with exact values", vals.every((v) => nums.some((n) => near(n, v, 0.00001))), `${vals.filter((v) => nums.some((n) => near(n, v, 0.00001))).length}/7`));
      out.push(check("total = 48,102,647.46 by formula", tot, tot?.ref));
      return out;
    },
  },
  {
    id: "file-word-table", category: "files", level: 2,
    workbook: { sheets: { Sheet1: [] } },
    attachments: ["budget-memo.doc"],
    prompt: "Put the department budget table from the attached memo into a new sheet called Budget, and calculate Variance as Budget minus Actual with formulas.",
    check: (api) => {
      const out = [check("Budget sheet exists", api.hasSheet("Budget"))];
      if (!api.hasSheet("Budget")) return out;
      const col = findColumn(api, /variance/i, "Budget");
      const v = (re) => col && valueAt(api, col, re);
      out.push(check("Marketing variance = -11,500", near(v(/^Marketing$/)?.value, -11500), v(/^Marketing$/)?.value));
      out.push(check("Operations variance = 17,250", near(v(/^Operations$/)?.value, 17250)));
      out.push(check("variance uses formulas", v(/^IT$/)?.formula));
      return out;
    },
  },

  // ===== multi-sheet =====
  {
    id: "multisheet-quarter", category: "multi-sheet", level: 3,
    workbook: {
      active: "Jan",
      sheets: Object.fromEntries(["Jan", "Feb", "Mar"].map((m, i) => [m, [["Product", "Units", "Revenue"], ["Coffee", 120 + i * 10, 1500 + i * 125], ["Tea", 80 + i * 5, 640 + i * 40], ["Cocoa", 40 + i * 4, 410 + i * 41], ["Mate", 25 + i * 15, 168.75 + i * 101.25]]])),
    },
    prompt: "Create a Q1 sheet that totals Units and Revenue for each product across the Jan, Feb and Mar sheets, using formulas.",
    check: (api) => {
      const out = [check("Q1 sheet exists", api.hasSheet("Q1"))];
      if (!api.hasSheet("Q1")) return out;
      const coffee = findLabel(api, /^coffee$/i, "Q1")[0];
      const nums = coffee ? rowNumbersRight(api, coffee) : [];
      const rev = nums.find((c) => near(c.value, 1500 + 1625 + 1750));
      const units = nums.find((c) => c.value === 120 + 130 + 140);
      const mate = findLabel(api, /^mate$/i, "Q1")[0];
      const mateRev = mate && rowNumbersRight(api, mate).find((c) => near(c.value, 168.75 + 270 + 371.25));
      out.push(check("Coffee units = 390", units, nums.map((n) => n.value).join(",")));
      out.push(check("Coffee revenue = 4,875", rev));
      out.push(check("Mate revenue = 810", mateRev));
      out.push(rev ? dynamicCheck("Q1 is live (Feb Coffee revenue → 2,000)", api, "Feb!C2", 2000, () => api.value(`Q1!${rev.ref}`), () => 1500 + 2000 + 1750) : check("Q1 is live", false, "cell not found"));
      out.push(noNewErrors(api));
      return out;
    },
  },
];

// ---------- fixture files for attachment tasks ----------
export const CSV_ROWS = (() => {
  const r = rng(7);
  const vendors = ["Andes Coffee Co.", "Tigo", "Entel", "+Plus Market", "Hipermaxi", "YPFB", "-Minus Supplies"];
  return Array.from({ length: 50 }, (_, i) => [
    `2026-03-${String(1 + (i % 28)).padStart(2, "0")}`,
    vendors[Math.floor(r() * vendors.length)],
    CATS[Math.floor(r() * CATS.length)],
    round2(5 + r() * 495),
  ]);
})();
// Make sure the "+Plus Market" text-escaping case is present exactly once.
CSV_ROWS[3][1] = "+Plus Market";
for (let i = 0; i < CSV_ROWS.length; i++) if (i !== 3 && CSV_ROWS[i][1] === "+Plus Market") CSV_ROWS[i][1] = "Tigo";
export const CSV_TOTAL = round2(CSV_ROWS.reduce((a, r) => a + r[3], 0));
