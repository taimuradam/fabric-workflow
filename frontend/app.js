/* ============================================================================
   app.js — Textile Lot Costing frontend behaviour.

   Talks to Python (calculations / storage / excel_export) exclusively through
   window.pywebview.api, which api.py exposes. This file never does any costing
   maths itself — it collects inputs, asks Python to compute, and repaints. That
   guarantees the numbers match the underlying (unchanged) calculation module.
   ========================================================================== */

"use strict";

// The lot-info input ids, matching the LotInfo dataclass field names.
const LOT_FIELDS = [
  "reference", "fabric_type", "date", "gsm", "width_in",
  "total_kg", "rate_per_meter", "transport_cost",
];

// Which lot-calculation tiles are currency (get the "Rs" prefix).
const MONEY_CALCS = new Set([
  "fabric_cost", "total_cost", "base_cost_per_kg",
  "wastage_cost_per_kg", "adjusted_cost_per_kg",
]);

let lastResults = null;       // most recent compute() result (for Excel export)
let recomputeTimer = null;    // debounce handle

/* ----------------------------------------------------------- formatting --- */

function todayISO() {
  // Local calendar date as yyyy-mm-dd (avoids the UTC shift of toISOString).
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}

function fmt(value, { money = false, decimals = 2 } = {}) {
  // Mirror calculations.py: a null/None dependent value renders as an em dash.
  if (value === null || value === undefined || Number.isNaN(value)) return "—";
  const n = Number(value).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  return money ? `Rs ${n}` : n;
}

/* --------------------------------------------------------- reading state --- */

function readLot() {
  const lot = {};
  for (const id of LOT_FIELDS) {
    const el = document.getElementById(id);
    const val = el.value.trim();
    // Text fields keep their string; numeric blanks become null so Python sees
    // "not entered" rather than an empty string (both are handled, this is tidy).
    lot[id] = val === "" ? null : val;
  }
  return lot;
}

function readProducts() {
  const rows = document.querySelectorAll("#product-rows tr");
  const products = [];
  rows.forEach((row) => {
    const name = row.querySelector(".p-name").value;
    const weight = row.querySelector(".p-weight").value.trim();
    const pieces = row.querySelector(".p-pieces").value.trim();
    const isWastage = row.querySelector(".p-wastage").checked;
    products.push({
      name: name,
      weight_kg: weight === "" ? null : weight,
      pieces: pieces === "" ? null : pieces,
      is_wastage: isWastage,
    });
  });
  return products;
}

/* ------------------------------------------------------------- rendering --- */

function paint(results) {
  lastResults = results;

  // --- Lot calculation tiles ---
  document.querySelectorAll("[data-calc]").forEach((el) => {
    const key = el.dataset.calc;
    el.textContent = fmt(results[key], { money: MONEY_CALCS.has(key) });
  });

  // --- Per-row computed cells (order matches readProducts()) ---
  const rows = document.querySelectorAll("#product-rows tr");
  const pr = results.product_results || [];
  rows.forEach((row, i) => {
    const res = pr[i] || {};
    const isWastage = row.querySelector(".p-wastage").checked;
    const wtpc = row.querySelector(".c-wtpc");
    const costpc = row.querySelector(".c-costpc");
    const revenue = row.querySelector(".c-revenue");

    if (isWastage) {
      // Wastage row is pure cost — no per-piece figures, no revenue.
      setCalc(wtpc, "—");
      setCalc(costpc, "—");
      setCalc(revenue, "—");
    } else {
      // weight_per_piece === null means pieces was 0/blank -> "N/A".
      setCalc(wtpc, res.weight_per_piece == null ? "N/A" : fmt(res.weight_per_piece));
      setCalc(costpc, fmt(res.cost_per_piece, { money: true }));
      setCalc(revenue, fmt(res.revenue, { money: true }));
    }
  });

  // --- Totals (shared by the table footer and the Summary card) ---
  setTotals("total_weight_produced", fmt(results.total_weight_produced));
  setTotals("total_pieces", fmt(results.total_pieces, { decimals: 0 }));
  setTotals("total_payment", fmt(results.total_payment, { money: true }));
  setTotals("total_receipt", fmt(results.total_receipt, { money: true }));

  // --- Profit (colour-coded) ---
  const profitEl = document.getElementById("profit-value");
  profitEl.classList.remove("profit--pos", "profit--neg");
  if (results.profit === null || results.profit === undefined) {
    profitEl.textContent = "—";
  } else {
    profitEl.textContent = fmt(results.profit, { money: true });
    profitEl.classList.add(results.profit >= 0 ? "profit--pos" : "profit--neg");
  }

  // --- Reconciliation warning banner ---
  const banner = document.getElementById("recon-warning");
  if (results.recon_mismatch) {
    document.getElementById("recon-message").textContent = results.recon_message;
    banner.hidden = false;
  } else {
    banner.hidden = true;
  }
}

function setCalc(el, text) {
  el.textContent = text;
  el.classList.toggle("is-empty", text === "—" || text === "N/A");
}

function setTotals(key, text) {
  document.querySelectorAll(`[data-total="${key}"]`).forEach((el) => {
    el.textContent = text;
  });
}

/* --------------------------------------------------------------- compute --- */

async function recompute() {
  const api = window.pywebview && window.pywebview.api;
  if (!api) return; // bridge not ready yet
  try {
    const results = await api.compute(readLot(), readProducts());
    paint(results);
  } catch (err) {
    console.error("compute failed", err);
  }
}

function scheduleRecompute() {
  // Debounce rapid typing so we don't hammer the Python bridge on every keypress.
  clearTimeout(recomputeTimer);
  recomputeTimer = setTimeout(recompute, 150);
}

/* ----------------------------------------------------------- product rows -- */

function makeRow(product = {}) {
  const tr = document.createElement("tr");
  const weight = product.weight_kg == null ? "" : product.weight_kg;
  const pieces = product.pieces == null ? "" : product.pieces;
  const wastage = !!product.is_wastage;

  tr.innerHTML = `
    <td><input class="cell-input p-name" type="text" placeholder="Product name" /></td>
    <td><input class="cell-input num p-weight" type="text" inputmode="decimal" /></td>
    <td><input class="cell-input num p-pieces" type="text" inputmode="decimal" /></td>
    <td class="waste"><input class="waste-check p-wastage" type="checkbox" title="Mark as wastage" /></td>
    <td class="num"><span class="cell-calc c-wtpc is-empty">—</span></td>
    <td class="num"><span class="cell-calc c-costpc is-empty">—</span></td>
    <td class="num"><span class="cell-calc c-revenue is-empty">—</span></td>
    <td class="waste"><button class="row-del" type="button" title="Delete row">✕</button></td>
  `;

  // Populate values (set via .value to avoid HTML-escaping issues with names).
  tr.querySelector(".p-name").value = product.name || "";
  tr.querySelector(".p-weight").value = weight;
  tr.querySelector(".p-pieces").value = pieces;
  const check = tr.querySelector(".p-wastage");
  check.checked = wastage;
  applyWastageStyling(tr, wastage);

  // Wire up events.
  tr.querySelectorAll(".p-name, .p-weight, .p-pieces").forEach((inp) =>
    inp.addEventListener("input", scheduleRecompute)
  );
  check.addEventListener("change", () => onWastageToggle(tr));
  tr.querySelector(".row-del").addEventListener("click", () => {
    tr.remove();
    recompute();
  });

  return tr;
}

function addRow(product) {
  document.getElementById("product-rows").appendChild(makeRow(product));
}

function onWastageToggle(row) {
  const check = row.querySelector(".p-wastage");
  if (check.checked) {
    // Enforce a single wastage row: clear the flag on every other row.
    document.querySelectorAll("#product-rows tr").forEach((other) => {
      if (other !== row) {
        const c = other.querySelector(".p-wastage");
        if (c.checked) {
          c.checked = false;
          applyWastageStyling(other, false);
        }
      }
    });
  }
  applyWastageStyling(row, check.checked);
  recompute();
}

function applyWastageStyling(row, isWastage) {
  row.classList.toggle("products__wastage", isWastage);
  // A wastage row has no pieces — disable + clear that input to make it obvious.
  const pieces = row.querySelector(".p-pieces");
  pieces.disabled = isWastage;
  if (isWastage) pieces.value = "";
}

/* ------------------------------------------------------- load / new / clear */

function seedDefaultRows() {
  document.getElementById("product-rows").innerHTML = "";
  addRow({ name: "" });
  addRow({ name: "" });
  addRow({ name: "Wastage", is_wastage: true });
}

function newLot() {
  for (const id of LOT_FIELDS) {
    document.getElementById(id).value = id === "date" ? todayISO() : "";
  }
  seedDefaultRows();
  recompute();
}

function populate(lot, products) {
  for (const id of LOT_FIELDS) {
    const val = lot[id];
    document.getElementById(id).value = val == null ? "" : String(val);
  }
  const tbody = document.getElementById("product-rows");
  tbody.innerHTML = "";
  (products || []).forEach((p) => addRow(p));
  if (!products || products.length === 0) seedDefaultRows();
  recompute();
}

/* ------------------------------------------------------------- toolbar ops */

function toast(message, isError = false) {
  const el = document.getElementById("toast");
  el.textContent = message;
  el.classList.toggle("toast--error", isError);
  el.hidden = false;
  // force reflow so the transition runs each time
  void el.offsetWidth;
  el.classList.add("toast--show");
  setTimeout(() => {
    el.classList.remove("toast--show");
    setTimeout(() => (el.hidden = true), 220);
  }, 2600);
}

async function saveLot() {
  const res = await window.pywebview.api.save_lot(readLot(), readProducts());
  if (res.ok) toast("Lot saved.");
  else if (res.error) toast("Could not save: " + res.error, true);
}

async function openLot() {
  const res = await window.pywebview.api.load_lot();
  if (res.ok) {
    populate(res.lot, res.products);
    toast("Lot opened.");
  } else if (res.error) {
    toast("Could not open: " + res.error, true);
  }
}

async function exportExcel() {
  const res = await window.pywebview.api.export_excel(
    readLot(), readProducts(), lastResults
  );
  if (res.ok) toast("Excel file created.");
  else if (res.error) toast("Could not export: " + res.error, true);
}

/* ------------------------------------------------------------------- init */

function bindToolbar() {
  document.getElementById("btn-new").addEventListener("click", newLot);
  document.getElementById("btn-save").addEventListener("click", saveLot);
  document.getElementById("btn-open").addEventListener("click", openLot);
  document.getElementById("btn-export").addEventListener("click", exportExcel);
  document.getElementById("btn-add-row").addEventListener("click", () => {
    addRow({ name: "" });
    recompute();
  });

  // Live recompute when any lot-info field changes.
  for (const id of LOT_FIELDS) {
    document.getElementById(id).addEventListener("input", scheduleRecompute);
  }
}

function init() {
  bindToolbar();
  // Default date + starter rows before the bridge is up; numbers fill in once
  // pywebview is ready (compute() no-ops until then).
  document.getElementById("date").value = todayISO();
  seedDefaultRows();
}

document.addEventListener("DOMContentLoaded", init);
// pywebview injects its api asynchronously; run the first real compute then.
window.addEventListener("pywebviewready", recompute);
