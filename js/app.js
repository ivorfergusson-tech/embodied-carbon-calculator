// UI controller: reads state, calls carbon.js, re-renders. No carbon maths
// lives here — if a number's wrong, the bug is in carbon.js or factors-data.js.

const STORAGE_KEY = "embodied-carbon-calculator-v1";

let state = loadState();

function defaultState() {
  return {
    assumptions: { ...DEFAULT_ASSUMPTIONS },
    rows: SAMPLE_ELEMENTS.map(r => ({ ...r })),
  };
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved && saved.assumptions && Array.isArray(saved.rows)) return saved;
  } catch (e) { /* storage unavailable or corrupt: fall back to the worked example */ }
  return defaultState();
}

function saveState() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (e) { /* not critical */ }
}

// ---------- formatting ----------
const fmt = (n, dp = 0) => n.toLocaleString("en-GB", { minimumFractionDigits: dp, maximumFractionDigits: dp });
const tonnes = kg => fmt(kg / 1000, 2);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ---------- assumptions ----------
const asFields = {
  projectName: document.getElementById("as-name"),
  distanceKm: document.getElementById("as-distance"),
  transportFactor: document.getElementById("as-transport"),
  a5Pct: document.getElementById("as-a5"),
};

function fillAssumptions() {
  for (const [key, el] of Object.entries(asFields)) el.value = state.assumptions[key];
}

for (const [key, el] of Object.entries(asFields)) {
  el.addEventListener("input", () => {
    state.assumptions[key] = key === "projectName" ? el.value : Math.max(0, Number(el.value) || 0);
    saveState();
    renderResults();
  });
}

// ---------- take-off rows ----------
const rowsBody = document.getElementById("rows-body");

function materialOptions(selected) {
  return FACTORS.map(f => `<option value="${f.id}"${f.id === selected ? " selected" : ""}>${esc(f.name)}</option>`).join("");
}

function renderRows() {
  if (!state.rows.length) {
    rowsBody.innerHTML = `<tr><td colspan="10" class="muted empty">No rows yet. Add one, or load the worked example.</td></tr>`;
    return;
  }
  rowsBody.innerHTML = state.rows.map((r, i) => {
    const mat = getFactor(r.material);
    const canM3 = mat && mat.density;
    return `<tr data-i="${i}">
      <td><input type="text" data-k="element" value="${esc(r.element)}" placeholder="e.g. Ground floor slab" aria-label="Element" /></td>
      <td><select data-k="material" aria-label="Material">${materialOptions(r.material)}</select></td>
      <td class="num"><input type="number" data-k="qty" value="${r.qty}" step="any" min="0" aria-label="Quantity" /></td>
      <td><select data-k="unit" aria-label="Unit">
        <option value="kg"${r.unit === "kg" ? " selected" : ""}>kg</option>
        <option value="m3"${r.unit === "m3" ? " selected" : ""}${canM3 ? "" : " disabled"}>m³</option>
      </select></td>
      <td class="num" data-out="mass"></td>
      <td class="num" data-out="a13"></td>
      <td class="num" data-out="a4"></td>
      <td class="num" data-out="a5"></td>
      <td class="num strong" data-out="total"></td>
      <td><button type="button" class="link-btn" data-remove aria-label="Remove row">Remove</button></td>
    </tr>`;
  }).join("");
}

rowsBody.addEventListener("input", e => {
  const tr = e.target.closest("tr[data-i]");
  const key = e.target.dataset.k;
  if (!tr || !key) return;
  const row = state.rows[Number(tr.dataset.i)];
  row[key] = e.target.value;

  if (key === "material") {
    // Materials without a density can only be taken off by mass
    const mat = getFactor(row.material);
    if (!mat.density && row.unit === "m3") row.unit = "kg";
    saveState();
    renderRows();
  } else {
    saveState();
  }
  renderResults();
});

rowsBody.addEventListener("click", e => {
  if (!e.target.matches("[data-remove]")) return;
  state.rows.splice(Number(e.target.closest("tr").dataset.i), 1);
  saveState();
  renderAll();
});

document.getElementById("btn-add").addEventListener("click", () => {
  state.rows.push({ element: "", material: "concrete", qty: "", unit: "m3" });
  saveState();
  renderAll();
  const inputs = rowsBody.querySelectorAll('input[data-k="element"]');
  inputs[inputs.length - 1].focus();
});

document.getElementById("btn-sample").addEventListener("click", () => {
  state = defaultState();
  saveState();
  fillAssumptions();
  renderAll();
});

document.getElementById("btn-clear").addEventListener("click", () => {
  if (!confirm("Remove every row from the take-off?")) return;
  state.rows = [];
  saveState();
  renderAll();
});

// ---------- results ----------
function renderResults() {
  const a = state.assumptions;
  const p = calcProject(state.rows, a);

  // per-row outputs
  rowsBody.querySelectorAll("tr[data-i]").forEach(tr => {
    const r = p.results[Number(tr.dataset.i)];
    const set = (k, v) => { tr.querySelector(`[data-out="${k}"]`).textContent = v; };
    if (!r) { ["mass", "a13", "a4", "a5", "total"].forEach(k => set(k, "–")); return; }
    set("mass", fmt(r.mass)); set("a13", fmt(r.a13)); set("a4", fmt(r.a4, 1)); set("a5", fmt(r.a5, 1)); set("total", fmt(r.total));
  });

  document.getElementById("foot-a13").textContent = fmt(p.a13);
  document.getElementById("foot-a4").textContent = fmt(p.a4, 1);
  document.getElementById("foot-a5").textContent = fmt(p.a5, 1);
  document.getElementById("foot-total").textContent = fmt(p.total);

  // headline
  document.getElementById("res-total").textContent = tonnes(p.total);
  document.getElementById("res-project").textContent = a.projectName;
  document.getElementById("res-a13").textContent = tonnes(p.a13) + " t";
  document.getElementById("res-a4").textContent = tonnes(p.a4) + " t";
  document.getElementById("res-a5").textContent = tonnes(p.a5) + " t";
  const pct = v => (p.total ? (v / p.total) * 100 : 0);
  document.getElementById("stage-bar").innerHTML =
    `<span class="s-a13" style="width:${pct(p.a13)}%"></span><span class="s-a4" style="width:${pct(p.a4)}%"></span><span class="s-a5" style="width:${pct(p.a5)}%"></span>`;

  // recycled vs virgin steel
  const c = calcVirginComparison(state.rows, a);
  const compare = document.getElementById("compare");
  if (!c.count) {
    compare.innerHTML = `<p class="muted small">Add a row using <em>UK average (recycled) rebar</em> to see how much virgin steel would add.</p>`;
  } else {
    const newTotal = p.total + c.extra;
    const times = c.recycled ? (c.virgin / c.recycled).toFixed(1) : "–";
    compare.innerHTML = `
      <p class="compare-title">What if the reinforcement were virgin steel instead of UK-average recycled?</p>
      <div class="compare-grid">
        <div><span class="muted small">Rebar as specified (recycled)</span><strong>${tonnes(c.recycled)} tCO₂e</strong></div>
        <div><span class="muted small">Same mass, virgin steel</span><strong class="warn">${tonnes(c.virgin)} tCO₂e</strong></div>
        <div><span class="muted small">Project total would become</span><strong class="warn">${tonnes(newTotal)} tCO₂e</strong></div>
      </div>
      <p class="muted small">That one sourcing choice adds <strong>${tonnes(c.extra)} tCO₂e</strong> (${times}× the rebar's carbon,
      +${p.total ? fmt((c.extra / p.total) * 100, 0) : 0}% on the whole project). Recycled content is one of the biggest single levers a specifier has.</p>`;
  }

  renderChart(p);
}

function renderChart(p) {
  const items = state.rows
    .map((r, i) => ({ label: r.element || getFactor(r.material)?.name || "Unnamed", v: p.results[i]?.total || 0 }))
    .filter(x => x.v > 0)
    .sort((x, y) => y.v - x.v);
  const chart = document.getElementById("chart");
  if (!items.length) { chart.innerHTML = `<p class="muted small">Nothing to chart yet.</p>`; return; }
  const max = items[0].v;
  chart.innerHTML = items.map(x => `
    <div class="bar-row">
      <span class="bar-label">${esc(x.label)}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${(x.v / max) * 100}%"></span></span>
      <span class="bar-val">${tonnes(x.v)}</span>
    </div>`).join("");
}

function renderFactors() {
  document.getElementById("factors-body").innerHTML = FACTORS.map(f => `
    <tr><td>${esc(f.name)}</td><td class="num">${f.density ? fmt(f.density) : "–"}</td><td class="num">${f.factor}</td>
    <td class="num">${Math.round(f.waste * 100)}%</td><td class="muted">${esc(f.note)}</td></tr>`).join("");
}

function renderAll() {
  renderRows();
  renderResults();
}

fillAssumptions();
renderFactors();
renderAll();
