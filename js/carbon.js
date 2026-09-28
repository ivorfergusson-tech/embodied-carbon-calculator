// The calculation itself — the same formulas as each row of the workbook's
// Calculator tab. No DOM access here, so every number on screen can be traced
// back to these few lines.

// Mass incl. waste (kg): volume x density for m3 rows, else quantity is
// already in kg; then grossed up by the material's waste allowance.
function massInclWaste(row, mat) {
  const qty = Number(row.qty) || 0;
  if (row.unit === "m3") {
    if (!mat.density) return null; // can't convert a volume without a density
    return qty * mat.density * (1 + mat.waste);
  }
  return qty * (1 + mat.waste);
}

// Returns every stage for one row, in kgCO2e, or null if the row is incomplete.
function calcRow(row, assumptions, materialOverride) {
  const mat = getFactor(materialOverride || row.material);
  if (!mat) return null;
  const mass = massInclWaste(row, mat);
  if (mass === null) return null;

  const a13 = mass * mat.factor;                                                   // A1-A3 material
  const a4 = (mass / 1000) * assumptions.distanceKm * assumptions.transportFactor; // A4 transport: t x km x factor
  const a5 = (a13 + a4) * (assumptions.a5Pct / 100);                               // A5 construction allowance
  return { mat, mass, a13, a4, a5, total: a13 + a4 + a5 };
}

function calcProject(rows, assumptions) {
  const results = rows.map(r => calcRow(r, assumptions));
  const sum = key => results.reduce((s, r) => s + (r ? r[key] : 0), 0);
  return { results, a13: sum("a13"), a4: sum("a4"), a5: sum("a5"), total: sum("total") };
}

// Re-runs any recycled-steel rows through the virgin factor (same mass).
function calcVirginComparison(rows, assumptions) {
  let recycled = 0, virgin = 0, count = 0;
  rows.forEach(r => {
    const swap = VIRGIN_SWAPS[r.material];
    if (!swap) return;
    const a = calcRow(r, assumptions);
    const b = calcRow(r, assumptions, swap);
    if (!a || !b) return;
    recycled += a.total; virgin += b.total; count++;
  });
  return { count, recycled, virgin, extra: virgin - recycled };
}
