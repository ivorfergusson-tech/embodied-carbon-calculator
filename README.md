# Embodied Carbon Calculator

Estimates the embodied carbon of a building's main structural and envelope
materials (life-cycle stages A1–A5), using ICE Database factors, and shows
how material choices such as recycled vs virgin steel change the footprint.

**Live version:** https://ivorfergusson-tech.github.io/embodied-carbon-calculator/

No install, no server. Open `index.html` in a browser. There's also an Excel
version with live formulas in `downloads/`.

## What embodied carbon is

*Operational* carbon is what a building emits while it's in use (heating,
lighting). *Embodied* carbon is what's released making, transporting and
building it in the first place. As buildings get more energy-efficient,
embodied carbon becomes a bigger share of their whole-life footprint. That's
why UK clients and contractors are increasingly asked to measure it for net-zero
targets and carbon reporting.

## The calculation

Each row of the take-off works through the same steps as the workbook's
Calculator tab:

1. **Mass incl. waste (kg)** = quantity × density (if taken off in m³) or
   quantity (if already in kg), × (1 + waste allowance). Real orders include
   offcuts and over-ordering, so carbon is worked out on what's bought.
2. **A1–A3 materials** = mass × ICE carbon factor (kgCO₂e/kg).
3. **A4 transport** = mass (tonnes) × distance (km) × transport factor
   (kgCO₂e per tonne·km, UK HGV road-freight average).
4. **A5 construction** = (A1–A3 + A4) × a site allowance %.
5. **Total** = A1–A3 + A4 + A5.

The page also re-runs any recycled rebar through the virgin-steel factor on
the same mass, so you can see what that one sourcing choice is worth in tCO₂e.

## How the code is organised

| File | Responsibility |
|---|---|
| `js/factors-data.js` | Reference data: material factors, densities and waste allowances (ICE Database), default project assumptions and the worked example. **Illustrative, not live data.** |
| `js/carbon.js` | The calculation. Pure functions, no DOM, so every number on screen traces back to a few lines here. |
| `js/app.js` | UI controller. Reads state, calls `carbon.js`, re-renders. Saves your take-off to the browser's `localStorage`. |
| `downloads/Embodied-Carbon-Calculator.xlsx` | The Excel version: live VLOOKUP formulas, material dropdowns, blank template rows, methodology and a summary chart. |
| `tools/carbon_calculator.py` | Python/openpyxl script that generates the workbook. |

## Known simplifications

- **Stages A1–A5 only** (BS EN 15978). It leaves out B (in-use maintenance and
  replacement) and C (end of life).
- **One transport distance and vehicle type** for every material. A real
  project would use each supplier's actual delivery details.
- **A5 is a flat % allowance**, because detailed site-energy data isn't
  usually available at early design stage. RICS whole life carbon guidance
  allows a default assumption in that case.
- **Timber excludes biogenic carbon** (CO₂ absorbed while the tree grew).
- **Factors are representative ICE figures**, not product-specific EPDs. Check
  the live ICE Database (v4.1) before relying on any number.
