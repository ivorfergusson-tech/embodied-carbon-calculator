// Reference data: ICE Database (Inventory of Carbon and Energy),
// University of Bath / Circular Ecology.
// Illustrative, commonly-cited cradle-to-gate (A1-A3, EN 15804+A2) values —
// the same figures as the "Carbon Factors" tab in the Excel workbook.
// Check the live database (v4.1) or a product EPD before using on a real project.
//
// density: kg/m3, only for materials normally taken off by volume (null = take off in kg)
// factor:  A1-A3 kgCO2e per kg
// waste:   typical site over-order/offcut allowance (fraction)

const FACTORS = [
  { id: "concrete",   name: "Ready-mix concrete (general use, 30-40MPa)",                  density: 2400, factor: 0.103, waste: 0.05, note: "ICE DB, general concrete mix, cradle-to-gate A1-A3" },
  { id: "rebar-uk",   name: "Reinforcement steel bar (rebar) - UK average (~66% recycled)", density: null, factor: 0.456, waste: 0.05, note: "ICE DB / UK steel sector average recycled content" },
  { id: "rebar-virgin", name: "Reinforcement steel bar (rebar) - virgin (0% recycled)",     density: null, factor: 1.99,  waste: 0.05, note: "ICE DB, primary/virgin production" },
  { id: "steel-sections", name: "Structural steel sections (hot rolled, virgin)",           density: null, factor: 1.55,  waste: 0.05, note: "ICE DB" },
  { id: "brick",      name: "Common clay facing brick",                                     density: 1700, factor: 0.24,  waste: 0.10, note: "ICE DB, fired clay brick, ~2.3 kg per standard brick" },
  { id: "block",      name: "Concrete blockwork (dense aggregate)",                         density: 2000, factor: 0.107, waste: 0.08, note: "ICE DB" },
  { id: "mortar",     name: "General purpose cement mortar",                                density: null, factor: 0.163, waste: 0.10, note: "ICE DB" },
  { id: "timber",     name: "Sawn softwood timber (structural, e.g. C16/C24)",              density: 480,  factor: 0.31,  waste: 0.10, note: "ICE DB — excludes biogenic carbon storage" },
  { id: "plasterboard", name: "Plasterboard",                                               density: null, factor: 0.39,  waste: 0.10, note: "ICE DB, gypsum plasterboard" },
  { id: "mineral-wool", name: "Mineral wool insulation",                                    density: null, factor: 1.28,  waste: 0.05, note: "ICE DB, glass/stone mineral wool" },
  { id: "roof-tile",  name: "Concrete roof tile",                                           density: null, factor: 0.098, waste: 0.08, note: "ICE DB" },
  { id: "glass",      name: "Float glass (double-glazed units)",                            density: null, factor: 0.86,  waste: 0.03, note: "ICE DB, flat glass" },
];

// Recycled -> virgin swaps used by the "what if virgin steel?" comparison
const VIRGIN_SWAPS = { "rebar-uk": "rebar-virgin" };

const DEFAULT_ASSUMPTIONS = {
  projectName: "Small single-storey domestic side extension (~15m² footprint)",
  distanceKm: 50,
  transportFactor: 0.105, // kgCO2e per tonne.km, representative UK HGV road freight average
  a5Pct: 5,               // construction/site energy allowance, % of A1-A4
};

// Worked example: the same 8 rows that ship pre-filled in the workbook
const SAMPLE_ELEMENTS = [
  { element: "Foundation (strip footing)",                     material: "concrete",       qty: 4.5,  unit: "m3" },
  { element: "Ground floor slab",                              material: "concrete",       qty: 2.25, unit: "m3" },
  { element: "Reinforcement (foundation + slab)",              material: "rebar-uk",       qty: 350,  unit: "kg" },
  { element: "Lintels over openings",                          material: "steel-sections", qty: 45,   unit: "kg" },
  { element: "External walls - outer leaf (facing brick)",     material: "brick",          qty: 4140, unit: "kg" },
  { element: "External walls - inner leaf (blockwork)",        material: "block",          qty: 3000, unit: "kg" },
  { element: "Mortar (brick/block bedding & pointing)",        material: "mortar",         qty: 500,  unit: "kg" },
  { element: "Roof structure (rafters, purlins, wall plates)", material: "timber",         qty: 600,  unit: "kg" },
];

function getFactor(id) {
  return FACTORS.find(f => f.id === id) || null;
}
