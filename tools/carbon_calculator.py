"""
Builds a reusable embodied carbon calculator workbook (A1-A5: raw material,
transport and construction) covering any small construction project, using
representative factors based on the ICE Database (Inventory of Carbon and
Energy, University of Bath / Circular Ecology).

Unlike a script that just prints a final number, every result cell in the
"Calculator" sheet is a live Excel formula that references the "Carbon
Factors" reference table and the project-level assumptions cells -- open the
workbook and click any total to see exactly how it was derived.

The Calculator sheet ships pre-filled with a worked example (a small
single-storey domestic extension) followed by blank template rows with the
same formulas and a Material dropdown, so it can be reused for a different
project just by clearing/editing rows in Excel -- no need to re-run this
script or touch Python at all.

Usage:
    python tools/carbon_calculator.py [output_xlsx]

This is a deterministic formatting/writing tool -- the material factors are
fixed reference data (see the "Carbon Factors" sheet and workflows/
embodied_carbon_calculator.md for how to refresh them).
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_FILL = PatternFill(start_color="1F3B57", end_color="1F3B57", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F3B57")
SUBTITLE_FONT = Font(italic=True, size=10, color="555555")
SECTION_FONT = Font(bold=True, size=12, color="1F3B57")
NOTE_FONT = Font(italic=True, size=9, color="777777")
INPUT_FILL = PatternFill(start_color="FFF6DD", end_color="FFF6DD", fill_type="solid")
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
THIN_BORDER = Border(bottom=Side(style="thin", color="D9D9D9"))
TOTAL_FILL = PatternFill(start_color="E8EEF4", end_color="E8EEF4", fill_type="solid")
COMPARE_FILL = PatternFill(start_color="FBEFE0", end_color="FBEFE0", fill_type="solid")

# --- Reference data: ICE Database, Univ. of Bath / Circular Ecology -------
# Illustrative, commonly-cited cradle-to-gate (A1-A3, EN 15804+A2) values.
# Check the live database (circularecology.com/embodied-carbon-footprint-
# database.html, current edition v4.1) before using on a real project.
# Waste % = typical site over-order/offcut allowance (illustrative rule-of-
# thumb figures -- refer to WRAP UK guidance or your own take-off records
# for project-specific rates).
FACTORS = [
    # material, unit, density kg/m3 (None if N/A), A1-A3 factor kgCO2e/kg, waste %, source note
    ("Ready-mix concrete (general use, 30-40MPa)", "kg", 2400, 0.103, 0.05,
     "ICE DB, general concrete mix, cradle-to-gate A1-A3"),
    ("Reinforcement steel bar (rebar) - UK average (~66% recycled)", "kg", None, 0.456, 0.05,
     "ICE DB / UK steel sector average recycled content"),
    ("Reinforcement steel bar (rebar) - virgin (0% recycled)", "kg", None, 1.99, 0.05,
     "ICE DB, primary/virgin production - shown for comparison only"),
    ("Structural steel sections (hot rolled, virgin)", "kg", None, 1.55, 0.05,
     "ICE DB"),
    ("Common clay facing brick", "kg", 1700, 0.24, 0.10,
     "ICE DB, fired clay brick, ~2.3 kg per standard brick"),
    ("Concrete blockwork (dense aggregate)", "kg", 2000, 0.107, 0.08,
     "ICE DB"),
    ("General purpose cement mortar", "kg", None, 0.163, 0.10,
     "ICE DB"),
    ("Sawn softwood timber (structural, e.g. C16/C24)", "kg", 480, 0.31, 0.10,
     "ICE DB - excludes biogenic carbon storage, see Methodology"),
    ("Plasterboard", "kg", None, 0.39, 0.10,
     "ICE DB, gypsum plasterboard"),
    ("Mineral wool insulation", "kg", None, 1.28, 0.05,
     "ICE DB, glass/stone mineral wool"),
    ("Concrete roof tile", "kg", None, 0.098, 0.08,
     "ICE DB"),
    ("Float glass (double-glazed units)", "kg", None, 0.86, 0.03,
     "ICE DB, flat glass"),
]

# --- Worked example: small single-storey domestic side extension ----------
# element, material (must match FACTORS name exactly), quantity, quantity unit
ELEMENTS = [
    ("Foundation (strip footing)", "Ready-mix concrete (general use, 30-40MPa)", 4.5, "m3"),
    ("Ground floor slab", "Ready-mix concrete (general use, 30-40MPa)", 2.25, "m3"),
    ("Reinforcement (foundation + slab)", "Reinforcement steel bar (rebar) - UK average (~66% recycled)", 350, "kg"),
    ("Lintels over openings", "Structural steel sections (hot rolled, virgin)", 45, "kg"),
    ("External walls - outer leaf (facing brick)", "Common clay facing brick", 4140, "kg"),
    ("External walls - inner leaf (blockwork)", "Concrete blockwork (dense aggregate)", 3000, "kg"),
    ("Mortar (brick/block bedding & pointing)", "General purpose cement mortar", 500, "kg"),
    ("Roof structure (rafters, purlins, wall plates)", "Sawn softwood timber (structural, e.g. C16/C24)", 600, "kg"),
]
NUM_TEMPLATE_ROWS = 15  # extra blank, fully-formula'd rows so this can be reused for a new project

# Project-level assumption defaults (all editable cells in the Calculator sheet)
DEFAULT_PROJECT_NAME = "Small single-storey domestic side extension (~15m2 footprint)"
DEFAULT_DISTANCE_KM = 50
DEFAULT_TRANSPORT_FACTOR = 0.105  # kgCO2e per tonne.km, representative UK HGV road freight average
DEFAULT_A5_PCT = 0.05  # construction/site energy allowance, % of A1-A4, default assumption

CALC_HEADERS = [
    ("Element", 34), ("Material", 44), ("Quantity", 11), ("Unit", 7),
    ("Waste %", 9), ("Density (kg/m3)", 13), ("Mass incl. waste (kg)", 14),
    ("A1-A3 Factor (kgCO2e/kg)", 13), ("A1-A3 Carbon (kgCO2e)", 13),
    ("A4 Transport (kgCO2e)", 13), ("A5 Construction (kgCO2e)", 14),
    ("Total Carbon (kgCO2e)", 14), ("Total Carbon (tCO2e)", 13),
]
# Column letters, named for readability against CALC_HEADERS above
(C_ELEMENT, C_MATERIAL, C_QTY, C_UNIT, C_WASTE, C_DENSITY, C_MASS,
 C_A13_FACTOR, C_A13_CARBON, C_A4_CARBON, C_A5_CARBON, C_TOTAL_KG, C_TOTAL_T) = [
    get_column_letter(i) for i in range(1, len(CALC_HEADERS) + 1)
]


def _title_block(ws, title, subtitle, width_col="A", width=130):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUBTITLE_FONT
    ws.column_dimensions[width_col].width = width


def _header_row(ws, row, headers):
    for col_idx, (header, width) in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[row].height = 32


def _calc_row_formulas(calc, row, cf_last_row, distance_cell, transport_factor_cell, a5_pct_cell):
    """Write the (identical, unit-agnostic) formula set for one Calculator data row.

    Works whether the row is pre-filled or left blank as a template row -- every
    formula is wrapped in IFERROR so blank rows render as empty rather than #N/A.
    """
    cf_range = f"'Carbon Factors'!$A$5:$F${cf_last_row}"
    calc.cell(row=row, column=5,  # Waste %
              value=f"=IFERROR(VLOOKUP({C_MATERIAL}{row},{cf_range},5,FALSE),\"\")")
    calc.cell(row=row, column=6,  # Density
              value=f"=IFERROR(VLOOKUP({C_MATERIAL}{row},{cf_range},3,FALSE),\"\")")
    # Mass incl. waste (kg): volume x density for m3 rows, else quantity is already in kg;
    # then grossed up by the material's waste allowance.
    calc.cell(row=row, column=7,
              value=(f"=IFERROR(IF({C_UNIT}{row}=\"m3\",{C_QTY}{row}*{C_DENSITY}{row},{C_QTY}{row})"
                      f"*(1+{C_WASTE}{row}),\"\")"))
    calc.cell(row=row, column=8,  # A1-A3 factor
              value=f"=IFERROR(VLOOKUP({C_MATERIAL}{row},{cf_range},4,FALSE),\"\")")
    calc.cell(row=row, column=9,  # A1-A3 carbon
              value=f"=IFERROR({C_MASS}{row}*{C_A13_FACTOR}{row},\"\")")
    calc.cell(row=row, column=10,  # A4 transport carbon = tonnes x distance x factor
              value=f"=IFERROR(({C_MASS}{row}/1000)*{distance_cell}*{transport_factor_cell},\"\")")
    calc.cell(row=row, column=11,  # A5 construction/site energy, % of A1-A4
              value=f"=IFERROR(({C_A13_CARBON}{row}+{C_A4_CARBON}{row})*{a5_pct_cell},\"\")")
    calc.cell(row=row, column=12,  # Total kgCO2e
              value=f"=IFERROR({C_A13_CARBON}{row}+{C_A4_CARBON}{row}+{C_A5_CARBON}{row},\"\")")
    calc.cell(row=row, column=13,  # Total tCO2e
              value=f"=IFERROR({C_TOTAL_KG}{row}/1000,\"\")")
    for c in range(1, 14):
        calc.cell(row=row, column=c).border = THIN_BORDER
    calc.row_dimensions[row].height = 30


def build(output_path: str) -> None:
    wb = Workbook()

    # --- Read Me First ---
    tips = wb.active
    tips.title = "Read Me First"
    _title_block(tips, "Embodied Carbon Calculator", "Reusable template, pre-filled with a worked example")
    lines = [
        "",
        "What this is",
        "- Estimates the embodied carbon (A1-A3 material + A4 transport + A5 construction) of the "
        "main structural/envelope materials for a construction project.",
        "- Every total in the 'Calculator' tab is a live formula, not a typed-in number. Click any "
        "cell in columns E-M to see exactly how it was calculated and where the factor came from.",
        "",
        "Using it for a DIFFERENT project",
        "- The Calculator tab has 8 pre-filled rows (the worked example) followed by 15 blank "
        "template rows with the same live formulas already in place.",
        "- To model a new project: clear the Element/Quantity/Unit cells you don't need, or use the "
        "blank rows below. Pick the Material from the dropdown list (so it always matches the "
        "Carbon Factors tab exactly) and type the Unit as either 'm3' or 'kg'.",
        "- Update the project assumptions at the top of the Calculator tab (project name, transport "
        "distance, A5 allowance) for the new project.",
        "- Need more than 23 rows? Select a row inside the block, copy it, and insert/paste extra "
        "rows -- the formulas use fixed references to the Carbon Factors and assumption cells, so "
        "they copy down correctly.",
        "- If you add new materials, add them to the 'Carbon Factors' tab first (with a density if "
        "it's a volume-based material) so the dropdown and VLOOKUPs pick them up.",
        "",
        "Tabs in this workbook",
        "- Methodology: plain-English explanation of what embodied carbon means, the formula used, "
        "and the scope/limitations of this estimate.",
        "- Carbon Factors: the reference table of material carbon factors and waste allowances, with "
        "sources.",
        "- Calculator: the reusable worked example.",
        "- Summary: total result and a chart breaking it down by building element (worked example).",
        "",
        "Top-line caveat",
        "- The factors in this workbook are illustrative, representative figures based on the ICE "
        "Database (University of Bath / Circular Ecology) and UK Government transport conversion "
        "factors. For a real project, check current product-specific Environmental Product "
        "Declarations (EPDs) or the live ICE Database (v4.1) rather than relying on these numbers.",
    ]
    r = 3
    for line in lines:
        cell = tips.cell(row=r, column=1, value=line)
        if line and not line.startswith("-"):
            cell.font = SECTION_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    # --- Methodology ---
    meth = wb.create_sheet("Methodology")
    _title_block(meth, "Methodology", "How the numbers in this workbook are worked out")
    meth_lines = [
        "",
        "1. Embodied carbon vs operational carbon",
        "- Operational carbon is the CO2e a building produces while it's in use (heating, lighting, "
        "etc). Embodied carbon is the CO2e released making, transporting and building it in the "
        "first place - before anyone even moves in. This workbook only covers embodied carbon.",
        "",
        "2. The calculation",
        "- Mass incl. waste (kg) = Quantity x Density (for volume-based materials, e.g. concrete in "
        "m3) or just Quantity (for materials already taken off by mass, e.g. kg of steel) x "
        "(1 + Waste allowance %).",
        "- A1-A3 Carbon (kgCO2e) = Mass incl. waste x A1-A3 Carbon Factor.",
        "- A4 Transport Carbon (kgCO2e) = Mass (tonnes) x Transport distance (km) x Transport "
        "emission factor (kgCO2e per tonne.km).",
        "- A5 Construction Carbon (kgCO2e) = (A1-A3 + A4) x A5 allowance % - a default percentage "
        "uplift for on-site plant/energy use, used because granular construction-phase data isn't "
        "normally available at this stage (this is a recognised simplification - RICS whole life "
        "carbon guidance permits default assumptions for A5 when project-specific data is missing).",
        "- Total Carbon = A1-A3 + A4 + A5. Every one of these is a live formula in the 'Calculator' "
        "tab, referencing the 'Carbon Factors' tab and the project assumption cells at the top of "
        "the sheet - nothing is pasted in as a static number.",
        "",
        "3. Scope: what 'A1-A5' means (and what's still excluded)",
        "- Construction carbon assessments are broken into life-cycle stages (BS EN 15978). This "
        "workbook covers stages A1-A3 (raw material extraction, transport to factory, manufacture), "
        "A4 (transport to site) and A5 (construction/installation, via a default % allowance).",
        "- It does NOT include B (maintenance, repair, replacement while in use) or C (demolition, "
        "disposal at end of life). A full whole-life carbon assessment would need those stages too. "
        "This is a named limitation, not an oversight.",
        "- A4 uses one assumed average transport distance and a single UK HGV road-freight emission "
        "factor for every material - a real project would use each supplier's actual delivery "
        "distance and vehicle type.",
        "",
        "4. Waste allowance",
        "- Real quantity take-offs never use 100% of the material ordered - there's always cutting "
        "waste, breakages and over-ordering. Each material in the 'Carbon Factors' tab carries a "
        "typical waste allowance (e.g. 5% for concrete, 10% for brick/timber), which grosses up the "
        "mass before any carbon is calculated - so the estimate reflects what's actually bought, not "
        "just what ends up in the permanent structure.",
        "",
        "5. Where the factors come from",
        "- Material (A1-A3) carbon factors are based on the ICE Database (Inventory of Carbon and "
        "Energy), originally developed by Prof. Geoff Hammond and Dr Craig Jones at the University "
        "of Bath, now maintained by Circular Ecology - the standard free UK reference for embodied "
        "carbon, widely cited in industry.",
        "- The transport (A4) emission factor is a representative UK HGV road-freight average, in "
        "the style of the UK Government's (DESNZ/DEFRA) annual greenhouse gas conversion factors for "
        "company reporting.",
        "- All figures here are representative/typical values, not product-specific EPD or logistics "
        "data. For a real project you'd use the actual supplier's EPD and delivery details.",
        "",
        "6. Why the steel sourcing matters",
        "- The 'Carbon Factors' tab lists two rebar factors: UK-average recycled steel (~0.46 "
        "kgCO2e/kg) and virgin steel (~1.99 kgCO2e/kg) - over 4x higher. The 'Comparison' section on "
        "the Calculator tab re-runs the same reinforcement quantity through the virgin factor so you "
        "can see the difference in real kgCO2e, not just percentages. This is one of the biggest "
        "single levers a specifier has over a project's embodied carbon.",
        "",
        "7. Timber caveat (biogenic carbon)",
        "- Growing timber absorbs (sequesters) CO2. Some carbon accounting methods show this as a "
        "negative figure at A1-A3 and then add it back at end of life when the timber decays or is "
        "burned; others exclude it entirely, which is the convention used here. In practice this "
        "means timber's true climate impact is more favourable than the flat 0.31 kgCO2e/kg figure "
        "alone suggests - worth knowing if asked about it.",
    ]
    r = 3
    for line in meth_lines:
        cell = meth.cell(row=r, column=1, value=line)
        if line and not line.startswith("-"):
            cell.font = SECTION_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    # --- Carbon Factors ---
    cf = wb.create_sheet("Carbon Factors")
    _title_block(cf, "Carbon Factors (Reference Table)",
                 "Representative A1-A3 factors + typical waste allowances, ICE Database (Univ. of Bath / Circular Ecology)",
                 width=46)
    cf_headers = [("Material", 46), ("Unit", 8), ("Density (kg/m3)", 15),
                  ("Carbon Factor A1-A3 (kgCO2e/kg)", 15), ("Typical Waste Allowance (%)", 14),
                  ("Source / Notes", 50)]
    _header_row(cf, 4, cf_headers)
    for i, (material, unit, density, factor, waste, note) in enumerate(FACTORS, start=5):
        cf.cell(row=i, column=1, value=material).alignment = WRAP
        cf.cell(row=i, column=2, value=unit)
        cf.cell(row=i, column=3, value=density if density is not None else "-")
        cf.cell(row=i, column=4, value=factor)
        cf.cell(row=i, column=5, value=waste)
        cf.cell(row=i, column=5).number_format = "0%"
        cf.cell(row=i, column=6, value=note).alignment = WRAP
        for c in range(1, 7):
            cf.cell(row=i, column=c).border = THIN_BORDER
        cf.row_dimensions[i].height = 30
    cf_last_row = 4 + len(FACTORS)

    # --- Calculator ---
    calc = wb.create_sheet("Calculator")
    _title_block(calc, "Calculator",
                 "Reusable template - pre-filled with a worked example (small single-storey domestic extension)",
                 width=34)

    # Project-level assumptions (editable input cells, referenced by every row's formulas below)
    assumptions = [
        ("Project name:", DEFAULT_PROJECT_NAME),
        ("Average transport distance to site (km):", DEFAULT_DISTANCE_KM),
        ("Transport emission factor (kgCO2e per tonne.km - UK HGV road avg):", DEFAULT_TRANSPORT_FACTOR),
        ("A5 construction/site energy allowance (% of A1-A4):", DEFAULT_A5_PCT),
    ]
    for i, (label, value) in enumerate(assumptions):
        row = 4 + i
        calc.cell(row=row, column=1, value=label).font = Font(bold=True)
        cell = calc.cell(row=row, column=3, value=value)
        cell.fill = INPUT_FILL
        cell.border = Border(bottom=Side(style="thin", color="C9A227"), top=Side(style="thin", color="C9A227"),
                              left=Side(style="thin", color="C9A227"), right=Side(style="thin", color="C9A227"))
    distance_cell = "$C$5"
    transport_factor_cell = "$C$6"
    a5_pct_cell = "$C$7"
    calc["C7"].number_format = "0%"
    calc.cell(row=8, column=1,
              value="(Yellow cells are editable assumptions - change them here and every row below recalculates.)").font = NOTE_FONT

    note_row = 10
    calc.cell(row=note_row, column=1,
              value="Worked example rows below, followed by blank template rows for a new project. "
                    "Material has a dropdown list; Unit must be 'm3' or 'kg'.").font = NOTE_FONT
    calc.cell(row=note_row, column=1).alignment = WRAP

    header_row = note_row + 1
    _header_row(calc, header_row, CALC_HEADERS)

    first_data_row = header_row + 1
    for offset, (element, material, qty, unit) in enumerate(ELEMENTS):
        row = first_data_row + offset
        calc.cell(row=row, column=1, value=element).alignment = WRAP
        calc.cell(row=row, column=2, value=material).alignment = WRAP
        calc.cell(row=row, column=3, value=qty)
        calc.cell(row=row, column=4, value=unit)
        _calc_row_formulas(calc, row, cf_last_row, distance_cell, transport_factor_cell, a5_pct_cell)

    template_first_row = first_data_row + len(ELEMENTS)
    for offset in range(NUM_TEMPLATE_ROWS):
        row = template_first_row + offset
        _calc_row_formulas(calc, row, cf_last_row, distance_cell, transport_factor_cell, a5_pct_cell)

    last_data_row = template_first_row + NUM_TEMPLATE_ROWS - 1

    # Material dropdown + Unit dropdown across every data row (filled and blank template rows)
    material_dv = DataValidation(type="list", formula1=f"='Carbon Factors'!$A$5:$A${cf_last_row}",
                                  allow_blank=True, showDropDown=False)
    material_dv.error = "Pick a material from the Carbon Factors tab so it matches exactly."
    material_dv.errorTitle = "Unknown material"
    calc.add_data_validation(material_dv)
    material_dv.add(f"B{first_data_row}:B{last_data_row}")

    unit_dv = DataValidation(type="list", formula1='"m3,kg"', allow_blank=True, showDropDown=False)
    calc.add_data_validation(unit_dv)
    unit_dv.add(f"D{first_data_row}:D{last_data_row}")

    total_row = last_data_row + 1
    calc.cell(row=total_row, column=1, value="TOTAL (A1-A5 embodied carbon)").font = Font(bold=True)
    calc.cell(row=total_row, column=12, value=f"=SUM(L{first_data_row}:L{last_data_row})").font = Font(bold=True)
    calc.cell(row=total_row, column=13, value=f"=SUM(M{first_data_row}:M{last_data_row})").font = Font(bold=True)
    for c in range(1, 14):
        calc.cell(row=total_row, column=c).fill = TOTAL_FILL

    # --- Comparison: same reinforcement mass, virgin steel instead of UK-average recycled ---
    compare_row = total_row + 2
    calc.cell(row=compare_row, column=1,
               value="COMPARISON (not included in total): same reinforcement mass, virgin steel instead of UK-average recycled").font = SECTION_FONT
    header2_row = compare_row + 1
    _header_row(calc, header2_row, CALC_HEADERS)
    data2_row = header2_row + 1
    rebar_row = first_data_row + 2  # "Reinforcement (foundation + slab)" row above
    calc.cell(row=data2_row, column=1, value="Reinforcement (foundation + slab) - IF VIRGIN STEEL").alignment = WRAP
    calc.cell(row=data2_row, column=2, value="Reinforcement steel bar (rebar) - virgin (0% recycled)").alignment = WRAP
    calc.cell(row=data2_row, column=3, value=f"=C{rebar_row}")
    calc.cell(row=data2_row, column=4, value="kg")
    _calc_row_formulas(calc, data2_row, cf_last_row, distance_cell, transport_factor_cell, a5_pct_cell)
    for c in range(1, 14):
        calc.cell(row=data2_row, column=c).fill = COMPARE_FILL

    diff_row = data2_row + 1
    calc.cell(row=diff_row, column=1, value="Difference vs. UK-average recycled steel (same mass)").font = Font(italic=True)
    calc.cell(row=diff_row, column=12, value=f"=L{data2_row}-L{rebar_row}").font = Font(italic=True)
    calc.cell(row=diff_row, column=13, value=f"=M{data2_row}-M{rebar_row}").font = Font(italic=True)

    # --- Summary ---
    summ = wb.create_sheet("Summary")
    _title_block(summ, "Summary", "Total result and breakdown by building element (worked example)", width=40)
    summ["A4"] = "Total embodied carbon (A1-A5)"
    summ["A4"].font = SECTION_FONT
    summ["B4"] = f"=Calculator!M{total_row}"
    summ["B4"].font = Font(bold=True, size=14, color="1F3B57")
    summ["A5"] = "tCO2e"
    summ["A5"].font = NOTE_FONT
    summ["A7"] = "If reinforcement had used virgin steel instead of UK-average recycled steel:"
    summ["A7"].alignment = WRAP
    summ["B7"] = f"=Calculator!M{total_row}+Calculator!M{diff_row}"
    summ["B7"].font = Font(bold=True, color="B35B00")
    summ["A8"] = "tCO2e total (the extra carbon from that one sourcing choice is broken out in the 'Difference' row on the Calculator tab)"
    summ["A8"].font = NOTE_FONT
    summ["A8"].alignment = WRAP
    summ.column_dimensions["A"].width = 46
    summ.column_dimensions["B"].width = 16

    # Breakdown table feeding the chart (element name + tCO2e), copied by formula from Calculator
    # (references the 8 worked-example rows only; extend this range manually if you add rows for a new project)
    summ_header_row = 11
    summ.cell(row=summ_header_row, column=1, value="Element").font = HEADER_FONT
    summ.cell(row=summ_header_row, column=1).fill = HEADER_FILL
    summ.cell(row=summ_header_row, column=2, value="Embodied Carbon (tCO2e)").font = HEADER_FONT
    summ.cell(row=summ_header_row, column=2).fill = HEADER_FILL
    for offset in range(len(ELEMENTS)):
        row = summ_header_row + 1 + offset
        src_row = first_data_row + offset
        summ.cell(row=row, column=1, value=f"=Calculator!A{src_row}")
        summ.cell(row=row, column=2, value=f"=Calculator!M{src_row}")
    summ_last_row = summ_header_row + len(ELEMENTS)
    summ.cell(row=summ_last_row + 1, column=1,
               value="(Chart covers the worked example rows only - extend the data range in Excel's "
                     "'Select Data' if you add rows for a new project.)").font = NOTE_FONT
    summ.cell(row=summ_last_row + 1, column=1).alignment = WRAP

    chart = BarChart()
    chart.title = "Embodied carbon by building element (tCO2e, A1-A5)"
    chart.y_axis.title = "tCO2e"
    chart.x_axis.title = "Element"
    chart.style = 10
    data_ref = Reference(summ, min_col=2, min_row=summ_header_row, max_row=summ_last_row)
    cats_ref = Reference(summ, min_col=1, min_row=summ_header_row + 1, max_row=summ_last_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.height = 10
    chart.width = 24
    summ.add_chart(chart, "D4")

    wb.save(output_path)
    print(f"Wrote embodied carbon calculator to {output_path}")
    print(f"Worked example materials: {len(ELEMENTS)} | Blank template rows: {NUM_TEMPLATE_ROWS} | Reference factors: {len(FACTORS)}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Embodied Carbon Calculator.xlsx"
    build(out)
