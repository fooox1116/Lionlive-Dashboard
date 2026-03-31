# Colombia Team Tracker Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate `Colombia_Teams.xlsx` — a fully-wired Excel workbook for tracking monthly unit economics of Colombia guild teams, with configurable staff allocation, actual-value overrides, and investment recovery tracking.

**Architecture:** Python script (`generate_colombia_tracker.py`) builds the entire workbook with openpyxl. All inter-sheet formulas are written as Excel formula strings. User edits `Monthly_Data` and `Staff` directly in Excel after generation; all other sheets recalculate automatically.

**Tech Stack:** Python 3, openpyxl, Excel (xlsx format)

**Spec:** `docs/superpowers/specs/2026-03-31-colombia-team-tracker-design.md`

---

## Chunk 1: Script scaffold + Config sheet

### Task 1: Script skeleton

**Files:**
- Create: `generate_colombia_tracker.py`

- [ ] Create script with imports and constants

```python
#!/usr/bin/env python3
"""Generate Colombia_Teams.xlsx — Lionlive guild team tracker."""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
import datetime

OUTPUT_FILE = "Colombia_Teams.xlsx"

# ── Colour palette ────────────────────────────────────────────────
FILL_HEADER    = PatternFill("solid", fgColor="1F3864")   # dark navy
FILL_SECTION   = PatternFill("solid", fgColor="2E75B6")   # mid blue
FILL_LABEL     = PatternFill("solid", fgColor="D6E4F7")   # light blue
FILL_ENTRY     = PatternFill("solid", fgColor="FFF2CC")   # yellow — data entry
FILL_OVERRIDE  = PatternFill("solid", fgColor="DDEEFF")   # blue — actual overrides
FILL_CALC      = PatternFill("solid", fgColor="F2F2F2")   # grey — auto-calc
FILL_WHITE     = PatternFill("solid", fgColor="FFFFFF")

FONT_HDR  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
FONT_BOLD = Font(name="Arial", bold=True, size=10)
FONT_NORM = Font(name="Arial", size=10)
FONT_TINY = Font(name="Arial", size=9, color="666666")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")
ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center")

THIN = Side(style="thin", color="CCCCCC")
MED  = Side(style="medium", color="2E75B6")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ── Team definitions ──────────────────────────────────────────────
TEAMS = [
    {"id": "COL01", "name": "Team 1", "market": "Colombia",
     "launch": "2025-09-01", "status": "Active"},
    {"id": "COL02", "name": "Team 2", "market": "Colombia",
     "launch": "2026-02-01", "status": "Active"},
]

# ── Staff definitions ─────────────────────────────────────────────
# teams_covered: list of team IDs this person covers
STAFF = [
    {"name": "Ops Manager",    "role": "Operations",    "salary": 800,  "teams": ["COL01", "COL02"]},
    {"name": "Dance Teacher",  "role": "Dance Teacher", "salary": 600,  "teams": ["COL01"]},
    {"name": "MUA",            "role": "MUA",           "salary": 500,  "teams": ["COL01", "COL02"]},
    {"name": "Tech Support",   "role": "Tech",          "salary": 400,  "teams": ["COL01", "COL02"]},
]

# ── Sample months (extend as needed) ─────────────────────────────
MONTHS = [
    "2026-02", "2026-03",
]

def main():
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    build_config(wb)
    build_staff(wb)
    build_monthly_data(wb)
    build_ue_summary(wb)
    for team in TEAMS:
        build_team_sheet(wb, team)

    wb.save(OUTPUT_FILE)
    print(f"✅  Saved {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
```

- [ ] Run to confirm no import errors:
  ```bash
  python3 generate_colombia_tracker.py
  ```
  Expected: `✅  Saved Colombia_Teams.xlsx` (empty sheets)

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: scaffold colombia team tracker generator"
  ```

---

### Task 2: Config sheet

**Files:**
- Modify: `generate_colombia_tracker.py` — add `build_config(wb)` function

- [ ] Add helper utilities above `main()`:

```python
def h(ws, row, col, value, fill=None, font=None, align=None, fmt=None, border=None):
    """Write a cell with optional styling."""
    c = ws.cell(row=row, column=col, value=value)
    if fill:   c.fill   = fill
    if font:   c.font   = font
    if align:  c.alignment = align
    if fmt:    c.number_format = fmt
    if border: c.border = border
    return c

def section_header(ws, row, col, text, width=2):
    """Write a section header spanning `width` columns."""
    c = ws.cell(row=row, column=col, value=text)
    c.fill  = FILL_SECTION
    c.font  = FONT_HDR
    c.alignment = ALIGN_CENTER
    if width > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + width - 1)
```

- [ ] Add `build_config(wb)` function:

```python
def build_config(wb):
    ws = wb.create_sheet("Config")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    # ── Title ─────────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "⚙️  CONFIG — Lionlive Colombia Guild Tracker"
    c.fill  = FILL_HEADER
    c.font  = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28

    # ── Section A: Team Registry ───────────────────────────────────
    section_header(ws, 3, 1, "TEAM REGISTRY", width=5)
    headers = ["Team ID", "Team Name", "Market", "Launch Date", "Status"]
    for i, h_text in enumerate(headers, 1):
        c = ws.cell(row=4, column=i, value=h_text)
        c.fill = FILL_LABEL; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN

    for r, team in enumerate(TEAMS, 5):
        data = [team["id"], team["name"], team["market"], team["launch"], team["status"]]
        for i, val in enumerate(data, 1):
            c = ws.cell(row=r, column=i, value=val)
            c.fill = FILL_WHITE; c.font = FONT_NORM
            c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

    # ── Section B: UE Assumptions ─────────────────────────────────
    row = 5 + len(TEAMS) + 2
    section_header(ws, row, 1, "UE ASSUMPTIONS (edit here → auto-applied to all sheets)", width=3)
    row += 1

    assumptions = [
        ("Take Rate",             0.65,   "0%",    "CONFIG_TAKE_RATE"),
        ("Diamond → USD rate",    0.01,   "$0.0000","CONFIG_DIAMOND_RATE"),
        ("Monthly Rent (USD)",    600,    "$#,##0", "CONFIG_RENT"),
        ("Monthly Utility (USD)", 800,    "$#,##0", "CONFIG_UTILITY"),
        ("Monthly Buffer (USD)",  1000,   "$#,##0", "CONFIG_BUFFER"),
        ("Initial Investment",    11430,  "$#,##0", "CONFIG_INITIAL_INVEST"),
        ("Creator Base Salary",   500,    "$#,##0", "CONFIG_CREATOR_BASE"),
        ("Revenue Share %",       0.25,   "0%",    "CONFIG_CREATOR_REVSHARE"),
    ]
    for label, val, fmt, name in assumptions:
        c_lbl = ws.cell(row=row, column=1, value=label)
        c_lbl.fill = FILL_LABEL; c_lbl.font = FONT_BOLD
        c_lbl.alignment = ALIGN_LEFT; c_lbl.border = BORDER_THIN

        c_val = ws.cell(row=row, column=2, value=val)
        c_val.fill = FILL_ENTRY; c_val.font = FONT_NORM
        c_val.number_format = fmt; c_val.alignment = ALIGN_CENTER
        c_val.border = BORDER_THIN

        # Define named range so other sheets can reference Config!B<row>
        ws.cell(row=row, column=3, value=f"← Named: {name}").font = FONT_TINY
        wb.defined_names[name] = openpyxl.workbook.defined_name.DefinedName(
            name, attr_text=f"Config!$B${row}"
        )
        row += 1

    # ── Section C: Stage Thresholds ───────────────────────────────
    row += 1
    section_header(ws, row, 1, "STAGE THRESHOLDS (diamonds/month)", width=3)
    row += 1
    stages = [
        ("Incubation",  0,         "CONFIG_STAGE_INC"),
        ("Breakeven",   500000,    "CONFIG_STAGE_BE"),
        ("Stable",      1060000,   "CONFIG_STAGE_STB"),
        ("Optimistic",  2000000,   "CONFIG_STAGE_OPT"),
    ]
    ws.cell(row=row, column=1, value="Stage").fill = FILL_LABEL
    ws.cell(row=row, column=1).font = FONT_BOLD
    ws.cell(row=row, column=2, value="Min Diamonds").fill = FILL_LABEL
    ws.cell(row=row, column=2).font = FONT_BOLD
    row += 1
    for label, val, name in stages:
        ws.cell(row=row, column=1, value=label).font = FONT_NORM
        c_val = ws.cell(row=row, column=2, value=val)
        c_val.fill = FILL_ENTRY; c_val.font = FONT_NORM
        c_val.number_format = "#,##0"
        wb.defined_names[name] = openpyxl.workbook.defined_name.DefinedName(
            name, attr_text=f"Config!$B${row}"
        )
        row += 1

    # ── Legend ────────────────────────────────────────────────────
    row += 1
    section_header(ws, row, 1, "LEGEND", width=2)
    row += 1
    legend = [
        (FILL_ENTRY,    "Yellow = Data entry required"),
        (FILL_OVERRIDE, "Blue   = Actual override (leave blank → use assumption)"),
        (FILL_CALC,     "Grey   = Auto-calculated, do not edit"),
    ]
    for fill, label in legend:
        ws.cell(row=row, column=1).fill = fill
        ws.cell(row=row, column=2, value=label).font = FONT_TINY
        row += 1
```

- [ ] Run script and open `Colombia_Teams.xlsx` — verify Config sheet looks correct with assumptions table and team registry.

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add Config sheet to colombia tracker"
  ```

---

## Chunk 2: Staff sheet + Monthly_Data sheet

### Task 3: Staff sheet

**Files:**
- Modify: `generate_colombia_tracker.py` — add `build_staff(wb)`

- [ ] Add `build_staff(wb)` function:

```python
def build_staff(wb):
    ws = wb.create_sheet("Staff")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16

    # Dynamic columns: one per team, then Teams Covered, Coeff, then cost per team
    team_ids = [t["id"] for t in TEAMS]
    n_teams  = len(team_ids)

    col_first_team = 4  # D onward: one column per team (checkmark)
    col_teams_covered = col_first_team + n_teams
    col_coeff         = col_teams_covered + 1
    col_cost_start    = col_coeff + 1   # one cost col per team

    # Set column widths for team check columns
    for i in range(n_teams):
        ws.column_dimensions[get_column_letter(col_first_team + i)].width = 10
    ws.column_dimensions[get_column_letter(col_teams_covered)].width = 14
    ws.column_dimensions[get_column_letter(col_coeff)].width = 14
    for i in range(n_teams):
        ws.column_dimensions[get_column_letter(col_cost_start + i)].width = 14

    # Title
    total_cols = col_cost_start + n_teams - 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(row=1, column=1, value="👥  STAFF — Cost Allocation by Team")
    c.fill = FILL_HEADER; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 26

    # Instructions row
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    c = ws.cell(row=2, column=1,
        value='✏️  Mark "✓" in team columns to assign staff. Allocation & cost auto-calculate.')
    c.font = Font(name="Arial", italic=True, size=9, color="444444")
    c.alignment = ALIGN_LEFT

    # Headers row 3
    headers_fixed = ["Staff Name", "Role", "Salary (USD)"]
    for i, hdr in enumerate(headers_fixed, 1):
        c = ws.cell(row=3, column=i, value=hdr)
        c.fill = FILL_LABEL; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

    for i, tid in enumerate(team_ids):
        col = col_first_team + i
        c = ws.cell(row=3, column=col, value=tid)
        c.fill = FILL_LABEL; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

    for col, label in [
        (col_teams_covered, "Teams Covered"),
        (col_coeff,         "Alloc Coeff"),
    ]:
        c = ws.cell(row=3, column=col, value=label)
        c.fill = FILL_CALC; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

    for i, tid in enumerate(team_ids):
        col = col_cost_start + i
        c = ws.cell(row=3, column=col, value=f"{tid} Cost")
        c.fill = FILL_CALC; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

    # Data rows
    for r_offset, staff in enumerate(STAFF):
        row = 4 + r_offset
        # Fixed columns
        ws.cell(row=row, column=1, value=staff["name"]).font = FONT_NORM
        ws.cell(row=row, column=2, value=staff["role"]).font  = FONT_NORM
        c_sal = ws.cell(row=row, column=3, value=staff["salary"])
        c_sal.fill = FILL_ENTRY; c_sal.font = FONT_NORM
        c_sal.number_format = "$#,##0"; c_sal.alignment = ALIGN_RIGHT
        c_sal.border = BORDER_THIN

        # Team checkmark columns
        check_cols = []
        for i, tid in enumerate(team_ids):
            col = col_first_team + i
            val = "✓" if tid in staff["teams"] else ""
            c = ws.cell(row=row, column=col, value=val)
            c.fill = FILL_ENTRY; c.font = FONT_NORM
            c.alignment = ALIGN_CENTER; c.border = BORDER_THIN
            check_cols.append(get_column_letter(col))

        # Teams Covered formula: COUNTIF across checkmark columns
        check_range = ",".join(f"{cl}{row}" for cl in check_cols)
        coeff_col_letter = get_column_letter(col_teams_covered)
        covered_formula = f'=COUNTIF({check_cols[0]}{row}:{check_cols[-1]}{row},"✓")'
        c_cov = ws.cell(row=row, column=col_teams_covered, value=covered_formula)
        c_cov.fill = FILL_CALC; c_cov.font = FONT_NORM
        c_cov.number_format = "0"; c_cov.alignment = ALIGN_CENTER
        c_cov.border = BORDER_THIN

        # Allocation Coefficient: 1/Teams_Covered (0 if 0)
        c_coeff = ws.cell(row=row, column=col_coeff,
            value=f"=IF({coeff_col_letter}{row}=0,0,1/{coeff_col_letter}{row})")
        c_coeff.fill = FILL_CALC; c_coeff.font = FONT_NORM
        c_coeff.number_format = "0.00"; c_coeff.alignment = ALIGN_CENTER
        c_coeff.border = BORDER_THIN

        # Per-team cost = Salary × Coeff × IF(team assigned, 1, 0)
        coeff_letter = get_column_letter(col_coeff)
        sal_letter   = get_column_letter(3)
        for i, tid in enumerate(team_ids):
            col = col_cost_start + i
            team_check_col = get_column_letter(col_first_team + i)
            formula = (f'=IF({team_check_col}{row}="✓",'
                       f'{sal_letter}{row}*{coeff_letter}{row},0)')
            c_cost = ws.cell(row=row, column=col, value=formula)
            c_cost.fill = FILL_CALC; c_cost.font = FONT_NORM
            c_cost.number_format = "$#,##0.00"; c_cost.alignment = ALIGN_RIGHT
            c_cost.border = BORDER_THIN

    # Totals row
    total_row = 4 + len(STAFF)
    ws.cell(row=total_row, column=1, value="TOTAL").font = FONT_BOLD
    c_sal_tot = ws.cell(row=total_row, column=3,
        value=f"=SUM(C4:C{total_row-1})")
    c_sal_tot.number_format = "$#,##0"; c_sal_tot.font = FONT_BOLD
    c_sal_tot.fill = FILL_LABEL; c_sal_tot.border = BORDER_THIN

    for i in range(n_teams):
        col = col_cost_start + i
        col_letter = get_column_letter(col)
        c_tot = ws.cell(row=total_row, column=col,
            value=f"=SUM({col_letter}4:{col_letter}{total_row-1})")
        c_tot.number_format = "$#,##0.00"; c_tot.font = FONT_BOLD
        c_tot.fill = FILL_LABEL; c_tot.border = BORDER_THIN

    # Store total row and cost column info as sheet attributes for cross-ref
    ws._staff_total_row    = total_row
    ws._col_cost_start     = col_cost_start
    ws._col_first_team_chk = col_first_team
```

- [ ] Run and verify Staff sheet renders with checkmarks, formulas, and totals.

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add Staff sheet with auto allocation formulas"
  ```

---

### Task 4: Monthly_Data sheet

**Files:**
- Modify: `generate_colombia_tracker.py` — add `build_monthly_data(wb)`

- [ ] Add `build_monthly_data(wb)`:

```python
# Column layout constants for Monthly_Data
# For each team: diamonds, creator count, creator salary override,
#                rent override, utility override, buffer override, other override
TEAM_COLS = 7  # columns per team block
MONTH_COL = 1  # column A

def monthly_data_team_col_start(team_index):
    """Column index (1-based) for start of team block in Monthly_Data."""
    return MONTH_COL + 1 + team_index * TEAM_COLS

def build_monthly_data(wb):
    ws = wb.create_sheet("Monthly_Data")
    team_ids   = [t["id"] for t in TEAMS]
    n_teams    = len(team_ids)

    # Admin cols start after all team blocks
    admin_col_start = MONTH_COL + 1 + n_teams * TEAM_COLS

    ws.column_dimensions["A"].width = 12

    # ── Title ─────────────────────────────────────────────────────
    last_col = admin_col_start + 4
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws["A1"]
    c.value = "📋  MONTHLY DATA — Enter actual values here"
    c.fill = FILL_HEADER; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER; ws.row_dimensions[1].height = 26

    # ── Team section headers (row 2) ───────────────────────────────
    for ti, tid in enumerate(team_ids):
        col_s = monthly_data_team_col_start(ti)
        ws.merge_cells(start_row=2, start_column=col_s,
                       end_row=2, end_column=col_s + TEAM_COLS - 1)
        c = ws.cell(row=2, column=col_s, value=f"── {tid} ──")
        c.fill = FILL_SECTION; c.font = FONT_HDR; c.alignment = ALIGN_CENTER

    # Admin header
    ws.merge_cells(start_row=2, start_column=admin_col_start,
                   end_row=2, end_column=admin_col_start + 3)
    c = ws.cell(row=2, column=admin_col_start, value="── SHARED ADMIN (split equally) ──")
    c.fill = FILL_SECTION; c.font = FONT_HDR; c.alignment = ALIGN_CENTER

    # ── Column sub-headers (row 3) ────────────────────────────────
    ws.cell(row=3, column=1, value="Month").fill   = FILL_LABEL
    ws.cell(row=3, column=1).font                  = FONT_BOLD
    ws.cell(row=3, column=1).alignment             = ALIGN_CENTER
    ws.column_dimensions["A"].width = 12

    team_sub_headers = [
        ("Diamonds",          FILL_ENTRY,    "#,##0"),
        ("Creators",          FILL_ENTRY,    "0"),
        ("Creator Sal. Act.", FILL_OVERRIDE, "$#,##0"),
        ("Rent Act.",         FILL_OVERRIDE, "$#,##0"),
        ("Utility Act.",      FILL_OVERRIDE, "$#,##0"),
        ("Buffer Act.",       FILL_OVERRIDE, "$#,##0"),
        ("Other Cost Act.",   FILL_OVERRIDE, "$#,##0"),
    ]
    for ti in range(n_teams):
        col_s = monthly_data_team_col_start(ti)
        for ci, (label, fill, _fmt) in enumerate(team_sub_headers):
            col = col_s + ci
            c = ws.cell(row=3, column=col, value=label)
            c.fill = fill; c.font = FONT_BOLD
            c.alignment = ALIGN_CENTER; c.border = BORDER_THIN
            ws.column_dimensions[get_column_letter(col)].width = 14

    admin_headers = ["Clothes", "Taxi", "Meals", "Other Admin"]
    for ci, label in enumerate(admin_headers):
        col = admin_col_start + ci
        c = ws.cell(row=3, column=col, value=label)
        c.fill = FILL_ENTRY; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(col)].width = 14

    # ── Data rows ─────────────────────────────────────────────────
    for ri, month in enumerate(MONTHS):
        row = 4 + ri
        c = ws.cell(row=row, column=1, value=month)
        c.fill = FILL_WHITE; c.font = FONT_NORM
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN

        for ti in range(n_teams):
            col_s = monthly_data_team_col_start(ti)
            for ci, (_label, fill, fmt) in enumerate(team_sub_headers):
                col = col_s + ci
                c = ws.cell(row=row, column=col, value=None)
                c.fill = fill; c.font = FONT_NORM
                c.number_format = fmt; c.border = BORDER_THIN

        for ci in range(4):
            col = admin_col_start + ci
            c = ws.cell(row=row, column=col, value=None)
            c.fill = FILL_ENTRY; c.font = FONT_NORM
            c.number_format = "$#,##0"; c.border = BORDER_THIN

    # Store metadata for UE_Summary to use
    ws._admin_col_start = admin_col_start
    ws._data_row_start  = 4
    ws._data_row_end    = 4 + len(MONTHS) - 1
```

- [ ] Run and verify Monthly_Data sheet has correct team blocks, colour coding, and admin columns.

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add Monthly_Data sheet with override columns"
  ```

---

## Chunk 3: UE_Summary + Team sheets

### Task 5: UE_Summary sheet

**Files:**
- Modify: `generate_colombia_tracker.py` — add `build_ue_summary(wb)`

- [ ] Add `build_ue_summary(wb)`:

```python
def build_ue_summary(wb):
    ws = wb.create_sheet("UE_Summary")
    team_ids = [t["id"] for t in TEAMS]

    headers = [
        "Month", "Team ID", "Team Name",
        "Diamonds", "Revenue (USD)",
        "Creator Salary", "Staff Cost",
        "Rent", "Utility", "Buffer", "Admin Share", "Other",
        "Total Cost", "Gross Profit", "Margin %",
        "Stage",
        "Cumul. Net", "Initial Invest.", "Recovered?", "Month #"
    ]
    header_fmts = [
        None, None, None,
        "#,##0", "$#,##0.00",
        "$#,##0.00", "$#,##0.00",
        "$#,##0.00", "$#,##0.00", "$#,##0.00", "$#,##0.00", "$#,##0.00",
        "$#,##0.00", "$#,##0.00", "0.0%",
        None,
        "$#,##0.00", "$#,##0.00", None, "0"
    ]

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws["A1"]
    c.value = "📊  UE SUMMARY — Auto-calculated, do not edit"
    c.fill = FILL_HEADER; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER; ws.row_dimensions[1].height = 26

    # Column headers
    for ci, hdr in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=hdr)
        c.fill = FILL_CALC; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(ci)].width = 14

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 14

    md_ws = wb["Monthly_Data"]

    row = 3
    for month_idx, month in enumerate(MONTHS):
        md_row = md_ws._data_row_start + month_idx

        for ti, team in enumerate(TEAMS):
            tid = team["id"]
            col_s = monthly_data_team_col_start(ti)   # diamonds col in Monthly_Data

            # Column letters in Monthly_Data
            dia_col  = get_column_letter(col_s)       # Diamonds
            crt_col  = get_column_letter(col_s + 1)   # Creator count
            csa_col  = get_column_letter(col_s + 2)   # Creator salary actual
            rnt_col  = get_column_letter(col_s + 3)   # Rent actual
            utl_col  = get_column_letter(col_s + 4)   # Utility actual
            buf_col  = get_column_letter(col_s + 5)   # Buffer actual
            oth_col  = get_column_letter(col_s + 6)   # Other actual

            # Admin columns (split equally)
            adm_start = md_ws._admin_col_start
            admin_sum_range = (
                f"Monthly_Data!{get_column_letter(adm_start)}{md_row}:"
                f"{get_column_letter(adm_start+3)}{md_row}"
            )

            # Revenue formula
            revenue_f = (
                f"=IF(Monthly_Data!{dia_col}{md_row}=\"\",0,"
                f"Monthly_Data!{dia_col}{md_row}*CONFIG_DIAMOND_RATE*CONFIG_TAKE_RATE)"
            )
            # Creator salary: max(base × count, revenue × revshare%)
            # actual override if provided
            creator_sal_f = (
                f"=IF(Monthly_Data!{csa_col}{md_row}<>\"\","
                f"Monthly_Data!{csa_col}{md_row},"
                f"MAX(CONFIG_CREATOR_BASE*IF(Monthly_Data!{crt_col}{md_row}=\"\",0,Monthly_Data!{crt_col}{md_row}),"
                f"{revenue_f[1:]}*CONFIG_CREATOR_REVSHARE))"   # strip leading =
            )
            # But creator_sal_f above is messy — simplify with intermediate
            # We'll write it directly:
            dia_ref = f"Monthly_Data!{dia_col}{md_row}"
            crt_ref = f"Monthly_Data!{crt_col}{md_row}"
            csa_ref = f"Monthly_Data!{csa_col}{md_row}"
            rev_expr = f"{dia_ref}*CONFIG_DIAMOND_RATE*CONFIG_TAKE_RATE"
            creator_sal_f = (
                f"=IF({csa_ref}<>\"\",{csa_ref},"
                f"MAX(CONFIG_CREATOR_BASE*IF({crt_ref}=\"\",1,{crt_ref}),"
                f"{rev_expr}*CONFIG_CREATOR_REVSHARE))"
            )

            # Staff cost: sum of this team's cost column in Staff sheet
            staff_ws = wb["Staff"]
            n_staff = len(STAFF)
            cost_col = get_column_letter(staff_ws._col_cost_start + ti)
            staff_f = f"=SUM(Staff!{cost_col}4:Staff!{cost_col}{3+n_staff})"

            # Individual cost overrides
            def override_f(actual_col, config_name):
                return (f"=IF(Monthly_Data!{actual_col}{md_row}<>\"\","
                        f"Monthly_Data!{actual_col}{md_row},{config_name})")

            rent_f    = override_f(rnt_col, "CONFIG_RENT")
            utility_f = override_f(utl_col, "CONFIG_UTILITY")
            buffer_f  = override_f(buf_col, "CONFIG_BUFFER")
            other_f   = f"=IF(Monthly_Data!{oth_col}{md_row}<>\"\",Monthly_Data!{oth_col}{md_row},0)"

            # Admin share = total admin / active team count (always equal split)
            n_active = sum(1 for t in TEAMS if t["status"] == "Active")
            admin_f = f"=SUM({admin_sum_range})/{n_active}"

            # Total cost (cols 6-12 = F-L in this sheet, row `row`)
            total_f = f"=SUM(F{row}:L{row})"
            profit_f = f"=E{row}-M{row}"
            margin_f = f"=IF(E{row}=0,0,N{row}/E{row})"

            # Stage
            stage_f = (
                f'=IFS(D{row}>=CONFIG_STAGE_OPT,"🚀 Optimistic",'
                f'D{row}>=CONFIG_STAGE_STB,"✅ Stable",'
                f'D{row}>=CONFIG_STAGE_BE,"⚠️ Breakeven",'
                f'TRUE,"🌱 Incubation")'
            )

            # Cumulative net: sum of all previous profit rows for same team
            # Find previous rows for this team
            prev_rows = [3 + (m * len(TEAMS)) + ti for m in range(month_idx)]
            if prev_rows:
                prev_refs = "+".join(f"N{r}" for r in prev_rows)
                cumul_f = f"=N{row}+{prev_refs}"
            else:
                cumul_f = f"=N{row}"

            invest_f   = "=CONFIG_INITIAL_INVEST"
            recovered_f = f'=IF(Q{row}>=CONFIG_INITIAL_INVEST,"✅ YES","⏳ No")'

            # Month number since launch
            launch_str = team["launch"]
            month_num_f = (
                f'=DATEDIF(DATE(LEFT("{launch_str}",4),MID("{launch_str}",6,2),1),'
                f'DATE(LEFT(A{row},4),RIGHT(A{row},2),1),"M")+1'
            )

            values = [
                (month,       None),
                (tid,         None),
                (team["name"],None),
                (f"=IF(Monthly_Data!{dia_ref[13:]}=\"\",0,Monthly_Data!{dia_ref[13:]})",
                              "#,##0"),
                (revenue_f,   "$#,##0.00"),
                (creator_sal_f,"$#,##0.00"),
                (staff_f,     "$#,##0.00"),
                (rent_f,      "$#,##0.00"),
                (utility_f,   "$#,##0.00"),
                (buffer_f,    "$#,##0.00"),
                (admin_f,     "$#,##0.00"),
                (other_f,     "$#,##0.00"),
                (total_f,     "$#,##0.00"),
                (profit_f,    "$#,##0.00"),
                (margin_f,    "0.0%"),
                (stage_f,     None),
                (cumul_f,     "$#,##0.00"),
                (invest_f,    "$#,##0.00"),
                (recovered_f, None),
                (month_num_f, "0"),
            ]

            for ci, (val, fmt) in enumerate(values, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = FILL_CALC if ci > 3 else FILL_WHITE
                c.font = FONT_NORM; c.border = BORDER_THIN
                c.alignment = ALIGN_CENTER
                if fmt: c.number_format = fmt

            row += 1

    # Freeze top rows
    ws.freeze_panes = "D3"
```

**Note on diamonds formula:** The `dia_ref` is like `Monthly_Data!B4`. The formula for column D uses `=IF(Monthly_Data!B4="",0,Monthly_Data!B4)` — this is straightforward but note the string construction skips "Monthly_Data!" prefix when building the inner reference. Double-check this during testing.

- [ ] Run script, open UE_Summary, manually verify one row's numbers match expected calculation.

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add UE_Summary sheet with live formulas"
  ```

---

### Task 6: Individual team sheets

**Files:**
- Modify: `generate_colombia_tracker.py` — add `build_team_sheet(wb, team)`

- [ ] Add `build_team_sheet(wb, team)`:

```python
def build_team_sheet(wb, team):
    tid = team["id"]
    ws = wb.create_sheet(f"Team_{tid}")
    ue = wb["UE_Summary"]

    # Find row indices in UE_Summary for this team
    team_rows = []
    for ri, month in enumerate(MONTHS):
        ti = next(i for i, t in enumerate(TEAMS) if t["id"] == tid)
        ue_row = 3 + ri * len(TEAMS) + ti
        team_rows.append(ue_row)

    # ── Title ─────────────────────────────────────────────────────
    ws.merge_cells("A1:T1")
    c = ws["A1"]
    c.value = f"🏠  {tid} — {team['name']} ({team['market']})"
    c.fill = FILL_HEADER; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER; ws.row_dimensions[1].height = 28

    # Meta info
    meta = [
        ("Launch Date:", team["launch"]),
        ("Status:", team["status"]),
        ("Market:", team["market"]),
    ]
    for ci, (label, val) in enumerate(meta):
        c = ws.cell(row=2, column=1 + ci*2, value=label)
        c.font = FONT_BOLD
        ws.cell(row=2, column=2 + ci*2, value=val).font = FONT_NORM

    # ── Monthly table header ───────────────────────────────────────
    headers = [
        "Month", "Diamonds", "Revenue", "Creator Sal.", "Staff Cost",
        "Rent", "Utility", "Buffer", "Admin", "Other",
        "Total Cost", "Gross Profit", "Margin %",
        "Stage", "Cumul. Net", "Recovered?", "Month #"
    ]
    # These map to UE_Summary columns D–T (4–20)
    ue_col_map = [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 20]

    for ci, hdr in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=hdr)
        c.fill = FILL_LABEL; c.font = FONT_BOLD
        c.alignment = ALIGN_CENTER; c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(ci)].width = 13

    # Data rows linking to UE_Summary
    for ri, ue_row in enumerate(team_rows):
        row = 5 + ri
        for ci, ue_col in enumerate(ue_col_map, 1):
            col_letter = get_column_letter(ue_col)
            c = ws.cell(row=row, column=ci,
                        value=f"=UE_Summary!{col_letter}{ue_row}")
            c.fill = FILL_CALC; c.font = FONT_NORM
            c.border = BORDER_THIN; c.alignment = ALIGN_CENTER
            # Copy number format
            ue_fmt_map = {
                4: "#,##0", 5: "$#,##0.00", 6: "$#,##0.00", 7: "$#,##0.00",
                8: "$#,##0.00", 9: "$#,##0.00", 10: "$#,##0.00", 11: "$#,##0.00",
                12: "$#,##0.00", 13: "$#,##0.00", 14: "$#,##0.00", 15: "0.0%",
                17: "$#,##0.00", 20: "0"
            }
            if ue_col in ue_fmt_map:
                c.number_format = ue_fmt_map[ue_col]

    # ── Investment Recovery Chart ──────────────────────────────────
    if len(team_rows) >= 1:
        chart = LineChart()
        chart.title = f"{tid} — Cumulative Net vs Break-even"
        chart.style = 10
        chart.y_axis.title = "USD"
        chart.x_axis.title = "Month"
        chart.height = 12; chart.width = 22

        # Cumulative net = col 15 in team sheet
        data_row_start = 5
        data_row_end   = 5 + len(team_rows) - 1
        data_ref = Reference(ws, min_col=15, min_row=data_row_start, max_row=data_row_end)
        chart.add_data(data_ref, titles_from_data=False)
        chart.series[0].title = SeriesLabel(v="Cumulative Net Profit")
        chart.series[0].graphicalProperties.line.solidFill = "2E75B6"

        chart_row = 5 + len(team_rows) + 2
        ws.add_chart(chart, f"A{chart_row}")
```

- [ ] Run script and open each Team_COL01 and Team_COL02 sheet. Verify data links back to UE_Summary and chart appears.

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add per-team sheets with investment recovery chart"
  ```

---

## Chunk 4: Polish + sample data + final output

### Task 7: Add sample data and freeze panes

**Files:**
- Modify: `generate_colombia_tracker.py` — pre-fill MONTHS with sample data

- [ ] Add sample monthly data to `build_monthly_data()` for testing:

Add a `SAMPLE_DATA` dict before `main()`:
```python
# Pre-fill some sample values so UE formulas are testable immediately
# Format: (month, team_index): {col_offset: value}
# col_offset: 0=Diamonds, 1=Creators, rest=overrides (blank = use Config)
SAMPLE_DATA = {
    ("2026-02", 0): {0: 890000, 1: 4},   # COL01 Feb
    ("2026-02", 1): {0: 450000, 1: 3},   # COL02 Feb
    ("2026-03", 0): {0: 1200000, 1: 4},  # COL01 Mar
    ("2026-03", 1): {0: 620000, 1: 3},   # COL02 Mar
}
```

In `build_monthly_data()`, after the blank row loop, add:
```python
    # Fill sample data
    for (month, ti), cols in SAMPLE_DATA.items():
        if month not in MONTHS:
            continue
        md_row = md_ws._data_row_start + MONTHS.index(month)
        col_s  = monthly_data_team_col_start(ti)
        for col_offset, val in cols.items():
            c = ws.cell(row=md_row, column=col_s + col_offset, value=val)
```

- [ ] Run, open UE_Summary, verify numbers are populated and plausible:
  - COL01 Feb: ~890,000 diamonds → revenue ≈ $5,785 → breakeven stage
  - COL02 Feb: ~450,000 diamonds → revenue ≈ $2,925 → incubation stage

- [ ] Add freeze panes to Monthly_Data and freeze top 2 rows + col A:
  ```python
  ws.freeze_panes = "B4"  # in build_monthly_data
  ```

- [ ] Commit:
  ```bash
  git add generate_colombia_tracker.py
  git commit -m "feat: add sample data and freeze panes"
  ```

---

### Task 8: Final QA checklist

- [ ] Open `Colombia_Teams.xlsx` and verify:
  - [ ] Config: All assumptions editable (yellow), named ranges exist
  - [ ] Staff: Changing a ✓ updates Teams Covered, Coeff, and per-team cost
  - [ ] Monthly_Data: Blank override cells leave Config defaults in UE_Summary
  - [ ] Monthly_Data: Entering a rent override for COL01 changes UE_Summary COL01 rent for that month
  - [ ] UE_Summary: Stage shows correct emoji label for each diamond count
  - [ ] UE_Summary: Cumulative net increases month over month for same team
  - [ ] Team_COL01: All values match UE_Summary for COL01 rows
  - [ ] Team_COL01: Chart renders with a line

- [ ] Commit final file:
  ```bash
  git add Colombia_Teams.xlsx generate_colombia_tracker.py
  git commit -m "feat: deliver Colombia_Teams.xlsx v1.0"
  ```

---

## Chunk 5: Documentation

### Task 9: Update memory + Obsidian

- [ ] Update `MEMORY.md` to reflect Excel tracker completion.

- [ ] Add Obsidian note:
  - Path: `03_Guild_Operations/Colombia/2026-03-31_colombia-team-tracker.md`
  - Content: overview, file location, how to add a new team, key formula notes

---

## Summary

| Sheet | Purpose | Who edits |
|-------|---------|-----------|
| Config | Assumptions + team registry | Jas (rarely) |
| Staff | Staff list + team assignment | Jas (when team changes) |
| Monthly_Data | Diamond counts + actual overrides | Jas (monthly) |
| UE_Summary | Calculated UE per team/month | No one (formulas) |
| Team_COL0X | Per-team view + chart | No one (formulas) |

**To add a new team:** Add a row to Config → add a column to Staff → re-run generator (or manually insert the team block in Monthly_Data and a new UE_Summary section).
