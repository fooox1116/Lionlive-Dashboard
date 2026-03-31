#!/usr/bin/env python3
"""
Generate Colombia_Teams.xlsx — Lionlive guild team performance tracker.

Usage:
    python3 generate_colombia_tracker.py

Outputs Colombia_Teams.xlsx in the current directory. Re-run to regenerate.
After generation, edit Monthly_Data and Staff sheets directly in Excel;
all other sheets recalculate automatically via formulas.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.workbook.defined_name import DefinedName

OUTPUT_FILE = "Colombia_Teams.xlsx"

# ── Colour palette ─────────────────────────────────────────────────────────────
FILL_HEADER   = PatternFill("solid", fgColor="1F3864")   # dark navy
FILL_SECTION  = PatternFill("solid", fgColor="2E75B6")   # mid blue
FILL_LABEL    = PatternFill("solid", fgColor="D6E4F7")   # light blue
FILL_ENTRY    = PatternFill("solid", fgColor="FFF2CC")   # yellow — data entry
FILL_OVERRIDE = PatternFill("solid", fgColor="DDEEFF")   # blue — actual overrides
FILL_CALC     = PatternFill("solid", fgColor="F2F2F2")   # grey — auto-calc
FILL_WHITE    = PatternFill("solid", fgColor="FFFFFF")

FONT_HDR  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
FONT_BOLD = Font(name="Arial", bold=True, size=10)
FONT_NORM = Font(name="Arial", size=10)
FONT_TINY = Font(name="Arial", size=9, color="666666")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")
ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center")

THIN = Side(style="thin", color="CCCCCC")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ── Team definitions ───────────────────────────────────────────────────────────
# Add new teams here; sheets will be auto-generated.
TEAMS = [
    {"id": "COL01", "name": "Team 1", "market": "Colombia",
     "launch": "2025-09-01", "status": "Active"},
    {"id": "COL02", "name": "Team 2", "market": "Colombia",
     "launch": "2026-02-01", "status": "Active"},
]

# ── Staff definitions ──────────────────────────────────────────────────────────
# "teams": list of team IDs this person is assigned to.
# Allocation coefficient = 1 / len(teams). Per-team cost = salary × coeff.
STAFF = [
    {"name": "Ops Manager",   "role": "Operations",    "salary": 800,  "teams": ["COL01", "COL02"]},
    {"name": "Dance Teacher", "role": "Dance Teacher",  "salary": 600,  "teams": ["COL01"]},
    {"name": "MUA",           "role": "MUA",            "salary": 500,  "teams": ["COL01", "COL02"]},
    {"name": "Tech Support",  "role": "Tech",           "salary": 400,  "teams": ["COL01", "COL02"]},
]

# ── Months to include ──────────────────────────────────────────────────────────
# Add more rows here; each month appears in Monthly_Data and UE_Summary.
MONTHS = ["2026-02", "2026-03"]

# ── Sample data (for testing; leave real cells blank or override in Excel) ─────
# Key: (month, team_index_in_TEAMS), Value: {col_offset: value}
# col_offset 0 = Diamonds, 1 = Creator count; 2-6 = cost overrides (leave blank → Config default)
SAMPLE_DATA = {
    ("2026-02", 0): {0: 890_000, 1: 4},
    ("2026-02", 1): {0: 450_000, 1: 3},
    ("2026-03", 0): {0: 1_200_000, 1: 4},
    ("2026-03", 1): {0: 620_000,   1: 3},
}

# ── Layout constants ───────────────────────────────────────────────────────────
TEAM_BLOCK_COLS = 7    # columns per team in Monthly_Data
MONTH_COL       = 1    # column A = month label

def team_col_start(team_index: int) -> int:
    """1-based column index of the first column of a team block in Monthly_Data."""
    return MONTH_COL + 1 + team_index * TEAM_BLOCK_COLS


# ── Helpers ────────────────────────────────────────────────────────────────────

def cell(ws, row, col, value=None, fill=None, font=None, align=None,
         fmt=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill   is not None: c.fill            = fill
    if font   is not None: c.font            = font
    if align  is not None: c.alignment       = align
    if fmt    is not None: c.number_format   = fmt
    if border is not None: c.border          = border
    return c


def section_hdr(ws, row, col, text, span=2):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = FILL_SECTION; c.font = FONT_HDR; c.alignment = ALIGN_CENTER
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return c


def title_row(ws, row, text, total_cols, height=28):
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=total_cols)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = FILL_HEADER
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.alignment = ALIGN_CENTER
    ws.row_dimensions[row].height = height
    return c


def add_named_range(wb, name, sheet_name, cell_ref):
    """Add a workbook-level named range. Sheet name is always quoted."""
    wb.defined_names[name] = DefinedName(
        name, attr_text=f"'{sheet_name}'!{cell_ref}"
    )


# ── Sheet builders ─────────────────────────────────────────────────────────────

def build_config(wb):
    ws = wb.create_sheet("Config")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 12

    title_row(ws, 1, "⚙️  CONFIG — Single source of truth for all assumptions", 5)

    # ── Section A: Team Registry ───────────────────────────────────
    section_hdr(ws, 3, 1, "TEAM REGISTRY", span=5)
    for ci, hdr in enumerate(["Team ID", "Team Name", "Market", "Launch Date", "Status"], 1):
        cell(ws, 4, ci, hdr, fill=FILL_LABEL, font=FONT_BOLD,
             align=ALIGN_CENTER, border=BORDER_THIN)

    # Store launch-date cell refs so UE_Summary can reference them
    team_launch_cells = {}   # team_id → absolute cell ref like $D$5
    for ri, team in enumerate(TEAMS, 5):
        data = [team["id"], team["name"], team["market"], team["launch"], team["status"]]
        for ci, val in enumerate(data, 1):
            cell(ws, ri, ci, val, fill=FILL_WHITE, font=FONT_NORM,
                 align=ALIGN_CENTER, border=BORDER_THIN)
        team_launch_cells[team["id"]] = f"$D${ri}"

    # ── Section B: UE Assumptions ──────────────────────────────────
    row = 5 + len(TEAMS) + 2
    section_hdr(ws, row, 1,
                "UE ASSUMPTIONS  (edit yellow cells → all sheets update automatically)", span=3)
    row += 1

    assumptions = [
        ("Take Rate",              0.65,    "0%",      "CONFIG_TAKE_RATE"),
        ("Diamond → USD rate",     0.01,    "$0.0000", "CONFIG_DIAMOND_RATE"),
        ("Monthly Rent (USD)",     600,     "$#,##0",  "CONFIG_RENT"),
        ("Monthly Utility (USD)",  800,     "$#,##0",  "CONFIG_UTILITY"),
        ("Monthly Buffer (USD)",   1000,    "$#,##0",  "CONFIG_BUFFER"),
        ("Initial Investment/team",11430,   "$#,##0",  "CONFIG_INITIAL_INVEST"),
        ("Creator Base Salary",    500,     "$#,##0",  "CONFIG_CREATOR_BASE"),
        ("Revenue Share %",        0.25,    "0%",      "CONFIG_CREATOR_REVSHARE"),
    ]
    for label, val, fmt, name in assumptions:
        cell(ws, row, 1, label, fill=FILL_LABEL, font=FONT_BOLD,
             align=ALIGN_LEFT, border=BORDER_THIN)
        cell(ws, row, 2, val, fill=FILL_ENTRY, font=FONT_NORM,
             align=ALIGN_CENTER, fmt=fmt, border=BORDER_THIN)
        cell(ws, row, 3, f"← Named range: {name}", font=FONT_TINY)
        add_named_range(wb, name, "Config", f"$B${row}")
        row += 1

    # ── Section C: Stage Thresholds ───────────────────────────────
    row += 1
    section_hdr(ws, row, 1, "STAGE THRESHOLDS (diamonds / month)", span=3)
    row += 1
    for ci, hdr in enumerate(["Stage", "Min Diamonds (≥)"], 1):
        cell(ws, row, ci, hdr, fill=FILL_LABEL, font=FONT_BOLD,
             align=ALIGN_CENTER, border=BORDER_THIN)
    row += 1

    stages = [
        ("🌱 Incubation",  0),
        ("⚠️ Breakeven",   500_000,  "CONFIG_STAGE_BE"),
        ("✅ Stable",      1_060_000,"CONFIG_STAGE_STB"),
        ("🚀 Optimistic",  2_000_000,"CONFIG_STAGE_OPT"),
    ]
    stage_rows = {}  # stage_name → row  (for IFS formula reference)
    for stage_data in stages:
        stage_label = stage_data[0]
        stage_val   = stage_data[1]
        stage_name  = stage_data[2] if len(stage_data) > 2 else None
        cell(ws, row, 1, stage_label, font=FONT_NORM, border=BORDER_THIN)
        cell(ws, row, 2, stage_val, fill=FILL_ENTRY, font=FONT_NORM,
             align=ALIGN_CENTER, fmt="#,##0", border=BORDER_THIN)
        if stage_name:
            add_named_range(wb, stage_name, "Config", f"$B${row}")
        stage_rows[stage_label] = row
        row += 1

    # ── Legend ─────────────────────────────────────────────────────
    row += 1
    section_hdr(ws, row, 1, "LEGEND", span=2)
    row += 1
    for fill, label in [
        (FILL_ENTRY,    "Yellow = Data entry required"),
        (FILL_OVERRIDE, "Blue   = Actual override (blank → use Config assumption)"),
        (FILL_CALC,     "Grey   = Auto-calculated — do not edit"),
    ]:
        ws.cell(row=row, column=1).fill = fill
        cell(ws, row, 2, label, font=FONT_TINY)
        row += 1

    # Store for cross-sheet reference
    ws._team_launch_cells = team_launch_cells


def build_staff(wb):
    ws = wb.create_sheet("Staff")
    team_ids = [t["id"] for t in TEAMS]
    n_teams  = len(team_ids)

    col_first_chk = 4              # D: first team checkbox column
    col_covered   = col_first_chk + n_teams
    col_coeff     = col_covered + 1
    col_cost_s    = col_coeff + 1  # first per-team cost column
    total_cols    = col_cost_s + n_teams - 1

    # Column widths
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    for i in range(n_teams):
        ws.column_dimensions[get_column_letter(col_first_chk + i)].width = 10
    ws.column_dimensions[get_column_letter(col_covered)].width = 14
    ws.column_dimensions[get_column_letter(col_coeff)].width   = 14
    for i in range(n_teams):
        ws.column_dimensions[get_column_letter(col_cost_s + i)].width = 14

    title_row(ws, 1,
              '👥  STAFF — Mark "✓" in team columns; cost allocation auto-calculates',
              total_cols)

    # Instructions
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    cell(ws, 2, 1,
         "Allocation Coeff = 1 ÷ Teams Covered.  Per-team cost = Salary × Coeff (if assigned).",
         font=Font(name="Arial", italic=True, size=9, color="444444"),
         align=ALIGN_LEFT)

    # Header row 3
    for ci, hdr in enumerate(["Staff Name", "Role", "Salary (USD)"], 1):
        cell(ws, 3, ci, hdr, fill=FILL_LABEL, font=FONT_BOLD,
             align=ALIGN_CENTER, border=BORDER_THIN)
    for i, tid in enumerate(team_ids):
        cell(ws, 3, col_first_chk + i, tid,
             fill=FILL_LABEL, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
    cell(ws, 3, col_covered, "Teams Covered",
         fill=FILL_CALC, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
    cell(ws, 3, col_coeff,   "Alloc Coeff",
         fill=FILL_CALC, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
    for i, tid in enumerate(team_ids):
        cell(ws, 3, col_cost_s + i, f"{tid} Cost (USD)",
             fill=FILL_CALC, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)

    # Data rows
    for r_off, staff in enumerate(STAFF):
        row = 4 + r_off
        cell(ws, row, 1, staff["name"], font=FONT_NORM, border=BORDER_THIN)
        cell(ws, row, 2, staff["role"],  font=FONT_NORM, border=BORDER_THIN)
        cell(ws, row, 3, staff["salary"],
             fill=FILL_ENTRY, font=FONT_NORM, align=ALIGN_RIGHT,
             fmt="$#,##0", border=BORDER_THIN)

        chk_letters = []
        for i, tid in enumerate(team_ids):
            col = col_first_chk + i
            val = "✓" if tid in staff["teams"] else ""
            cell(ws, row, col, val,
                 fill=FILL_ENTRY, font=FONT_NORM, align=ALIGN_CENTER, border=BORDER_THIN)
            chk_letters.append(get_column_letter(col))

        # Teams Covered = COUNTIF across checkmark columns
        first_chk = chk_letters[0];  last_chk = chk_letters[-1]
        cov_letter = get_column_letter(col_covered)
        cell(ws, row, col_covered,
             f'=COUNTIF({first_chk}{row}:{last_chk}{row},"✓")',
             fill=FILL_CALC, font=FONT_NORM, align=ALIGN_CENTER,
             fmt="0", border=BORDER_THIN)

        # Coeff = 1/covered (0 if uncovered)
        cell(ws, row, col_coeff,
             f"=IF({cov_letter}{row}=0,0,1/{cov_letter}{row})",
             fill=FILL_CALC, font=FONT_NORM, align=ALIGN_CENTER,
             fmt="0.00", border=BORDER_THIN)

        # Per-team cost
        coeff_l = get_column_letter(col_coeff)
        sal_l   = get_column_letter(3)
        for i, tid in enumerate(team_ids):
            chk_l = chk_letters[i]
            formula = f'=IF({chk_l}{row}="✓",{sal_l}{row}*{coeff_l}{row},0)'
            cell(ws, row, col_cost_s + i, formula,
                 fill=FILL_CALC, font=FONT_NORM, align=ALIGN_RIGHT,
                 fmt="$#,##0.00", border=BORDER_THIN)

    # Totals row
    tot_row = 4 + len(STAFF)
    cell(ws, tot_row, 1, "TOTAL", font=FONT_BOLD, fill=FILL_LABEL, border=BORDER_THIN)
    cell(ws, tot_row, 3,
         f"=SUM(C4:C{tot_row-1})",
         fill=FILL_LABEL, font=FONT_BOLD, fmt="$#,##0", border=BORDER_THIN)
    for i in range(n_teams):
        cl = get_column_letter(col_cost_s + i)
        cell(ws, tot_row, col_cost_s + i,
             f"=SUM({cl}4:{cl}{tot_row-1})",
             fill=FILL_LABEL, font=FONT_BOLD, fmt="$#,##0.00", border=BORDER_THIN)

    # Store metadata for UE_Summary
    ws._col_cost_start  = col_cost_s
    ws._staff_data_rows = (4, tot_row - 1)


def build_monthly_data(wb):
    ws = wb.create_sheet("Monthly_Data")
    team_ids = [t["id"] for t in TEAMS]
    n_teams  = len(team_ids)

    admin_col_s = team_col_start(n_teams)   # first admin column (after all team blocks)
    last_col    = admin_col_s + 3

    ws.column_dimensions["A"].width = 12

    title_row(ws, 1, "📋  MONTHLY DATA — Enter actual values (blue = override; blank = use Config default)",
              last_col)

    # Team block headers (row 2)
    for ti, tid in enumerate(team_ids):
        col_s = team_col_start(ti)
        ws.merge_cells(start_row=2, start_column=col_s,
                       end_row=2, end_column=col_s + TEAM_BLOCK_COLS - 1)
        cell(ws, 2, col_s, f"── {tid} ──",
             fill=FILL_SECTION, font=FONT_HDR, align=ALIGN_CENTER)

    # Admin block header (row 2)
    ws.merge_cells(start_row=2, start_column=admin_col_s,
                   end_row=2, end_column=admin_col_s + 3)
    cell(ws, 2, admin_col_s, "── SHARED ADMIN (split equally across teams) ──",
         fill=FILL_SECTION, font=FONT_HDR, align=ALIGN_CENTER)

    # Column sub-headers (row 3)
    cell(ws, 3, 1, "Month",
         fill=FILL_LABEL, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)

    # (label, fill, number_format)
    team_sub = [
        ("Diamonds",           FILL_ENTRY,    "#,##0"),
        ("Creators",           FILL_ENTRY,    "0"),
        ("Creator Sal. Actual",FILL_OVERRIDE, "$#,##0"),
        ("Rent Actual",        FILL_OVERRIDE, "$#,##0"),
        ("Utility Actual",     FILL_OVERRIDE, "$#,##0"),
        ("Buffer Actual",      FILL_OVERRIDE, "$#,##0"),
        ("Other Cost Actual",  FILL_OVERRIDE, "$#,##0"),
    ]
    for ti in range(n_teams):
        col_s = team_col_start(ti)
        for ci, (label, fill, _fmt) in enumerate(team_sub):
            col = col_s + ci
            cell(ws, 3, col, label,
                 fill=fill, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
            ws.column_dimensions[get_column_letter(col)].width = 15

    admin_hdrs = ["Clothes", "Taxi", "Meals", "Other Admin"]
    for ci, label in enumerate(admin_hdrs):
        col = admin_col_s + ci
        cell(ws, 3, col, label,
             fill=FILL_ENTRY, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
        ws.column_dimensions[get_column_letter(col)].width = 14

    # Data rows
    data_row_start = 4
    for ri, month in enumerate(MONTHS):
        row = data_row_start + ri
        cell(ws, row, 1, month,
             fill=FILL_WHITE, font=FONT_NORM, align=ALIGN_CENTER, border=BORDER_THIN)

        for ti in range(n_teams):
            col_s = team_col_start(ti)
            for ci, (_label, fill, fmt) in enumerate(team_sub):
                col = col_s + ci
                # Pre-fill sample data if available
                val = None
                key = (month, ti)
                if key in SAMPLE_DATA and ci in SAMPLE_DATA[key]:
                    val = SAMPLE_DATA[key][ci]
                cell(ws, row, col, val,
                     fill=fill, font=FONT_NORM, fmt=fmt, border=BORDER_THIN)

        for ci in range(4):
            col = admin_col_s + ci
            cell(ws, row, col, None,
                 fill=FILL_ENTRY, font=FONT_NORM, fmt="$#,##0", border=BORDER_THIN)

    ws.freeze_panes = "B4"

    # Store metadata
    ws._admin_col_start = admin_col_s
    ws._data_row_start  = data_row_start
    ws._data_row_end    = data_row_start + len(MONTHS) - 1


def build_ue_summary(wb):
    ws    = wb.create_sheet("UE_Summary")
    md_ws = wb["Monthly_Data"]
    st_ws = wb["Staff"]
    cfg_ws = wb["Config"]

    team_ids   = [t["id"]   for t in TEAMS]
    team_names = [t["name"] for t in TEAMS]
    n_teams    = len(TEAMS)
    n_active   = sum(1 for t in TEAMS if t["status"] == "Active")

    headers = [
        "Month", "Team ID", "Team Name",
        "Diamonds", "Revenue (USD)",
        "Creator Salary", "Staff Cost",
        "Rent", "Utility", "Buffer", "Admin Share", "Other",
        "Total Cost", "Gross Profit", "Margin %",
        "Stage",
        "Cumul. Net", "Initial Invest.", "Recovered?", "Month #"
    ]
    fmts = [
        None, None, None,
        "#,##0", "$#,##0.00",
        "$#,##0.00", "$#,##0.00",
        "$#,##0.00", "$#,##0.00", "$#,##0.00", "$#,##0.00", "$#,##0.00",
        "$#,##0.00", "$#,##0.00", "0.0%",
        None,
        "$#,##0.00", "$#,##0.00", None, "0"
    ]

    title_row(ws, 1, "📊  UE SUMMARY — Auto-calculated. Do not edit.", len(headers))

    for ci, (hdr, fmt) in enumerate(zip(headers, fmts), 1):
        c = cell(ws, 2, ci, hdr,
                 fill=FILL_CALC, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
        ws.column_dimensions[get_column_letter(ci)].width = 14

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 14
    ws.freeze_panes = "D3"

    # Track written data rows per team for cumulative calculation
    team_profit_rows = {tid: [] for tid in team_ids}

    data_row = 3
    for month_idx, month in enumerate(MONTHS):
        md_row = md_ws._data_row_start + month_idx

        for ti, team in enumerate(TEAMS):
            tid  = team["id"]
            name = team["name"]
            cs   = team_col_start(ti)   # column start in Monthly_Data

            # Column letters in Monthly_Data
            dia_l = get_column_letter(cs)       # Diamonds
            crt_l = get_column_letter(cs + 1)   # Creator count
            csa_l = get_column_letter(cs + 2)   # Creator salary actual
            rnt_l = get_column_letter(cs + 3)   # Rent actual
            utl_l = get_column_letter(cs + 4)   # Utility actual
            buf_l = get_column_letter(cs + 5)   # Buffer actual
            oth_l = get_column_letter(cs + 6)   # Other cost actual

            dia_ref = f"Monthly_Data!{dia_l}{md_row}"
            crt_ref = f"Monthly_Data!{crt_l}{md_row}"
            csa_ref = f"Monthly_Data!{csa_l}{md_row}"

            admin_start_l = get_column_letter(md_ws._admin_col_start)
            admin_end_l   = get_column_letter(md_ws._admin_col_start + 3)
            admin_range   = f"Monthly_Data!{admin_start_l}{md_row}:{admin_end_l}{md_row}"

            # ── Formulas ─────────────────────────────────────────
            diamonds_f = f"=IF({dia_ref}=\"\",0,{dia_ref})"

            rev_expr   = f"{dia_ref}*CONFIG_DIAMOND_RATE*CONFIG_TAKE_RATE"
            revenue_f  = f"=IF({dia_ref}=\"\",0,{rev_expr})"

            # Creator salary: actual override OR max(base×count, rev×share%)
            # Blank creator count → 0 creators (not 1)
            creator_f  = (
                f"=IF({csa_ref}<>\"\",{csa_ref},"
                f"MAX(CONFIG_CREATOR_BASE*IF({crt_ref}=\"\",0,{crt_ref}),"
                f"{rev_expr}*CONFIG_CREATOR_REVSHARE))"
            )

            # Staff cost from Staff sheet (sum this team's cost column)
            sc_col  = get_column_letter(st_ws._col_cost_start + ti)
            sr_s, sr_e = st_ws._staff_data_rows
            staff_f = f"=SUM(Staff!{sc_col}{sr_s}:Staff!{sc_col}{sr_e})"

            def override_f(actual_l, config_name):
                return (f"=IF(Monthly_Data!{actual_l}{md_row}<>\"\","
                        f"Monthly_Data!{actual_l}{md_row},{config_name})")

            rent_f    = override_f(rnt_l, "CONFIG_RENT")
            utility_f = override_f(utl_l, "CONFIG_UTILITY")
            buffer_f  = override_f(buf_l, "CONFIG_BUFFER")
            other_f   = f"=IF(Monthly_Data!{oth_l}{md_row}<>\"\",Monthly_Data!{oth_l}{md_row},0)"
            admin_f   = f"=SUM({admin_range})/{n_active}"

            # Total cost = sum of cols F:L (creator … other) in UE_Summary
            # At data_row: creator=col6, staff=7, rent=8, util=9, buf=10, admin=11, other=12
            total_f  = f"=SUM(F{data_row}:L{data_row})"
            profit_f = f"=E{data_row}-M{data_row}"
            margin_f = f"=IF(E{data_row}=0,0,N{data_row}/E{data_row})"

            stage_f = (
                f'=IFS(D{data_row}>=CONFIG_STAGE_OPT,"🚀 Optimistic",'
                f'D{data_row}>=CONFIG_STAGE_STB,"✅ Stable",'
                f'D{data_row}>=CONFIG_STAGE_BE,"⚠️ Breakeven",'
                f'TRUE,"🌱 Incubation")'
            )

            # Cumulative net = this profit + all prior profits for same team
            prior_rows = team_profit_rows[tid]
            if prior_rows:
                prior_sum = "+".join(f"N{r}" for r in prior_rows)
                cumul_f = f"=N{data_row}+{prior_sum}"
            else:
                cumul_f = f"=N{data_row}"

            invest_f    = "=CONFIG_INITIAL_INVEST"
            recovered_f = f'=IF(Q{data_row}>=CONFIG_INITIAL_INVEST,"✅ YES","⏳ No")'

            # Month number: reference launch date from Config, not hardcoded string
            launch_cell_ref = cfg_ws._team_launch_cells[tid]  # e.g. $D$5
            month_num_f = (
                f'=DATEDIF(DATE(LEFT(\'Config\'!{launch_cell_ref},4),'
                f'MID(\'Config\'!{launch_cell_ref},6,2),1),'
                f'DATE(LEFT(A{data_row},4),RIGHT(A{data_row},2),1),"M")+1'
            )

            row_values = [
                (month,       None),
                (tid,         None),
                (name,        None),
                (diamonds_f,  "#,##0"),
                (revenue_f,   "$#,##0.00"),
                (creator_f,   "$#,##0.00"),
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

            for ci, (val, fmt) in enumerate(row_values, 1):
                c = ws.cell(row=data_row, column=ci, value=val)
                c.fill      = FILL_CALC if ci > 3 else FILL_WHITE
                c.font      = FONT_NORM
                c.border    = BORDER_THIN
                c.alignment = ALIGN_CENTER
                if fmt:
                    c.number_format = fmt

            team_profit_rows[tid].append(data_row)
            data_row += 1

    ws._data_row_start = 3
    ws._data_row_end   = data_row - 1
    ws._team_profit_rows = team_profit_rows


def build_team_sheet(wb, team):
    tid  = team["id"]
    name = team["name"]
    ws   = wb.create_sheet(f"Team_{tid}")
    ue   = wb["UE_Summary"]

    # Find this team's rows in UE_Summary
    team_ue_rows = ue._team_profit_rows[tid]

    # ── Title ──────────────────────────────────────────────────────
    title_row(ws, 1, f"🏠  {tid} — {name}  ({team['market']})", 17)

    # Meta
    for ci, (label, val) in enumerate([
        ("Launch Date:", team["launch"]),
        ("Status:", team["status"]),
        ("Market:", team["market"]),
    ]):
        cell(ws, 2, 1 + ci * 2,     label, font=FONT_BOLD)
        cell(ws, 2, 2 + ci * 2,     val,   font=FONT_NORM)

    # ── Monthly table ──────────────────────────────────────────────
    # Columns: Month, Diamonds, Revenue, Creator Sal, Staff Cost,
    #          Rent, Utility, Buffer, Admin, Other,
    #          Total Cost, Gross Profit, Margin %,
    #          Stage, Cumul. Net, Initial Invest., Recovered?, Month #
    col_headers = [
        "Month", "Diamonds", "Revenue", "Creator Sal.", "Staff Cost",
        "Rent", "Utility", "Buffer", "Admin", "Other",
        "Total Cost", "Gross Profit", "Margin %",
        "Stage", "Cumul. Net", "Initial Invest.", "Recovered?", "Month #"
    ]
    # Maps to UE_Summary columns (1-based)
    ue_col_map = [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
    col_fmts = {
        2: "#,##0", 3: "$#,##0.00", 4: "$#,##0.00", 5: "$#,##0.00",
        6: "$#,##0.00", 7: "$#,##0.00", 8: "$#,##0.00", 9: "$#,##0.00",
        10: "$#,##0.00", 11: "$#,##0.00", 12: "$#,##0.00", 13: "0.0%",
        15: "$#,##0.00", 16: "$#,##0.00", 18: "0"
    }

    hdr_row = 4
    for ci, hdr in enumerate(col_headers, 1):
        cell(ws, hdr_row, ci, hdr,
             fill=FILL_LABEL, font=FONT_BOLD, align=ALIGN_CENTER, border=BORDER_THIN)
        ws.column_dimensions[get_column_letter(ci)].width = 13

    for ri, ue_row in enumerate(team_ue_rows):
        row = hdr_row + 1 + ri
        for ci, ue_col in enumerate(ue_col_map, 1):
            ue_col_l = get_column_letter(ue_col)
            c = ws.cell(row=row, column=ci,
                        value=f"=UE_Summary!{ue_col_l}{ue_row}")
            c.fill = FILL_CALC; c.font = FONT_NORM
            c.border = BORDER_THIN; c.alignment = ALIGN_CENTER
            if ci in col_fmts:
                c.number_format = col_fmts[ci]

    # ── Investment Recovery Line Chart ────────────────────────────
    if team_ue_rows:
        data_start = hdr_row + 1
        data_end   = hdr_row + len(team_ue_rows)
        chart_row  = data_end + 3

        chart = LineChart()
        chart.title  = f"{tid} — Cumulative Net Profit vs Break-even"
        chart.style  = 10
        chart.height = 12
        chart.width  = 22
        chart.y_axis.title = "USD"
        chart.x_axis.title = "Month"

        # Cumulative Net is column 15 in the team sheet
        data_ref = Reference(ws, min_col=15, min_row=data_start, max_row=data_end)
        chart.add_data(data_ref, titles_from_data=False)
        chart.series[0].title = SeriesLabel(v="Cumulative Net Profit")
        chart.series[0].graphicalProperties.line.solidFill = "2E75B6"
        chart.series[0].graphicalProperties.line.width = 20000  # 2pt

        # Month labels from column 1
        cats = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
        chart.set_categories(cats)

        ws.add_chart(chart, f"A{chart_row}")


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    wb = Workbook()
    wb.remove(wb.active)   # remove default empty sheet

    build_config(wb)
    build_staff(wb)
    build_monthly_data(wb)
    build_ue_summary(wb)
    for team in TEAMS:
        build_team_sheet(wb, team)

    wb.save(OUTPUT_FILE)
    print(f"✅  Saved {OUTPUT_FILE}")
    print(f"    Sheets: {', '.join(ws.title for ws in wb.worksheets)}")
    print(f"    Teams: {len(TEAMS)} | Months: {len(MONTHS)} | Staff: {len(STAFF)}")


if __name__ == "__main__":
    main()
