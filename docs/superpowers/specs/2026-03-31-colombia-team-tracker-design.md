# Colombia Guild Team Performance Tracker — Design Spec

**Date:** 2026-03-31
**Project:** Lionlive Colombia Guild Operations
**Deliverable:** `Colombia_Teams.xlsx`
**Author:** Jas (Jasmine / 蒋子妍)

---

## Background

Colombia guild currently has 2 live-streaming teams (团播), scaling to 6 next month. Need an Excel model to track:
- Per-team unit economics (diamonds → revenue → costs → profit)
- Staff cost allocation across teams (shared or dedicated, configurable)
- Admin shared expenses (clothes, taxi, meals) — split equally
- Investment recovery tracking per team
- Monthly actual-value overrides over assumption defaults

Daily performance tracking is a separate future project (TikTok backstage → auto-pull). This Excel covers monthly financials only.

---

## Key Assumptions (Config defaults)

| Parameter | Value |
|-----------|-------|
| Take Rate | 65% |
| Diamond-to-USD | 1 diamond = $0.01 (i.e. 1M diamonds = $10,000) — NOTE: verify exact rate |
| Monthly Rent | $600 |
| Monthly Utility | $800 |
| Monthly Buffer | $1,000 |
| Initial Investment per Team | $11,430 (Decoration $3,000 + Device $6,000 + Clothing $2,000 + HR first month ~$430) |
| Creator Base Salary | $500 per creator |
| Creator Revenue Share | 25% of team revenue |
| Creator Total Salary | max(Base × count, Revenue Share × revenue) |

Stage thresholds (diamonds/month per team):
- Incubation: < 500,000
- Breakeven: 500,000 – 1,060,000
- Stable: 1,060,000 – 2,000,000
- Optimistic: > 2,000,000

---

## Sheet Architecture

### Sheet 1: Config
Single source of truth for all assumptions and team registry.

**Section A — Team Registry** (rows 3–20):
| Team ID | Team Name | Market | Launch Date | Status |
|---------|-----------|--------|-------------|--------|
| COL01 | Team 1 | Colombia | 2025-XX-XX | Active |
| COL02 | Team 2 | Colombia | 2026-XX-XX | Active |

**Section B — UE Assumptions** (rows 22–35):
All values are named cells (Excel named ranges) for formula reference.
- Take Rate: `CONFIG_TAKE_RATE`
- Rent: `CONFIG_RENT`
- Utility: `CONFIG_UTILITY`
- Buffer: `CONFIG_BUFFER`
- Initial Investment: `CONFIG_INITIAL_INVEST`
- Creator Base Salary: `CONFIG_CREATOR_BASE`
- Revenue Share %: `CONFIG_CREATOR_REVSHARE`

**Section C — Stage Thresholds** (rows 37–45):
Named ranges: `CONFIG_STAGE_INC`, `CONFIG_STAGE_BE`, `CONFIG_STAGE_STB`

---

### Sheet 2: Staff

Defines all staff, their roles, salaries, and team assignments.

| Column | Content |
|--------|---------|
| A | Staff Name |
| B | Role |
| C | Monthly Salary (USD) |
| D | COL01 (✓ or blank) |
| E | COL02 (✓ or blank) |
| F... | [future teams, one column each] |
| [N-2] | Teams Covered (COUNTIF of ✓ in D:F range) |
| [N-1] | Allocation Coefficient (1/Teams_Covered, 0 if 0 teams) |
| [N] per team | COL01 Cost (Salary × Coeff if ✓ in COL01 column) |

**Design note:** When a new team is added, insert a column for it before "Teams Covered." Staff cost per team auto-recalculates.

Example staff rows:
- Ops Manager (shared across both teams) → coeff = 0.5
- Dance Teacher (COL01 only) → coeff = 1.0
- MUA (both teams) → coeff = 0.5

---

### Sheet 3: Monthly_Data

Data entry sheet. One row per month. Columns are grouped per team.

**Column layout:**
```
A: Month (YYYY-MM)

[COL01 block — cols B to K]
B: COL01 Diamonds (actual entry)
C: COL01 Creator Count (actual entry)
D: COL01 Creator Salary Override (actual, leave blank to use formula)
E: COL01 Rent Override
F: COL01 Utility Override
G: COL01 Buffer Override
H: COL01 Other Cost Override (misc)

[COL02 block — cols I to P, same structure]
...

[Shared Admin block — last cols]
Z: Clothes (USD)
AA: Taxi (USD)
AB: Meals (USD)
AC: Other Admin (USD)
```

**Color coding:**
- Yellow headers: data entry required
- Blue cells: actual-value override (leave blank = use Config assumption)
- Grey cells: auto-calculated (no entry needed)

**Formula logic for each cost line (example: Rent for COL01):**
```
= IF(E_row <> "", E_row, CONFIG_RENT)
```

**Admin per team = total admin / number of active teams** (equal split, configurable in Config).

---

### Sheet 4: UE_Summary

Auto-calculated. One row per team per month.

| Column | Content |
|--------|---------|
| A | Month |
| B | Team ID |
| C | Team Name |
| D | Diamonds |
| E | Revenue (USD) = Diamonds × 0.01 × Take Rate |
| F | Creator Salary |
| G | Staff Cost (from Staff sheet, sum of per-team costs) |
| H | Rent |
| I | Utility |
| J | Buffer |
| K | Admin Share |
| L | Other |
| M | Total Cost |
| N | Gross Profit |
| O | Gross Margin % |
| P | Stage (IFS formula vs thresholds) |
| Q | Cumulative Net (running sum of Gross Profit) |
| R | Initial Investment |
| S | Recovered? (YES if Cumulative Net ≥ Initial Investment) |
| T | Month Number (since launch) |

---

### Sheet 5+: Team_COL01, Team_COL02, ...

Individual team view. One sheet per team.

Sections:
1. **Header**: Team name, launch date, market, current stage
2. **Monthly table**: All UE_Summary columns for this team only, formatted nicely
3. **Investment recovery chart**: Cumulative net profit over time (line chart vs $0 and break-even line)
4. **Cost breakdown chart**: Stacked bar by cost category per month

---

## Two-Layer Data Model

```
Config (assumptions)           Monthly_Data (actuals)
       ↓                              ↓
  IF(actual <> "", actual, assumption)
       ↓
  UE_Summary (calculated)
       ↓
  Team sheets (visualized)
```

This allows:
- Projections with default assumptions before actuals are known
- Overriding individual line items as the month closes
- Historical accuracy as data accumulates

---

## Scope Boundaries

**In scope:**
- Colombia teams only (COL01, COL02, expandable)
- Monthly granularity
- Manual diamond data entry
- Shared + dedicated staff cost allocation
- Admin expense equal split
- Investment recovery tracking

**Out of scope:**
- Mexico teams (separate file, same template)
- Daily diamond auto-pull (separate TikTok backstage project)
- Cross-country summary
- PDF export

---

## File Generation

Delivered as Python script (`generate_colombia_tracker.py`) using `openpyxl`.
Output: `Colombia_Teams.xlsx` in project root.

Script is re-runnable (overwrites existing file). After initial generation, user edits live data directly in Excel.
