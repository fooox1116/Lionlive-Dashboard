# Lionlive Dashboard — Project Instructions

## What this repo is
Financial dashboard + Colombia guild team tracker for Toplive Media Co. (Jas, 50/50 with Jay).

Two sub-projects:
1. **`index.html`** — Lionlive Financial Dashboard (GitHub Pages, v8.0)
2. **`generate_colombia_tracker.py` + `Colombia_Teams.xlsx`** — Monthly UE tracker for Colombia 团播 teams

## Key constraints

### Never change the settlement direction
Jay contributed more → Jasmine pays Jay. Before stating any settlement conclusion, write the math explicitly.

### Dashboard deploys to GitHub Pages
No build step. Edit `index.html` directly. Push `main` branch → live at https://fooox1116.github.io/Lionlive-Dashboard/

### Excel tracker: no script editing for team management
Teams are managed via the Config sheet Status dropdown (Not Launched Yet → 准备中 → Active → Closed). Only edit the Python script when adding months or permanent new slots beyond 10.

### Re-running the generator overwrites Colombia_Teams.xlsx
Back up any live data entries before re-running `python3 generate_colombia_tracker.py`.

## Commit conventions
- `feat:` new feature
- `fix:` bug fix
- `v8.x` version tags in dashboard commits

## After significant work
Run the obsidian-context-sync skill to save a note and update `06_Claude_Contexts/Guild_Colombia_Context.md`.

## Compact Instructions (on context compression)
When compressing, preserve in priority order:
1. Architecture decisions (NEVER summarize)
2. Modified files and their key changes
3. Current verification status (pass/fail)
4. Open TODOs and rollback notes
5. Tool outputs (can delete, keep pass/fail only)
