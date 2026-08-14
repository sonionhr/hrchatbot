# Prompt — Org Chart Update Workflow
*Run this whenever `raw_docs/Orgchart.csv` (the Masterlist export) is updated, to regenerate `html/orgchart.html`.*

> **2026-08-13 — source changed from `Orgchart.xlsx` to `Orgchart.csv`.** The Power BI export moved from an Excel workbook (sheet `Export`, columns named like `Masterlist[Employee Key]`) to a plain CSV (columns named like `MasterlistEmployee Key` — no brackets, no sheet). `scripts/generate_orgchart.py` was updated to read the CSV; `raw_docs/Orgchart.xlsx` is no longer read by the script (it may still exist in the repo as a leftover export, but it is not the source of truth — don't update it expecting the chart to change).

---

## THE FLOW

1. Update and **save** `raw_docs/Orgchart.csv` (close it if it's open in Excel/another program — see STEP 1).
2. Run:
   ```
   python scripts/generate_orgchart.py
   ```
3. Read the console output (STEP 2) and spot-check the result (STEP 3).
4. Commit both `raw_docs/Orgchart.csv` and `html/orgchart.html` to git (STEP 4) — **this is not optional**, see why below.

That's it for the routine case. The rest of this file is what to do when something looks off, and the business rules baked into the script so nobody has to re-derive them from scratch.

---

## STEP 1 — Before running: make sure the file is actually settled

`Orgchart.csv` is a live Power BI export. Twice during development (back when this was still an .xlsx export, but the same risk applies to any live export), reading it mid-save produced a stable-looking but **wrong** snapshot — fewer rows, a missing root, or a different column layout — with no error to signal it.

The script now protects against this itself: `load_csv_rows_stable()` copies the file repeatedly a couple of seconds apart and only proceeds once two consecutive copies are byte-identical (up to 5 attempts, ~10s worst case). If it never settles, it prints a warning and proceeds anyway with the latest snapshot — don't ignore that warning; re-export/re-save the file and re-run.

You do not need to close whatever program has the file open for the script to read it (it always works from a copy), but you DO need to have saved your changes — the script reads whatever was last written to disk, not unsaved in-memory edits.

## STEP 2 — Read the console output

The script prints, in order:
1. Total raw row count.
2. Any duplicate Employee Keys found (should be none — investigate if not).
3. The consolidation log for both tiers (see BUSINESS RULES below) — one line per department head, with the resulting bucket name and headcount. **Skim this** — it's the fastest way to notice if a head is missing, a department jumped size unexpectedly, or a name collision got auto-disambiguated in a way you didn't expect.
4. Renderable node count vs excluded-DL count.
5. The company root's name/title.
6. Final department list with true headcounts.

If the company root isn't "Christian Nielsen (CN) — CEO & President", or a `WARNING: expected exactly 1 blank-parent root` line appears, the export's top-of-hierarchy row is missing or has a non-blank parent — check STEP 1 first (most likely cause), then check the raw file itself.

## STEP 3 — Spot-check the output

Open `html/orgchart.html` in a browser and check:
- **Company Org Chart** tab: root is Christian Nielsen, 5 direct reports shown by default.
- **Department Org Chart** tab: dropdown count roughly matches expectations (compare against the previous run's numbers in `CHANGELOG.md` or your own memory of the org — a department halving or doubling in size is worth a second look, not necessarily wrong, but worth explaining).
- **Full Org Chart** tab: loads without errors, sections fit the page width.
- Browser console: no errors (F12 → Console).

If you have Claude Code available, the fastest way to do this spot-check is to ask it to open the file in headless Chrome and screenshot each tab — that's how this was verified throughout development (see git/session history for the exact `chromium-cli`/Selenium pattern used).

## STEP 4 — Commit to git

`html/orgchart.html` and `raw_docs/Orgchart.csv` are **not currently gitignored**, but historically neither has actually been committed — which meant a bad regeneration had no way to be reverted (this happened once during development: a mid-save read silently overwrote a verified-good `orgchart.html` with incomplete data, and there was no git history to fall back on). Always commit after a verified-good regeneration:

```
git add raw_docs/Orgchart.csv html/orgchart.html
git commit -m "Update org chart data — <one-line summary of what changed in the source>"
```

If a regeneration ever produces something wrong and you haven't committed yet, `git diff`/`git checkout` won't help you — the previous version is simply gone. This is the one step in this workflow that exists specifically to prevent repeating that mistake.

---

## BUSINESS RULES BAKED INTO THE SCRIPT

These are not obvious from the raw spreadsheet and were worked out iteratively — read this before changing `scripts/generate_orgchart.py`'s logic, so a future edit doesn't silently undo one of them.

### Department grounding is 100% structural, never by text label
Every person's department bucket is derived by walking the real `Employee Key` → `Parentkey_revise` chain from the company root down — **never** by matching the raw `Masterlist[Groupdept_shown_VN]` text. Two people can share identical department text while being on completely unrelated management chains (e.g. a global finance person's row happens to say "Finance (904)" even though they report through a totally different chain than the actual Vietnam Finance team) — grounding on text alone would silently merge unrelated real teams. See the `consolidate_heads()` / `bfs_subtree()` functions.

Consolidation runs in two tiers:
- **Tier 1** — every real direct report of the company root (Christian Nielsen) becomes a department head; their entire real subtree gets one bucket name.
- **Tier 2** — Jason King is treated specially: instead of his whole VN organization becoming one bucket, *his* direct reports each become their own department head (overwriting Tier 1's broad stamp for just that subtree). This is what produces "Quality & CQS", "Production (300)", etc. as separate departments instead of one giant "Jason King" bucket.

If a department bucket's name collides with a sibling's (two heads defaulting to the same name/text), the script auto-disambiguates by appending the head's own name — e.g. two people both tagged "Management & Fin Corp" become "Management & Fin Corp — Person A" and "— Person B".

### Manual name overrides
`CONSOLIDATED_NAME_OVERRIDES` currently has two entries:
- **Jason King** → "VN Operations (Country Head)" — his raw dept text ("HR & Admin") was wrong; he heads all of Vietnam, not HR.
- **Nguyễn Huề** → "Digital Transformation & VNII Factory" — his real team spans several raw-text fragments (VNII Factory, VNII_Quality, CIM&MES, VNII) that are all one real unit under him; this consolidates them under one explicit name rather than his own ambiguous raw text.

If the org changes and a similar text/reality mismatch appears for someone else, add them here rather than special-casing elsewhere.

### DL (Direct Labor) employees are counted everywhere but never rendered
Every headcount and report-count badge counts **every** employee, including rank-and-file DL production operators. But DL people never get their own card and are never expandable — the render set (`nodes`) explicitly excludes `jobtype == "DL"`. Where every one of a person's direct reports is DL, a dashed "(not shown)" badge appears instead of an expand toggle, still showing the true count.

Report-count badges break down by job type (IDL / DLS / DL), e.g. `+ 3 IDL · 45 DL`, not one aggregate number. The count is also always the person's **total subordinate count at every level down** (computed once via `subtree_breakdown()`, a memoized post-order walk over the whole tree), not just their direct reports — and it's deliberately independent of the UI's collapse/expand state, which only controls what's rendered. A card collapsed to level 1 still shows that person's true company-wide total on its badge.

### Entity flag
Anyone whose `Entity` column value isn't `"VNI"` (the Vietnam baseline) gets a small colored badge on their card showing their actual entity (e.g. `SONIONAS ENTITY`, `VNII ENTITY`). This exists because a handful of global/corporate people are grouped into VN-sounding departments by the consolidation logic above, and the entity badge is how a reader tells "this person is structurally in this department, but isn't actually a VN employee."

### Unlinked employees
Anyone whose real management chain never reaches the company root within this export (a manager key missing from the data) gets a dept label suffixed `— unlinked (manager not in this export)`, rather than silently falling back to raw text that could coincidentally collide with — and get merged into — a real department.

### Column layout is read by name, not position
`raw_docs/Orgchart.csv`'s column order has changed between exports before (Power BI re-exports don't guarantee stable column order — this is also what changed the column *naming style* itself when the export moved from .xlsx to .csv, see the note at the top of this file). The script reads every field via `row.get(header_name)` on the parsed CSV dict rows, resolved from the actual header row — never a hardcoded column index. If a column is renamed or removed, the script will fail fast with a clear error (`REQUIRED` list check) rather than silently reading the wrong field.

### Known data gap: ~740 employees with no `Parentkey_revise` (as of the 2026-08-13 CSV export)
Unlike the old .xlsx export, this CSV export contains roughly 740 rows (mostly `Entity=SonionAS`, a mix of DL/DLS/IDL, mostly blank `Groupdept_shown_VN`) whose `Parentkey_revise` is blank — including, misleadingly, the real company root (Christian Nielsen) himself, since he's *also* `Entity=SonionAS` with a blank parent (correctly, as the top of the tree). The **"expected exactly 1 blank-parent root" WARNING is therefore expected to fire on every run of this export** — don't treat it as a sign something broke. The script's existing orphan-with-most-children fallback correctly still picks Christian Nielsen (he's referenced as parent by thousands of rows; the ~740 broken rows are leaf orphans referenced by no one), and the "Unlinked employees" business rule above already handles the rest by labeling them `— unlinked (manager not in this export)` — they show up in the department dropdown as a large `Unassigned — unlinked...` bucket (~680 people) plus smaller `<Dept> — unlinked...` buckets. If this data gap gets fixed upstream in a future Power BI refresh, the WARNING and the unlinked buckets should shrink or disappear — worth a quick sanity check next time this is run.

---

## IF THE COLUMN LAYOUT CHANGES

The script assumes a header row containing (by name, any order): `MasterlistName_Initial`, `MasterlistEmployee Key`, `MasterlistParentkey_revise`, `MasterlistGroupdept_shown_VN`, `MasterlistChức vụ`, `MasterlistJob_type`, `MasterlistEntity`.

If Power BI drops/renames one of these columns (or reverts to the old `Masterlist[...]`-bracketed naming style), the script exits with a clear error naming the missing column(s) and the actual headers found — update the `REQUIRED` list and the corresponding `g(row, "...")` calls in `scripts/generate_orgchart.py` to match.
