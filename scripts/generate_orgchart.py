"""
Generate html/orgchart.html from raw_docs/Orgchart.csv.

Usage:
    python scripts/generate_orgchart.py

See prompt-orgchart-update.md at the repo root for the full workflow this
script is part of, and for the non-obvious business rules baked in below
(Jason King relabeling, Nguyễn Huề's department consolidation, DL exclusion,
entity baseline, etc.) — read that file before changing this script's logic.
"""
import csv
import hashlib
import shutil
import sys
import tempfile
import time
from collections import Counter, defaultdict

# Windows terminals often default to a legacy codepage (cp1252) that can't
# encode Vietnamese diacritics — without this, printing a name like "Nguyễn"
# crashes the script with UnicodeEncodeError even though the actual
# csv-reading and HTML-writing logic below is unaffected.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path

import json

REPO_ROOT = Path(__file__).resolve().parent.parent
SRC = REPO_ROOT / "raw_docs" / "Orgchart.csv"
OUT = REPO_ROOT / "html" / "orgchart.html"
# Second copy, written alongside the repo copy above — this path is the local
# OneDrive/SharePoint sync mirror of the "Shared Documents" library at
# https://sonion.sharepoint.com/departments/hr/vn/Shared%20Documents/Forms/Userview.aspx
# (confirmed 2026-08-17: same site as raw_docs/wiki_employee's own sync
# folder, just a different, sibling document library). Writing here lets
# OneDrive auto-sync the file up to that SharePoint location; the repo copy
# (OUT above) still exists separately so STEP 4's git history/revert safety
# net is unaffected.
OUT_SHAREPOINT = Path(r"C:\Users\vyho\Sonion A S\HR VN - Documents") / "orgchart.html"


def load_csv_rows_stable(path: Path, max_attempts: int = 5, interval: float = 2.0):
    """Read `path` safely, in two senses that have both bitten this project before
    (originally written for the old .xlsx export, same risks apply to this CSV export):

    1. Locked file — the exporting tool (Excel/Power BI) can keep an exclusive-ish
       lock while the file is open on someone's desktop, so reading it directly can
       raise PermissionError. Always work from a plain filesystem copy instead (a
       read-only copy succeeds even while the original is open elsewhere, since that
       lock is a share-mode, not a read-block) — this reads whatever was last
       *saved*, so the file must be saved first, not just edited.
    2. Mid-save file — the export can still be WRITING the file when this script
       runs, so a single copy can capture a half-written snapshot with silently
       wrong data (this has happened: row counts and even which columns exist
       changed between two reads taken seconds apart). Copy the file repeatedly and
       only proceed once two consecutive copies, `interval` seconds apart, are
       byte-identical.
    """
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        prev_hash = None
        copy_path = None
        for attempt in range(1, max_attempts + 1):
            copy_path = tmp / f"snapshot_{attempt}.csv"
            shutil.copy2(path, copy_path)
            h = hashlib.md5(copy_path.read_bytes()).hexdigest()
            if prev_hash == h:
                print(f"'{path.name}' is stable (unchanged over {interval}s) — proceeding.")
                with open(copy_path, encoding="utf-8-sig", newline="") as f:
                    reader = csv.DictReader(f)
                    return reader.fieldnames, list(reader)
            if attempt < max_attempts:
                if attempt > 1:
                    print(f"'{path.name}' changed since the last check (likely still being "
                          f"exported) — waiting {interval}s and re-checking ({attempt}/{max_attempts})...")
                else:
                    print(f"Checking '{path.name}' is stable before reading ({attempt}/{max_attempts})...")
                time.sleep(interval)
            prev_hash = h
        print(f"WARNING: '{path.name}' did not settle after {max_attempts} checks — "
              "proceeding with the latest snapshot anyway. If today's output looks off "
              "(missing people, wrong department counts), re-export the file "
              "and re-run this script.")
        with open(copy_path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            return reader.fieldnames, list(reader)


headers, csv_rows = load_csv_rows_stable(SRC)

REQUIRED = ["MasterlistName_Initial", "MasterlistEmployee Key",
            "MasterlistParentkey_revise", "MasterlistGroupdept_shown_VN",
            "MasterlistNew Department", "MasterlistChức vụ", "MasterlistJob_type",
            "MasterlistEntity"]
missing = [r for r in REQUIRED if r not in headers]
if missing:
    sys.exit(
        "ERROR: Orgchart.csv is missing expected column(s): " + ", ".join(missing) +
        f"\nActual columns found: {headers}\n"
        "This export's column layout has changed before (this script reads by header "
        "NAME, not position, specifically so re-ordering doesn't break it) — if a "
        "column was renamed or removed, this script's REQUIRED list / the g() lookups "
        "below need a matching update. See prompt-orgchart-update.md."
    )


def g(row, name):
    return row.get(name)


raw_rows = []
for row in csv_rows:
    key = g(row, "MasterlistEmployee Key")
    if key in (None, ""):
        continue
    name = (g(row, "MasterlistName_Initial") or "").strip()
    parent = g(row, "MasterlistParentkey_revise")
    parent = str(parent).strip() if parent not in (None, "") else ""
    dept_raw = (g(row, "MasterlistGroupdept_shown_VN") or "").strip()
    # Some rows carry a pipe-separated breadcrumb (e.g. "ME & EE & OPM|VNII|...")
    # instead of one clean department name — only the first segment is real.
    dept = dept_raw.split("|")[0].strip() if dept_raw else ""
    # MasterlistNew Department is a separate, per-person column used ONLY for
    # the Department Org Chart tab (and the Full Org Chart print view, which
    # shares its grouping) — a direct text grouping, not grounded through the
    # real manager chain like `dept`/`dept_label` below. It's kept for that
    # narrower purpose because, unlike Groupdept_shown_VN, this column is
    # clean/complete enough (no blanks, no pipe-breadcrumbs as of 2026-08-14)
    # to trust directly without the structural-grounding workaround.
    new_dept_raw = (g(row, "MasterlistNew Department") or "").strip()
    new_dept = new_dept_raw.split("|")[0].strip() if new_dept_raw else ""
    title = (g(row, "MasterlistChức vụ") or "").strip()
    jobtype = (g(row, "MasterlistJob_type") or "").strip()
    entity = (g(row, "MasterlistEntity") or "").strip()
    raw_rows.append({"key": str(key).strip(), "name": name, "parent": parent,
                      "dept": dept, "new_dept": new_dept, "title": title,
                      "jobtype": jobtype, "entity": entity})

print("Total raw rows:", len(raw_rows))
dupkeys = Counter(r["key"] for r in raw_rows)
dups = {k: v for k, v in dupkeys.items() if v > 1}
if dups:
    print("WARNING duplicate Employee Keys found (last row wins in lookups):", dups)

full_children = defaultdict(list)
for r in raw_rows:
    full_children[r["parent"]].append(r)

by_key_all = {r["key"]: r for r in raw_rows}

# Jason King heads ALL Vietnam employees (VNI + VNII operations) — the source file
# tags him with dept="HR & Admin", which would wrongly lump him into that
# department's chart/count. Give him his own distinct label instead.
for r in raw_rows:
    if r["name"].startswith("Jason King"):
        r["dept"] = "VN Operations (Country Head)"
        print("Relabeled department for", r["name"], "-> ", r["dept"])

# --- Consolidated department: grounded ENTIRELY on ORG STRUCTURE (real
# Employee Key -> Parentkey_revise chains, walked over EVERY employee
# including rank-and-file DL operators — so headcounts stay true even though
# DL people never get their own card), never on raw Group_dept text matching.
# Two people can share identical Group_dept text while being on completely
# unrelated management chains (e.g. Natalie Kwik's own row says "Finance
# (904)" but her real chain is Kwik -> Andreas Loibnegger (CTIO) -> Christian
# Nielsen (CEO) — nothing to do with Nguyễn Hạnh Thảo's actual Vietnam
# Finance team). Grounding on text alone would silently merge two unrelated
# real teams just because a spreadsheet cell happens to match.

# Manual name overrides for heads whose own raw dept text doesn't reflect
# what their team actually is. Default (no override) = the head's own dept,
# or the head's own name if their dept is blank.
CONSOLIDATED_NAME_OVERRIDES = {
    "Jason King": "VN Operations (Country Head)",
    "Nguyễn Huề": "Digital Transformation & VNII Factory",
}


def bfs_subtree(head):
    """All real descendants of `head` (every employee, any job type), head included."""
    queue = [head]
    seen = {}
    while queue:
        cur = queue.pop()
        if cur["key"] in seen:
            continue
        seen[cur["key"]] = cur
        queue.extend(full_children.get(cur["key"], []))
    return seen


def consolidate_heads(heads, label):
    """For a sibling group of heads (real direct reports of one manager),
    assign each head's own name/dept as the default bucket name, but
    disambiguate with the head's own name whenever two siblings in this
    SAME group would otherwise collide on an identical default name."""
    default_name = {}
    for h in heads:
        short = h["name"].split(" (")[0]
        default_name[h["key"]] = CONSOLIDATED_NAME_OVERRIDES.get(short, h["dept"] or h["name"])
    name_counts = Counter(default_name.values())
    result = {}
    print(f"\nConsolidating department-under-{label} structure:")
    for h in heads:
        name = default_name[h["key"]]
        if name_counts[name] > 1:
            name = f"{name} — {h['name']}"
        subtree = bfs_subtree(h)
        for key in subtree:
            result[key] = name
        print(f"  {h['name']} -> '{name}' ({len(subtree)} people incl. head)")
    return result


blank_parent_roots = [r for r in raw_rows if r["parent"] == ""]
company_root_key = blank_parent_roots[0]["key"] if len(blank_parent_roots) == 1 else None
if company_root_key is None:
    children_count = Counter(r["parent"] for r in raw_rows)
    orphans = [r for r in raw_rows if r["parent"] not in by_key_all]
    orphans.sort(key=lambda r: -children_count[r["key"]])
    company_root_key = orphans[0]["key"]
    print("WARNING: expected exactly 1 blank-parent root, found", len(blank_parent_roots),
          "-> falling back to orphan with most children:", by_key_all[company_root_key]["name"])

consolidated_override = {}
# Tier 1: every real direct report of the company root (broad-stamps their
# WHOLE subtree, including Jason King's — overwritten more granularly next).
tier1_heads = full_children.get(company_root_key, [])
consolidated_override.update(consolidate_heads(tier1_heads, "Company Root"))

# Tier 2: Jason King's own direct reports each become their own department
# head (overwrites the broad "VN Operations" stamp tier 1 gave his subtree).
jason_node = next((r for r in raw_rows if r["name"].startswith("Jason King")), None)
if jason_node:
    tier2_heads = full_children.get(jason_node["key"], [])
    consolidated_override.update(consolidate_heads(tier2_heads, "Jason King"))
else:
    print("WARNING: 'Jason King' not found in this export — tier-2 department "
          "consolidation was skipped. See prompt-orgchart-update.md.")

# SVN root — the "SVN Org Chart" tab shows Sonion Vietnam only, rooted at
# Jason King (VN Operations Country Head) rather than the global company
# root (Christian Nielsen). This is purely a display choice for that one
# tab; company_root_key above remains Christian Nielsen for every other
# purpose (Tier 1 consolidation, dept_label grounding, headcount totals),
# since those must stay grounded on the true full-company tree.
svn_root_key = jason_node["key"] if jason_node else company_root_key
if not jason_node:
    print("WARNING: 'Jason King' not found — SVN Org Chart tab falls back to "
          "the global company root instead.")

# Breakdown is grounded on real org structure (above), applied to EVERY
# employee. Anyone whose true chain never reaches the company root within
# this export (a broken/missing manager key upstream) is explicitly flagged
# as unlinked rather than silently falling back to raw dept text, which
# could coincidentally match — and get merged into — a real, unrelated
# consolidated department.
BASELINE_ENTITY = "VNI"
for r in raw_rows:
    if r["key"] == company_root_key:
        r["dept_label"] = r["dept"] if r["dept"] else "Unassigned"
    elif r["key"] in consolidated_override:
        r["dept_label"] = consolidated_override[r["key"]]
    else:
        base = r["dept"] if r["dept"] else "Unassigned"
        r["dept_label"] = f"{base} — unlinked (manager not in this export)"
    r["entity_flag"] = r["entity"] if r["entity"] and r["entity"] != BASELINE_ENTITY else ""
    # Department Org Chart tab's grouping — direct per-person text from
    # MasterlistNew Department, no structural grounding/consolidation applied
    # (see the note where new_dept is read above for why this one's trusted
    # directly). Falls back to "Unassigned" only if the column is ever blank.
    r["new_dept_label"] = r["new_dept"] if r["new_dept"] else "Unassigned"

# True headcount per department — counts EVERY employee, including DL
# operators who will never get their own card.
dept_counts = Counter(r["dept_label"] for r in raw_rows)
dept_list = sorted(dept_counts.keys())
new_dept_counts = Counter(r["new_dept_label"] for r in raw_rows)
new_dept_list = sorted(new_dept_counts.keys())

# True TOTAL subordinate count per person, broken down by Job_type
# (IDL/DLS/DL) — every person below them at ANY depth, not just direct
# reports. This is deliberately independent of the UI's collapse/expand
# state: "Collapse to level 1" only controls what's rendered on screen, the
# report-count badge always reflects the person's whole real reporting line.
sys.setrecursionlimit(10000)  # org depth is small in practice, but guard generously
_subtree_memo = {}


def subtree_breakdown(key):
    if key in _subtree_memo:
        return _subtree_memo[key]
    total = Counter()
    for child in full_children.get(key, []):
        total[child["jobtype"] or "?"] += 1
        for jt, cnt in subtree_breakdown(child["key"]).items():
            total[jt] += cnt
    _subtree_memo[key] = total
    return total


report_breakdown = {r["key"]: dict(subtree_breakdown(r["key"])) for r in raw_rows}

# Render set: DL (rank-and-file production) employees never get their own
# card / are never expandable, per explicit decision — but they're already
# counted above (dept_counts, report_breakdown) before this filter runs, so
# headcounts stay true to the full company.
nodes = [r for r in raw_rows if r["jobtype"] != "DL"]
print("\nRenderable nodes (never expand into DL):", len(nodes),
      " excluded DL employees:", len(raw_rows) - len(nodes))
by_key = {n["key"]: n for n in nodes}

print("Company root:", company_root_key, by_key_all[company_root_key]["name"], by_key_all[company_root_key]["title"])
print("SVN root:", svn_root_key, by_key_all[svn_root_key]["name"], by_key_all[svn_root_key]["title"])
print("Departments — structural, used for SVN Org Chart card labels (true headcount incl. DL):")
for d in dept_list:
    print(f"  {d}: {dept_counts[d]}")
print("Departments — Department Org Chart tab, direct from MasterlistNew Department (true headcount incl. DL):")
for d in new_dept_list:
    print(f"  {d}: {new_dept_counts[d]}")

data_json = json.dumps({
    "nodes": [{"key": n["key"], "parent": n["parent"], "name": n["name"],
               "title": n["title"], "dept_label": n["dept_label"],
               "new_dept_label": n["new_dept_label"],
               "entity_flag": n["entity_flag"],
               "report_breakdown": report_breakdown.get(n["key"], {})} for n in nodes],
    "companyRoot": company_root_key,
    "svnRoot": svn_root_key,
    "depts": dept_list,
    "deptTotals": dept_counts,
    "newDepts": new_dept_list,
    "newDeptTotals": new_dept_counts,
}, ensure_ascii=False)

HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Sonion Vietnam — Org Chart</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800&family=Raleway:wght@400;500;600&display=swap');

  :root{
    /* Sonion brand palette */
    --orange:#EB5F0A; --orange-red:#E74011; --black:#000000; --white:#FFFFFF;
    --charcoal:#262626; --grey-mid:#BFBFBF; --grey-text:#808080;
    --sky:#95D3E6; --lime:#C7D64F; --cream:#FBF8E9;

    --bg: var(--white); --card-bg:#ffffff; --card-border:#d9d9d9; --text:var(--black); --muted:var(--grey-text);
    --accent:var(--orange); --line:var(--grey-mid); --hover:#f2f2f2; --shadow:0 1px 3px rgba(0,0,0,0.10);
    --font-display:'Montserrat','Arial Black',Arial,sans-serif;
    --font-label:'Montserrat',Arial,sans-serif;
    --font-body:'Raleway',Calibri,Arial,sans-serif;
  }
  @media (prefers-color-scheme: dark){
    :root{
      --bg:var(--charcoal); --card-bg:#333333; --card-border:#4a4a4a; --text:#ffffff; --muted:#c7c7c7;
      --accent:#ff8a3d; --line:#5a5a5a; --hover:#3d3d3d; --shadow:0 1px 3px rgba(0,0,0,0.5);
    }
  }
  :root[data-theme="dark"]{
    --bg:var(--charcoal); --card-bg:#333333; --card-border:#4a4a4a; --text:#ffffff; --muted:#c7c7c7;
    --accent:#ff8a3d; --line:#5a5a5a; --hover:#3d3d3d; --shadow:0 1px 3px rgba(0,0,0,0.5);
  }
  :root[data-theme="light"]{
    --bg:var(--white); --card-bg:#ffffff; --card-border:#d9d9d9; --text:var(--black); --muted:var(--grey-text);
    --accent:var(--orange); --line:var(--grey-mid); --hover:#f2f2f2; --shadow:0 1px 3px rgba(0,0,0,0.10);
  }
  *{box-sizing:border-box;}
  body{
    margin:0; background:var(--bg); color:var(--text);
    font-family:var(--font-body);
  }
  header{
    position:sticky; top:0; z-index:20; background:var(--card-bg); border-bottom:1px solid var(--card-border);
    padding:14px 24px 0;
  }
  .header-top{ padding-bottom:12px; }
  .eyebrow{
    font-family:var(--font-label); font-weight:700; font-size:10.5px; letter-spacing:2px;
    text-transform:uppercase; color:var(--muted); margin:0 0 4px;
  }
  header h1{
    font-family:var(--font-display); font-weight:800; font-size:22px; letter-spacing:0.3px;
    text-transform:uppercase; color:var(--accent); margin:0 0 10px; line-height:1.1;
  }
  .orange-rule{ width:220px; max-width:45%; height:2px; background:var(--orange); margin:0 0 14px; border:0; }
  .header-toolbar{
    display:flex; align-items:center; gap:16px; flex-wrap:wrap; padding-bottom:12px;
  }
  .tabs{ display:flex; gap:6px; }
  .tab-btn{
    font-family:var(--font-label); font-weight:700; font-size:12.5px;
    border:1px solid var(--card-border); background:var(--card-bg); color:var(--text);
    padding:7px 16px; border-radius:3px; cursor:pointer;
  }
  .tab-btn.active{ background:var(--orange); color:#fff; border-color:var(--orange); }
  select#deptSelect{
    font-family:var(--font-body); padding:6px 10px; border-radius:3px; border:1px solid var(--card-border);
    background:var(--card-bg); color:var(--text); font-size:13px; max-width:320px;
  }
  .toolbar-group{ display:flex; align-items:center; gap:8px; }
  .toolbar-group.hidden{ display:none; }
  button.util{
    font-family:var(--font-label); font-weight:700;
    border:1px solid var(--grey-mid); background:var(--card-bg); color:var(--text);
    padding:6px 12px; border-radius:3px; cursor:pointer; font-size:11.5px;
  }
  button.util:hover{ background:var(--hover); }
  main{ padding:28px 24px 80px; overflow-x:auto; }
  .chart-wrap{ display:none; }
  .chart-wrap.active{ display:block; }
  .meta-line{ color:var(--muted); font-size:12.5px; margin:0 0 18px 2px; }

  /* Tree */
  ul.tree, ul.tree ul{
    list-style:none; margin:0; padding:0; display:flex; position:relative;
  }
  ul.tree{ padding-top:0; }
  ul.tree ul{ padding-top:32px; }
  ul.tree li{
    position:relative; display:flex; flex-direction:column; align-items:center;
    padding:0 10px; text-align:center;
  }
  /* connector: vertical line above every node except top-level roots */
  ul.tree ul li::before{
    content:""; position:absolute; top:0; left:50%; width:1px; height:32px;
    background:var(--line); transform:translateX(-0.5px);
  }
  /* horizontal connector across siblings, drawn on top of each li's own line */
  ul.tree ul li::after{
    content:""; position:absolute; top:0; height:1px; background:var(--line);
    left:0; right:0;
  }
  ul.tree ul li:only-child::after{ content:none; }
  ul.tree ul li:first-child::after{ left:50%; }
  ul.tree ul li:last-child::after{ right:50%; }

  .card{
    display:inline-flex; flex-direction:column; gap:2px; min-width:150px; max-width:210px;
    background:var(--card-bg); border:1px solid var(--card-border); border-radius:4px;
    padding:8px 12px; box-shadow:var(--shadow); position:relative; z-index:1; cursor:default;
  }
  .card .name{ font-family:var(--font-label); font-weight:700; font-size:13px; line-height:1.25; }
  .card .title{ font-family:var(--font-body); font-size:11.5px; color:var(--muted); line-height:1.25; }
  .card .dept{ font-family:var(--font-label); font-size:10.5px; color:var(--accent); font-weight:700; margin-top:2px; }
  .card .entity-flag{
    font-family:var(--font-label); font-size:9.5px; font-weight:700; text-transform:uppercase;
    letter-spacing:0.4px; color:var(--orange-red); background:rgba(231,64,17,0.12);
    border-radius:8px; padding:1px 7px; margin-top:3px; align-self:flex-start;
  }

  /* Diagram hierarchy palette (Sonion brand: orange apex, sky-blue level 2, lime level 3, neutral beyond) */
  .card.tier-0{ background:var(--orange); border-color:var(--orange); }
  .card.tier-0 .name, .card.tier-0 .title, .card.tier-0 .dept{ color:#fff; }
  .card.tier-0 .entity-flag{ background:rgba(255,255,255,0.28); color:#fff; }
  .card.tier-1{ background:var(--sky); border-color:var(--sky); }
  .card.tier-1 .name, .card.tier-1 .title, .card.tier-1 .dept{ color:var(--black); }
  .card.tier-2{ background:var(--lime); border-color:var(--lime); }
  .card.tier-2 .name, .card.tier-2 .title, .card.tier-2 .dept{ color:var(--black); }

  .toggle{
    font-family:var(--font-label); font-weight:700;
    margin-top:6px; border:1px solid var(--grey-mid); background:var(--card-bg); color:var(--text);
    border-radius:12px; font-size:10.5px; padding:2px 10px; cursor:pointer; align-self:center;
  }
  .toggle:hover{ background:var(--hover); }
  .toggle.dl-count{
    cursor:default; font-style:italic; font-weight:500; color:var(--muted); border-style:dashed;
  }
  .toggle.dl-count:hover{ background:var(--card-bg); }

  .children-wrap.collapsed{ display:none; }

  footer.note{
    background:var(--black); color:#fff; font-family:var(--font-body); font-size:11px;
    letter-spacing:0.4px; padding:12px 24px; margin-top:20px;
  }

  /* Full-company print view — same box/connector tree template as the
     Department Org Chart tab, just laid out as one section per department. */
  .print-dept-section{ margin:0 0 34px; }
  .print-tree-scalewrap{ overflow:hidden; width:100%; }
  .print-tree-inner{ display:inline-block; }
  .print-dept-title{
    font-family:var(--font-display); font-weight:800; font-size:14px; color:#fff;
    background:var(--orange); padding:7px 12px; text-transform:uppercase; letter-spacing:0.3px;
    margin:0 0 16px; break-after:avoid;
  }

  @media print{
    .no-print{ display:none !important; }
    header{ position:static; }
    main{ padding:0 6px; overflow:visible; }
    body{ background:#fff; }
    .card{ box-shadow:none; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .print-dept-title{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .print-dept-section{ break-before:page; }
    .print-dept-section:first-child{ break-before:auto; }
  }
</style>
</head>
<body>
<header>
  <div class="header-top">
    <p class="eyebrow">Sonion Vietnam · Human Resources</p>
    <h1>Org Chart</h1>
    <hr class="orange-rule" />
  </div>
  <div class="header-toolbar no-print">
    <div class="tabs">
      <button class="tab-btn active" data-tab="svn">SVN Org Chart</button>
      <button class="tab-btn" data-tab="dept">Department Org Chart</button>
      <button class="tab-btn" data-tab="print">Full Org Chart (All Departments)</button>
    </div>
    <div class="toolbar-group hidden" id="deptToolbar">
      <select id="deptSelect"></select>
    </div>
    <div class="toolbar-group" id="expandCollapseToolbar">
      <button class="util" id="expandAllBtn">Expand all</button>
      <button class="util" id="collapseAllBtn">Collapse to level 1</button>
    </div>
    <div class="toolbar-group hidden" id="printToolbar">
      <button class="util" id="printBtn">Print this page</button>
    </div>
  </div>
</header>
<main>
  <div class="chart-wrap active" id="wrap-svn">
    <p class="meta-line" id="svnMeta"></p>
    <div id="svnChart"></div>
  </div>
  <div class="chart-wrap" id="wrap-dept">
    <p class="meta-line" id="deptMeta"></p>
    <div id="deptChart"></div>
  </div>
  <div class="chart-wrap" id="wrap-print">
    <p class="meta-line no-print">Every department, one section per department, collapsed to head + direct reports by default — for printing the whole company org chart in one go. Each section auto-fits the page width. Each department starts on a new printed page.</p>
    <div id="printAllChart"></div>
  </div>
</main>
<footer class="note">Generated from raw_docs/Orgchart.csv (Masterlist export). Every report-count badge shows that person's TOTAL subordinates at every level down (not just direct reports), broken down by job type (IDL / DLS / DL) — this count never changes with the collapse/expand state, which only controls what's rendered on screen. Direct Labor (DL) operators are always counted but never get their own card — a dashed "(not shown)" badge appears instead wherever every one of a person's direct reports is DL. Click a card's button to expand its team; "Collapse to level 1" resets a view to the head and direct reports only.</footer>

<script id="org-data" type="application/json">__DATA_JSON__</script>
<script>
(function(){
  const DATA = JSON.parse(document.getElementById('org-data').textContent);
  const nodes = DATA.nodes;
  const byKey = {};
  nodes.forEach(n => byKey[n.key] = n);
  const childrenByParent = {};
  nodes.forEach(n => {
    if (!childrenByParent[n.parent]) childrenByParent[n.parent] = [];
    childrenByParent[n.parent].push(n);
  });
  Object.values(childrenByParent).forEach(list => list.sort((a,b)=>a.name.localeCompare(b.name)));

  function getChildren(key){ return childrenByParent[key] || []; }

  // TOTAL subordinate count (every level down, not just direct reports)
  // broken down by Job_type (IDL/DLS/DL), e.g. "2 IDL · 1 DLS · 30 DL".
  const JOBTYPE_ORDER = ['IDL', 'DLS', 'DL'];
  function formatBreakdown(breakdown){
    if (!breakdown) return '';
    const parts = JOBTYPE_ORDER.filter(t => breakdown[t]).map(t => breakdown[t] + ' ' + t);
    Object.keys(breakdown).forEach(t => {
      if (!JOBTYPE_ORDER.includes(t)) parts.push(breakdown[t] + ' ' + t);
    });
    return parts.join(' · ');
  }

  function renderNode(node, depth, showDept, defaultExpandDepth, childrenLookup){
    const kids = childrenLookup(node.key);
    const li = document.createElement('li');

    const card = document.createElement('div');
    const tier = depth <= 2 ? ('tier-' + depth) : '';
    card.className = ('card ' + tier).trim();
    const nameEl = document.createElement('div');
    nameEl.className = 'name';
    nameEl.textContent = node.name || '(No name)';
    const titleEl = document.createElement('div');
    titleEl.className = 'title';
    titleEl.textContent = node.title || '';
    card.appendChild(nameEl);
    card.appendChild(titleEl);
    if (showDept){
      const deptEl = document.createElement('div');
      deptEl.className = 'dept';
      deptEl.textContent = node.dept_label || '';
      card.appendChild(deptEl);
    }
    if (node.entity_flag){
      const entityEl = document.createElement('div');
      entityEl.className = 'entity-flag';
      entityEl.textContent = node.entity_flag + ' entity';
      card.appendChild(entityEl);
    }
    li.appendChild(card);

    // report_breakdown is the TRUE TOTAL subordinate count (every level
    // down, every job type incl. DL) — independent of collapse/expand state,
    // which only controls what's rendered, never the count. `kids` is only
    // the renderable (non-DL) direct-children subset used for the expand
    // toggle itself; DL reports are counted but never get their own card.
    const breakdownText = formatBreakdown(node.report_breakdown);
    const hasReports = breakdownText !== '';
    if (kids.length){
      const toggle = document.createElement('button');
      toggle.className = 'toggle';
      const expanded = depth < defaultExpandDepth;
      const label = '+ ' + breakdownText;
      toggle.textContent = expanded ? '− collapse' : label;
      li.appendChild(toggle);

      const childUl = document.createElement('ul');
      childUl.className = 'children-wrap' + (expanded ? '' : ' collapsed');
      kids.forEach(k => childUl.appendChild(renderNode(k, depth+1, showDept, defaultExpandDepth, childrenLookup)));
      li.appendChild(childUl);

      toggle.addEventListener('click', () => {
        const isCollapsed = childUl.classList.toggle('collapsed');
        toggle.textContent = isCollapsed ? label : '− collapse';
      });
    } else if (hasReports){
      // Every direct report is DL (rank-and-file) — never expandable, but
      // still counted so the headcount isn't silently lost.
      const badge = document.createElement('span');
      badge.className = 'toggle dl-count';
      badge.textContent = breakdownText + ' (not shown)';
      li.appendChild(badge);
    }
    return li;
  }

  function renderForest(containerEl, roots, showDept, defaultExpandDepth, childrenLookup){
    containerEl.innerHTML = '';
    const ul = document.createElement('ul');
    ul.className = 'tree';
    roots.forEach(r => ul.appendChild(renderNode(r, 0, showDept, defaultExpandDepth, childrenLookup)));
    containerEl.appendChild(ul);
  }

  const svnRoot = byKey[DATA.svnRoot];
  const svnContainer = document.getElementById('svnChart');
  function drawSvn(defaultExpandDepth){
    renderForest(svnContainer, [svnRoot], true, defaultExpandDepth, getChildren);
  }
  drawSvn(1);
  document.getElementById('svnMeta').textContent =
    `Root: ${svnRoot.name} — ${svnRoot.title} (Sonion Vietnam). Showing head + direct reports by default; use a card's button to expand further, or "Expand all" above.`;

  const deptSelect = document.getElementById('deptSelect');
  DATA.newDepts.forEach(d => {
    // newDeptTotals is the TRUE headcount (incl. DL); nodes.filter(...).length
    // would only count renderable (non-DL) people and undercount it.
    const count = DATA.newDeptTotals[d] || 0;
    const opt = document.createElement('option');
    opt.value = d;
    opt.textContent = `${d} (${count})`;
    deptSelect.appendChild(opt);
  });

  // Members of one dept bucket, restricted so traversal never crosses into a
  // different bucket just because the raw reporting line connects them
  // (e.g. a SonionAS-entity person managing otherwise-VNI staff).
  function getDeptGroup(dept){
    const members = nodes.filter(n => n.new_dept_label === dept);
    const memberKeys = new Set(members.map(m => m.key));
    const roots = members.filter(m => !memberKeys.has(m.parent));
    const childrenLookup = (key) => (childrenByParent[key] || []).filter(c => memberKeys.has(c.key));
    return { members, roots, childrenLookup };
  }

  const deptContainer = document.getElementById('deptChart');
  let currentDeptExpandDepth = 1;
  function drawDept(){
    const dept = deptSelect.value;
    const { roots, childrenLookup } = getDeptGroup(dept);
    const total = DATA.newDeptTotals[dept] || 0;
    document.getElementById('deptMeta').textContent =
      `${dept} — ${total} member(s) incl. DL, ${roots.length} head${roots.length>1?'s':''}.`;
    renderForest(deptContainer, roots, false, currentDeptExpandDepth, childrenLookup);
  }
  deptSelect.addEventListener('change', drawDept);
  if (DATA.newDepts.length){ deptSelect.value = DATA.newDepts[0]; drawDept(); }

  // ---- Full Org Chart (print) — every department, using the SAME
  // box/connector tree template as the Department Org Chart tab (just one
  // section per department instead of picking one from a dropdown).
  // Defaults to collapsed at level 1 (head + direct reports), and every
  // section auto-scales down to fit the page width instead of overflowing
  // — recalculated live whenever a toggle inside it expands/collapses. ----
  function fitToWidth(scaleWrap, inner){
    inner.style.transform = 'none';
    scaleWrap.style.height = 'auto';
    const naturalWidth = inner.scrollWidth;
    const naturalHeight = inner.scrollHeight;
    const availableWidth = scaleWrap.clientWidth;
    if (naturalWidth > availableWidth && availableWidth > 0){
      const scale = availableWidth / naturalWidth;
      inner.style.transform = `scale(${scale})`;
      inner.style.transformOrigin = 'top left';
      scaleWrap.style.height = (naturalHeight * scale) + 'px';
    }
  }

  let currentPrintExpandDepth = 1;
  function drawPrintAll(){
    const container = document.getElementById('printAllChart');
    container.innerHTML = '';
    DATA.newDepts.forEach(dept => {
      const { roots, childrenLookup } = getDeptGroup(dept);
      const section = document.createElement('div');
      section.className = 'print-dept-section';
      const title = document.createElement('h2');
      title.className = 'print-dept-title';
      title.textContent = `${dept} (${DATA.newDeptTotals[dept] || 0})`;
      section.appendChild(title);
      const scaleWrap = document.createElement('div');
      scaleWrap.className = 'print-tree-scalewrap';
      const treeContainer = document.createElement('div');
      treeContainer.className = 'print-tree-inner';
      scaleWrap.appendChild(treeContainer);
      section.appendChild(scaleWrap);
      renderForest(treeContainer, roots, false, currentPrintExpandDepth, childrenLookup);
      container.appendChild(section);

      const fit = () => fitToWidth(scaleWrap, treeContainer);
      fit();
      new ResizeObserver(fit).observe(treeContainer);
    });
  }
  document.getElementById('printBtn').addEventListener('click', () => window.print());

  // Persist the active tab (sessionStorage — survives a page reload but not
  // a new browser session) so that if this file gets reloaded out from under
  // the user — a live-preview/auto-refresh extension reloading on every save
  // is the usual cause — it comes back to the tab they were on instead of
  // silently resetting to SVN every time.
  let printDrawn = false;
  const TAB_STORAGE_KEY = 'orgchart_active_tab';
  const tabBtns = document.querySelectorAll('.tab-btn');

  function activateTab(tab, { persist = true } = {}){
    tabBtns.forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.getElementById('wrap-svn').classList.toggle('active', tab === 'svn');
    document.getElementById('wrap-dept').classList.toggle('active', tab === 'dept');
    document.getElementById('wrap-print').classList.toggle('active', tab === 'print');
    document.getElementById('deptToolbar').classList.toggle('hidden', tab !== 'dept');
    document.getElementById('printToolbar').classList.toggle('hidden', tab !== 'print');
    if (tab === 'print' && !printDrawn){ drawPrintAll(); printDrawn = true; }
    if (persist){
      try { sessionStorage.setItem(TAB_STORAGE_KEY, tab); } catch (e) { /* storage unavailable, e.g. file:// in some browsers — ignore */ }
    }
  }

  tabBtns.forEach(btn => btn.addEventListener('click', () => activateTab(btn.dataset.tab)));

  let restoredTab = null;
  try { restoredTab = sessionStorage.getItem(TAB_STORAGE_KEY); } catch (e) { /* ignore */ }
  if (restoredTab && ['svn', 'dept', 'print'].includes(restoredTab) && restoredTab !== 'svn'){
    activateTab(restoredTab, { persist: false });
  }

  function activeTab(){ return document.querySelector('.tab-btn.active').dataset.tab; }
  document.getElementById('expandAllBtn').addEventListener('click', () => {
    const tab = activeTab();
    if (tab === 'svn') drawSvn(9999);
    else if (tab === 'dept') { currentDeptExpandDepth = 9999; drawDept(); }
    else { currentPrintExpandDepth = 9999; drawPrintAll(); }
  });
  document.getElementById('collapseAllBtn').addEventListener('click', () => {
    const tab = activeTab();
    if (tab === 'svn') drawSvn(1);
    else if (tab === 'dept') { currentDeptExpandDepth = 1; drawDept(); }
    else { currentPrintExpandDepth = 1; drawPrintAll(); }
  });
})();
</script>
</body>
</html>
"""

data_json_safe = data_json.replace("</", "<\\/")
html = HTML_TEMPLATE.replace("__DATA_JSON__", data_json_safe)
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(html, encoding="utf-8")
print("\nWrote", OUT, len(html), "bytes")

try:
    OUT_SHAREPOINT.parent.mkdir(parents=True, exist_ok=True)
    OUT_SHAREPOINT.write_text(html, encoding="utf-8")
    print("Wrote", OUT_SHAREPOINT, len(html), "bytes (SharePoint sync copy)")
except OSError as e:
    print(f"WARNING: could not write the SharePoint sync copy at {OUT_SHAREPOINT} "
          f"({e}) — the repo copy above was still written fine. Check the OneDrive "
          "sync folder still exists at that path and isn't locked/offline.")
