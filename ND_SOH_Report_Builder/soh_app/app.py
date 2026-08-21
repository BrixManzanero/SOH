"""
ND SOH Report Builder
=====================

Local Streamlit app that replaces the manual Clean -> Pivot -> SUMIFS
process for the TW and AGT portions of the National Distributors
stocks-on-hand report.

SCOPE: TW and AGT only.

DCR (columns G/J/M) and QS (columns K/L) stay manual - pivot and SUMIFS
in Excel. The app never touches those columns, so your work there is safe.

Five steps:
    1. Select files      (TW raw, AGT raw, template)
    2. TW cleaning       (LDU merge, whitespace, aliases)
    3. AGT preview       (lookup against the roster)
    4. Validation        (totals, unmatched items, deltas)
    5. Generate          (writes H/I and N/O only)

Run with:  streamlit run app.py
"""

import json
import os
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from soh.config import (
    load_aliases, save_alias, delete_alias,
    load_rules, add_drop, add_fold, clear_rule,
)
from soh.reconcile import reconcile
from soh.report import (
    read_roster, write_report, clear_report, COLUMN_FIELDS,
)
from soh.stock_files import parse_stock_file

st.set_page_config(page_title="ND SOH Report Builder", page_icon="📦", layout="wide")

APP_DIR = Path(__file__).resolve().parent
STATE_PATH = APP_DIR / "config" / "last_session.json"
OUTPUT_DIR = APP_DIR / "output"

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
MAX_SCAN = 4000       # safety cap when walking a folder
MAX_SHOWN = 60        # results rendered per search


# ===================================================================
# Helpers
# ===================================================================

def fmt(n) -> str:
    try:
        return f"{float(n):,.0f}"
    except (TypeError, ValueError):
        return "-"


def human_size(num_bytes: int) -> str:
    value = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:,.0f} {unit}" if unit == "B" else f"{value:,.1f} {unit}"
        value /= 1024
    return f"{value:,.1f} GB"


@st.cache_data(ttl=20, show_spinner=False)
def scan_folder(folder: str, recursive: bool):
    """
    List Excel files in a folder, newest first.

    Returns a list of dicts so Streamlit can cache it cleanly.
    Skips Excel lock files (~$...) and hidden folders.
    """
    root = Path(folder).expanduser()
    if not root.is_dir():
        return []

    found = []
    try:
        iterator = root.rglob("*") if recursive else root.iterdir()
        for path in iterator:
            if len(found) >= MAX_SCAN:
                break
            try:
                if not path.is_file():
                    continue
                if path.suffix.lower() not in EXCEL_SUFFIXES:
                    continue
                if path.name.startswith("~$") or path.name.startswith("."):
                    continue
                stat = path.stat()
                found.append({
                    "name": path.name,
                    "path": str(path),
                    "parent": str(path.parent),
                    "size": stat.st_size,
                    "mtime": stat.st_mtime,
                })
            except (OSError, PermissionError):
                continue
    except (OSError, PermissionError):
        return []

    found.sort(key=lambda f: f["mtime"], reverse=True)
    return found


def matches(entry: dict, query: str) -> bool:
    """Every whitespace-separated term must appear somewhere in the name."""
    if not query.strip():
        return True
    haystack = entry["name"].lower()
    return all(term in haystack for term in query.lower().split())


def save_session(data: dict):
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, default=str)
    except OSError:
        pass


def load_session() -> dict:
    if not STATE_PATH.exists():
        return {}
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, OSError):
        return {}


# ===================================================================
# File picker
# ===================================================================

def file_slot(slot_key: str, label: str, help_text: str, entries: list) -> str:
    """
    One searchable file slot.

    Shows the current selection with a Change button. When nothing is
    selected (or the user clicks Change), shows a search box and the
    full list of matching files - not a truncated dropdown.
    """
    state_key = f"file_{slot_key}"
    editing_key = f"editing_{slot_key}"

    selected = st.session_state.get(state_key)
    if selected and not os.path.exists(selected):
        selected = None
        st.session_state[state_key] = None

    editing = st.session_state.get(editing_key, selected is None)

    st.markdown(f"**{label}**")
    st.caption(help_text)

    if selected and not editing:
        info = st.columns([5, 1])
        with info[0]:
            path = Path(selected)
            try:
                stat = path.stat()
                meta = (
                    f"{human_size(stat.st_size)} · "
                    f"{datetime.fromtimestamp(stat.st_mtime):%b %d, %Y %I:%M %p}"
                )
            except OSError:
                meta = ""
            st.success(f"**{path.name}**")
            st.caption(f"{path.parent}  \n{meta}")
        with info[1]:
            if st.button("Change", key=f"chg_{slot_key}", use_container_width=True):
                st.session_state[editing_key] = True
                st.rerun()
        return selected

    query = st.text_input(
        "Search",
        key=f"q_{slot_key}",
        placeholder="Type part of the file name, e.g. tw soh, agt, template",
        label_visibility="collapsed",
    )

    results = [e for e in entries if matches(e, query)]

    if not results:
        st.info(
            "No matching Excel files. Try a different search term, "
            "turn on **Include subfolders**, or paste the full path below."
        )
    else:
        st.caption(
            f"{len(results)} file{'s' if len(results) != 1 else ''} found"
            + (f" · showing first {MAX_SHOWN}" if len(results) > MAX_SHOWN else "")
        )
        with st.container(height=260, border=True):
            for i, entry in enumerate(results[:MAX_SHOWN]):
                row = st.columns([6, 1])
                with row[0]:
                    st.markdown(f"**{entry['name']}**")
                    st.caption(
                        f"{human_size(entry['size'])} · "
                        f"{datetime.fromtimestamp(entry['mtime']):%b %d, %Y %I:%M %p}"
                    )
                with row[1]:
                    if st.button("Select", key=f"pick_{slot_key}_{i}",
                                 use_container_width=True):
                        st.session_state[state_key] = entry["path"]
                        st.session_state[editing_key] = False
                        st.rerun()

    manual = st.text_input(
        "Or paste the full path",
        key=f"manual_{slot_key}",
        placeholder=r"C:\Users\you\Downloads\file.xlsx",
    )
    if manual.strip():
        candidate = manual.strip().strip('"')
        if os.path.exists(candidate):
            st.session_state[state_key] = candidate
            st.session_state[editing_key] = False
            st.rerun()
        else:
            st.error("That path does not exist.")

    return st.session_state.get(state_key)


# ===================================================================
# Header
# ===================================================================

st.title("📦 ND SOH Report Builder")
st.caption(
    "Cleans the TW and AGT raw files and writes them into your template "
    "without disturbing any formatting. DCR and QS stay manual."
)
st.info(
    "**What the app fills:** TW Intransit/Onhand (H, I) and AGT "
    "Intransit/Onhand (N, O).  \n"
    "**What it never touches:** DCR (G, J, M) and QS (K, L) — your pivot "
    "and SUMIFS work there is safe."
)

session = load_session()

# ===================================================================
# STEP 1 — Select files
# ===================================================================
st.header("1 · Select files")

source_mode = st.radio(
    "Where are the files?",
    ["📁 Browse a folder on this PC", "⬆ Upload files"],
    horizontal=True,
    help=(
        "Browse is faster and needs no upload. Use Upload when the app is "
        "deployed to Streamlit Cloud."
    ),
)
upload_mode = source_mode.startswith("⬆")

tw_path = agt_path = template_path = None

if upload_mode:
    st.caption("Files stay in memory for this session only.")
    up1, up2, up3 = st.columns(3)

    def stash(uploaded):
        if uploaded is None:
            return None
        tmp_dir = Path(tempfile.gettempdir()) / "nd_soh_uploads"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        target = tmp_dir / uploaded.name
        with open(target, "wb") as fh:
            fh.write(uploaded.getbuffer())
        return str(target)

    with up1:
        tw_path = stash(st.file_uploader("TW raw", type=["xlsx", "xlsm"], key="u_tw"))
    with up2:
        agt_path = stash(st.file_uploader("AGT raw", type=["xlsx", "xlsm"], key="u_agt"))
    with up3:
        template_path = stash(
            st.file_uploader("ND SOH template", type=["xlsx", "xlsm"], key="u_tpl")
        )
    report_date = st.date_input("Report date", value=date.today())

else:
    home = Path.home()
    shortcuts = {
        "Downloads": home / "Downloads",
        "Desktop": home / "Desktop",
        "Documents": home / "Documents",
        "OneDrive Desktop": home / "OneDrive" / "Desktop",
    }
    shortcuts = {k: v for k, v in shortcuts.items() if v.is_dir()}

    if "folder" not in st.session_state:
        st.session_state["folder"] = session.get("folder") or str(
            shortcuts.get("Downloads", home)
        )

    fcol, rcol = st.columns([4, 1])
    with fcol:
        folder = st.text_input(
            "Folder to search",
            key="folder",
            help="All Excel files in this folder are searchable below.",
        )
    with rcol:
        recursive = st.checkbox(
            "Include subfolders", value=False,
            help="Searches every folder inside. Slower on large trees.",
        )

    if shortcuts:
        st.caption("Jump to:")
        chips = st.columns(len(shortcuts) + 1)
        for i, (name, path) in enumerate(shortcuts.items()):
            if chips[i].button(name, key=f"sc_{i}", use_container_width=True):
                st.session_state["folder"] = str(path)
                st.rerun()
        if chips[-1].button("🔄 Refresh", key="sc_refresh", use_container_width=True):
            scan_folder.clear()
            st.rerun()

    entries = scan_folder(folder, recursive)

    if not Path(folder).expanduser().is_dir():
        st.error(f"`{folder}` is not a folder. Check the path or use a shortcut above.")
        st.stop()

    st.caption(
        f"{len(entries)} Excel file{'s' if len(entries) != 1 else ''} in this folder"
        + (" and its subfolders" if recursive else "")
        + (f" · scan capped at {MAX_SCAN:,}" if len(entries) >= MAX_SCAN else "")
    )

    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1:
        tw_path = file_slot(
            "tw", "TW raw file",
            "The raw TW stock file. Needs Item, Intransit and Onhand columns.",
            entries,
        )
    with c2:
        agt_path = file_slot(
            "agt", "AGT raw file",
            "The AGT SOH file. If it contains a pivot sheet, the app skips it "
            "and reads the raw sheet.",
            entries,
        )
    with c3:
        template_path = file_slot(
            "template", "ND SOH template",
            "Supplies the item roster and all formatting. Copied, never modified.",
            entries,
        )

    st.divider()
    report_date = st.date_input("Report date", value=date.today())

if not (tw_path and template_path):
    st.info("Select at least the **TW raw file** and the **template** to continue.")
    st.stop()

if not upload_mode:
    save_session({"folder": st.session_state.get("folder"), "tw": tw_path,
                  "agt": agt_path, "template": template_path})

# ---- Load roster and raw files
try:
    roster = read_roster(template_path, "SOH")
except (ValueError, KeyError, OSError) as exc:
    st.error(f"Could not read the template: {exc}")
    st.stop()

try:
    tw_file = parse_stock_file(tw_path)
except (ValueError, KeyError, OSError) as exc:
    st.error(f"Could not read the TW raw file: {exc}")
    st.stop()

agt_file = None
if agt_path:
    try:
        agt_file = parse_stock_file(agt_path)
    except (ValueError, OSError) as exc:
        st.warning(f"Could not read the AGT file, continuing without it: {exc}")

m1, m2, m3 = st.columns(3)
m1.metric("Items in roster", len(roster.items))
m2.metric("TW rows", len(tw_file.rows))
m3.metric("AGT rows", len(agt_file.rows) if agt_file else 0)

# ---- Refresher
with st.expander("🧹 Refresher — reset all data to zero (price list is kept)"):
    st.caption(
        "If your template is last week's report, it still holds old DCR and QS "
        "values. Clear them all here."
    )
    r1, r2 = st.columns([2, 1])
    with r1:
        clear_groups = st.multiselect(
            "Which columns to clear",
            ["TW", "QS", "AGT", "DCR"],
            default=["TW", "QS", "AGT", "DCR"],
            help="TW = H/I · QS = K/L · AGT = N/O · DCR = G/J/M",
        )
        as_blank = st.checkbox(
            "Leave blank instead of 0", value=False,
            help="Off matches how your report looks today.",
        )
    with r2:
        st.markdown("**Always kept**")
        st.markdown(
            "- Series, Item, Market Name\n"
            "- Launch Date\n"
            "- **SRP (price list)**\n"
            "- Total formulas\n"
            "- Grand Total / Diff. qty."
        )
    if st.button("Create blank template", use_container_width=True):
        if not clear_groups:
            st.warning("Pick at least one column group.")
        else:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            blank_path = OUTPUT_DIR / "ND_SOH_BLANK_template.xlsx"
            try:
                cinfo = clear_report(
                    template_path, str(blank_path), roster,
                    groups=tuple(clear_groups), blank_instead_of_zero=as_blank,
                )
                st.success(
                    f"Cleared {cinfo['cells_cleared']} cells across "
                    f"{cinfo['rows']} rows — columns "
                    f"{', '.join(cinfo['columns_cleared'])}."
                )
                st.caption("SRP, roster, formatting and every formula are intact.")
                with open(blank_path, "rb") as fh:
                    st.download_button(
                        "⬇ Download blank template", fh.read(),
                        file_name=blank_path.name,
                        mime="application/vnd.openxmlformats-officedocument."
                             "spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except (OSError, ValueError, KeyError) as exc:
                st.error(f"Could not clear: {exc}")

# ===================================================================
# STEP 2 — TW cleaning
# ===================================================================
st.header("2 · TW cleaning")

issue_rows = [r for r in tw_file.rows if r.issues]
ldu_rows = [r for r in tw_file.rows if r.was_ldu]

t1, t2, t3 = st.columns(3)
t1.metric("LDU rows merged", len(ldu_rows))
t2.metric("Rows cleaned", len([r for r in issue_rows if not r.was_ldu]))
t3.metric("Item codes after merge", len(tw_file.merged()))

if tw_file.totals_row:
    item_i, item_o = tw_file.item_totals()
    gt_i, gt_o = tw_file.totals_row
    if tw_file.totals_match():
        st.success(
            f"Raw total matches the Grand Total row — "
            f"Intransit {fmt(item_i)}, Onhand {fmt(item_o)}."
        )
    else:
        st.error(
            f"**Totals do not match.** Item sum: Intransit {fmt(item_i)} / "
            f"Onhand {fmt(item_o)}. Grand Total row: Intransit {fmt(gt_i)} / "
            f"Onhand {fmt(gt_o)}. Do not proceed until this is resolved."
        )

non_ldu_issues = [r for r in issue_rows if not r.was_ldu]
if non_ldu_issues:
    st.subheader("Items with hidden characters")
    st.dataframe(
        pd.DataFrame([{
            "Excel row": r.excel_row,
            "Found in raw": repr(r.raw_item),
            "Cleaned to": r.item,
            "Problem": ", ".join(r.issues),
            "Intransit": r.intransit,
            "Onhand": r.onhand,
        } for r in non_ldu_issues]),
        use_container_width=True, hide_index=True,
    )

if ldu_rows:
    with st.expander(f"View the {len(ldu_rows)} LDU rows that were merged"):
        st.dataframe(
            pd.DataFrame([{
                "Excel row": r.excel_row,
                "LDU row": r.raw_item,
                "Merged into": r.item,
                "Intransit": r.intransit,
                "Onhand": r.onhand,
            } for r in ldu_rows]),
            use_container_width=True, hide_index=True,
        )

# ===================================================================
# STEP 3 — AGT preview
# ===================================================================
st.header("3 · AGT preview")

dcr_result = None   # DCR stays manual
qs_manual = None    # QS stays manual

if agt_file is None:
    st.info("No AGT file selected — columns N and O will be left at zero.")
else:
    agt_merged = agt_file.merged()
    in_roster = {k: v for k, v in agt_merged.items() if k in roster.row_of}
    total_o = sum(v["onhand"] for v in agt_merged.values())
    matched_o = sum(v["onhand"] for v in in_roster.values())

    a1, a2, a3 = st.columns(3)
    a1.metric("AGT item codes", len(agt_merged))
    a2.metric("Onhand in file", fmt(total_o))
    a3.metric("Going into report", fmt(matched_o))

    st.caption(f"Reading sheet `{agt_file.sheet_name}`")

    if abs(total_o - matched_o) > 0.5:
        st.warning(
            f"{fmt(total_o - matched_o)} units are in the AGT file but have no "
            "match in the roster. See step 4."
        )
    else:
        st.success("Every AGT unit has a matching row in the roster.")

    with st.expander("View AGT items with stock"):
        st.dataframe(
            pd.DataFrame([
                {"Item": k, "Intransit": v["intransit"], "Onhand": v["onhand"],
                 "In roster": "yes" if k in roster.row_of else "no"}
                for k, v in sorted(agt_merged.items())
                if v["intransit"] or v["onhand"]
            ]),
            use_container_width=True, hide_index=True,
        )

# ===================================================================
# STEP 4 — Validation
# ===================================================================
st.header("4 · Validation")

result = reconcile(roster.items, tw_file, agt_file, dcr_result, qs_manual)

errors = result.errors()
warnings = result.warnings()
infos = [f for f in result.findings if f.level == "info"]

v1, v2, v3 = st.columns(3)
v1.metric("Errors", len(errors))
v2.metric("Warnings", len(warnings))
v3.metric("Rules applied", len(infos))

unmatched = [f for f in errors if f.category == "unmatched"]
if unmatched:
    st.error(
        f"**{len(unmatched)} item(s) have stock but no matching row in the "
        "roster.** Answer each one — the app remembers your choice for every "
        "future report."
    )
    for finding in unmatched:
        with st.container(border=True):
            st.markdown(
                f"**`{finding.item}`** — Intransit {fmt(finding.intransit)}, "
                f"Onhand {fmt(finding.onhand)}"
            )
            options = ["— choose —"] + finding.suggestions + ["‹ pick from full list ›"]
            b1, b2 = st.columns([2, 1])
            with b1:
                pick = st.selectbox(
                    "Where should this go?", options,
                    key=f"pick_{finding.item}", label_visibility="collapsed",
                )
                target = pick
                if pick == "‹ pick from full list ›":
                    target = st.selectbox(
                        "Pick from roster", roster.items, key=f"full_{finding.item}"
                    )
            with b2:
                valid = target not in ("— choose —", "‹ pick from full list ›")
                if st.button("Save as alias", key=f"al_{finding.item}",
                             use_container_width=True):
                    if valid:
                        save_alias(finding.item, target)
                        st.rerun()
                if st.button("Fold into item", key=f"fo_{finding.item}",
                             use_container_width=True):
                    if valid:
                        add_fold(finding.item, target)
                        st.rerun()
                if st.button("Exclude (drop)", key=f"dr_{finding.item}",
                             use_container_width=True):
                    add_drop(finding.item)
                    st.rerun()
            st.caption(
                "**Alias** — same item, spelled differently.  "
                "**Fold** — different item, but its stock rolls into another.  "
                "**Drop** — not part of the report at all."
            )

for finding in [f for f in errors if f.category != "unmatched"]:
    st.error(finding.message)

if warnings:
    with st.expander(f"{len(warnings)} warning(s)"):
        for finding in warnings:
            st.warning(finding.message)

if infos:
    with st.expander(f"{len(infos)} rule(s) applied automatically"):
        for finding in infos:
            st.info(
                f"{finding.message} (Intransit {fmt(finding.intransit)}, "
                f"Onhand {fmt(finding.onhand)})"
            )

st.subheader("Totals")
summary = pd.DataFrame([
    {"Source": "TW",
     "Intransit": result.stats.get("tw_intransit", 0),
     "Onhand": result.stats.get("tw_onhand", 0)},
    {"Source": "AGT",
     "Intransit": sum(v["AGT_intransit"] for v in result.values.values()),
     "Onhand": result.stats.get("agt_onhand", 0)},
])
summary.loc[len(summary)] = {
    "Source": "TW + AGT",
    "Intransit": summary["Intransit"].sum(),
    "Onhand": summary["Onhand"].sum(),
}
st.dataframe(
    summary.style.format({"Intransit": "{:,.0f}", "Onhand": "{:,.0f}"}),
    use_container_width=True, hide_index=True,
)
st.caption("DCR and QS are excluded — they stay manual and are never touched.")

# ---- Delta vs previous run
prev_path = OUTPUT_DIR / "_previous_totals.json"
current_totals = {
    "tw_intransit": result.stats.get("tw_intransit", 0),
    "tw_onhand": result.stats.get("tw_onhand", 0),
    "agt_intransit": sum(v["AGT_intransit"] for v in result.values.values()),
    "agt_onhand": result.stats.get("agt_onhand", 0),
}
if prev_path.exists():
    try:
        with open(prev_path, "r", encoding="utf-8") as fh:
            previous = json.load(fh)
        deltas = []
        for key, now in current_totals.items():
            before = previous.get("totals", {}).get(key, 0)
            change = now - before
            pct = (change / before * 100) if before else 0
            deltas.append({
                "Metric": key.replace("_", " ").title(),
                "Previous": before, "Now": now,
                "Change": change, "%": pct,
            })
        with st.expander(
            f"Compare with the previous report ({previous.get('date', '?')})"
        ):
            st.dataframe(
                pd.DataFrame(deltas).style.format({
                    "Previous": "{:,.0f}", "Now": "{:,.0f}",
                    "Change": "{:+,.0f}", "%": "{:+.1f}%",
                }),
                use_container_width=True, hide_index=True,
            )
            big = [d for d in deltas if abs(d["%"]) > 50 and d["Previous"]]
            if big:
                st.warning(
                    "Large swing in: " + ", ".join(d["Metric"] for d in big)
                    + ". Worth a second look before you submit."
                )
    except (json.JSONDecodeError, OSError):
        pass

# ===================================================================
# STEP 5 — Generate
# ===================================================================
st.header("5 · Generate")
st.caption(
    "Only **H, I** (TW) and **N, O** (AGT) are written. G, J, M (DCR) and "
    "K, L (QS) keep whatever is in the template."
)

if errors:
    st.warning(
        f"{len(errors)} error(s) still open. You can still generate, but some "
        "stock will not make it into the report."
    )

g1, g2 = st.columns(2)
with g1:
    blank_zeros = st.checkbox(
        "Leave blanks instead of 0", value=False,
        help="Off matches how your report looks today.",
    )
with g2:
    reset_manual = st.checkbox(
        "Zero out DCR and QS first", value=False,
        help=(
            "Turn on when the template is last week's report, so old DCR and "
            "QS values are not carried over. SRP is always kept."
        ),
    )

if st.button("📄 Generate report", type="primary", use_container_width=True):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUTPUT_DIR / f"ND_SOH_{report_date:%B_%d_%Y}.xlsx"

    try:
        source_template = template_path
        if reset_manual:
            staged = OUTPUT_DIR / "_staged_template.xlsx"
            clear_report(
                template_path, str(staged), roster, groups=("QS", "DCR"),
                blank_instead_of_zero=blank_zeros,
            )
            source_template = str(staged)

        info = write_report(
            source_template, str(out_path), roster, result.values,
            report_date=datetime.combine(report_date, datetime.min.time()),
            blank_zeros=blank_zeros, groups=("TW", "AGT"),
        )
        with open(prev_path, "w", encoding="utf-8") as fh:
            json.dump({"date": str(report_date), "totals": current_totals}, fh, indent=2)

        st.success(
            f"Done — {info['items_written']} items, {info['cells_written']} cells "
            f"written to columns {', '.join(info['columns_written'])}."
        )
        if reset_manual:
            st.info(
                f"**Zeroed:** columns {', '.join(info['columns_untouched'])} "
                "(DCR and QS) — clean and ready for you to fill. SRP intact."
            )
        else:
            st.info(
                f"**Untouched:** columns {', '.join(info['columns_untouched'])} "
                "(DCR and QS) — still holding whatever the template had."
            )
        with open(out_path, "rb") as fh:
            st.download_button(
                "⬇ Download report", fh.read(), file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption(f"Also saved to `{out_path}`")
    except (OSError, ValueError, KeyError) as exc:
        st.error(f"Could not generate the report: {exc}")

# ===================================================================
# Sidebar — what the app has learned
# ===================================================================
with st.sidebar:
    st.header("Aliases and rules")
    st.caption(
        "What the app has learned. Every answer you give here means one less "
        "question next time."
    )

    aliases = load_aliases()
    st.subheader(f"Aliases ({len(aliases)})")
    if aliases:
        for raw, target in aliases.items():
            s1, s2 = st.columns([4, 1])
            s1.text(f"{raw} → {target}")
            if s2.button("✕", key=f"da_{raw}"):
                delete_alias(raw)
                st.rerun()
    else:
        st.caption("None yet.")

    rules = load_rules()

    st.subheader(f"Dropped ({len(rules['drop'])})")
    if rules["drop"]:
        for item in rules["drop"]:
            s1, s2 = st.columns([4, 1])
            s1.text(item)
            if s2.button("✕", key=f"dd_{item}"):
                clear_rule(item)
                st.rerun()
    else:
        st.caption("None yet.")

    st.subheader(f"Folded ({len(rules['fold'])})")
    if rules["fold"]:
        for src, dst in rules["fold"].items():
            s1, s2 = st.columns([4, 1])
            s1.text(f"{src} → {dst}")
            if s2.button("✕", key=f"df_{src}"):
                clear_rule(src)
                st.rerun()
    else:
        st.caption("None yet.")

    st.divider()
    st.caption(
        "DCR and QS are manual and outside the app's scope. "
        "Columns G, J, M, K and L are never written to."
    )
