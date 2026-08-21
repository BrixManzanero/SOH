# ND SOH Report Builder

A Streamlit app that replaces the manual Clean → Pivot → SUMIFS work for the
**TW and AGT** portions of the National Distributors stocks-on-hand report.

It never modifies your template. It copies it, fills in the cells it owns, and
saves the result as a new file. Grouping, formatting and merged cells all
survive intact.

---

## Scope

| Column | Contents | Who fills it |
|---|---|---|
| H, I | TW Intransit / Onhand | **the app** |
| N, O | AGT Intransit / Onhand | **the app** |
| G, J, M | DCR Inventory | you (pivot + lookup) |
| K, L | QS Intransit / Onhand | you (from the photo) |
| P, Q, R | Totals | formulas already in the template |

**The app never writes to G, J, M, K or L.** They keep whatever the template
had, so your manual work there cannot be overwritten.

---

## Setup

### Run it locally (recommended)

Easiest: double-click `run_local.bat`.

From a terminal:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost
```

Opens at `http://localhost:8501`. Python 3.9 or newer.

If `python` is not recognised, try `py` instead. If neither works, install
Python from python.org and tick **Add python.exe to PATH** during setup.

**Why `--server.address localhost` matters.** By default Streamlit also serves
on your local network and prints a second "Network URL". On an office WiFi,
anyone on the same network could open that address and read your stock data.
Binding to localhost makes the app reachable from this machine only. This was
verified: the LAN address refuses the connection while localhost still works.
`run_local.bat` already includes the flag.

**A note on OneDrive.** If the app folder lives inside OneDrive, the `output/`
folder and `config/*.json` are synced to your OneDrive account, and OneDrive
can lock files mid-write. If you ever hit a "permission denied" while saving,
move the folder to `C:\soh_app`.

### Deploy to Streamlit Community Cloud (optional)

1. Create a GitHub repo (private is fine).
2. Upload the contents of this folder.
3. Go to share.streamlit.io → **New app** → select the repo.
4. Main file path: `app.py`
5. Deploy. It takes about a minute.

In the app, choose **Upload files** mode. The files are small (TW around 7 MB,
AGT around 18 KB, template around 32 KB), so the free tier handles them
comfortably.

Two things to know first. On the free tier, apps from a private repo are
private, but you only get **one** private app; everything else has to live in
a public repo. And the app still handles internal stock figures, so use the
viewer allowlist if you share it.

---

## The five steps

| # | Step | What happens |
|---|---|---|
| 1 | Select files | Searchable browser over any folder: TW raw, AGT raw, template |
| 2 | TW cleaning | LDU merge, whitespace fixes, Grand Total check |
| 3 | AGT preview | Lookup against the roster, flags anything with no match |
| 4 | Validation | Unit accounting, exclusions, unmatched items, deltas |
| 5 | Generate | Writes H/I and N/O only, then offers the download |

Afterwards, open the output in Excel and do the DCR pivot and the QS entry as
usual.

### File selection

The picker is a search box, not a dropdown. Type any part of a filename and
the list filters as you type; multiple words all have to match. Every result
shows its size and modified date, newest first. Shortcut buttons jump to
Downloads, Desktop, Documents and OneDrive Desktop. Tick **Include
subfolders** to search deeper, or paste a full path directly.

---

## What the app fixes automatically

Four classes of problem were found in the raw files.

**Class A — whitespace.** Trailing spaces (`'T1002 128+4 '`), double spaces
(`'Mega Pad Pro  KB'`), and non-breaking spaces, CHAR(160). The last one is
the dangerous one: it is invisible in Excel and silently breaks any lookup.

**Class B — LDU prefix.** 21 rows in a typical TW file. Display units belong
to their base item, so `LDU CN7c 256+8` (50 units) merges into `CN7c 256+8`,
turning 11,296 into 11,346.

**Class C — swapped columns.** In the TW raw, `BUDS 4 AIR` sits in the Item
column while `BD04 Air` sits in Market Name — the reverse of the report. No
find-and-replace can fix this, because nothing is misspelled; the values are
simply in the wrong columns.

**Class D — a different item code.** The TW raw writes `TSP-W03A` where some
weeks' rosters say `TSP-WP03A`.

Classes A and B are handled automatically: they are mechanical and safe.
Classes C and D are never guessed. The app asks once, then remembers.

---

## Aliases, equivalents and rules

Stored in `config/`. This is the app's memory.

### aliases.json

**`aliases`** — one-way, from a name in the raw file to a report code. Use
this only when the raw name could never itself be a roster code:

```json
{"BUDS 4 AIR": "BD04 Air", "T1002": "T1002 256+4"}
```

**`equivalents`** — codes that mean the same product, with no fixed direction:

```json
{"equivalents": [["TSP-W03A", "TSP-WP03A"]]}
```

This exists because the roster itself gets renamed. Across six weeks:

| | code used |
|---|---|
| roster Jul 27 – Aug 10 | `TSP-W03A` |
| roster Aug 20 | `TSP-WP03A` |
| TW raw writes | `TSP-W03A` |
| AGT raw writes | `TSP-WP03A` |

A one-way alias breaks one week or the other. An equivalence group only says
"these are the same thing", and the app resolves it against whichever form the
roster in front of it actually uses.

**On the `T1002` alias.** For five consecutive weeks `T1002 128+4` held zero
while `T1002 256+4` held everything, so a bare `T1002` in the raw belongs to
`256+4`. Worth watching, though: on Aug 20 the split changed to 3,000 / 5,480,
and that week TW wrote the two variants as separate rows. The bare form has so
far only appeared when the stock was entirely `256+4`. If a bare `T1002` ever
shows up while `128+4` also has stock, stop and ask TW how it splits.

### rules.json

```json
{
  "drop": ["Mega Pad Pro KB", "MegaPad 2"],
  "fold": {},
  "ask":  []
}
```

- **drop** — never included. You delete these from the TW raw by hand.
- **fold** — stock always rolls into another item code.
- **ask** — always flagged for a decision, never resolved automatically.

Dropped quantities are **never hidden**. The Validation step lists every item
removed by a rule, with its units, on every single run. A rule saves you a
click; it cannot quietly delete stock.

**Why the two keyboards are dropped.** `MegaPad 2` is the keyboard for the
MegaPad 2 tablet (T1103), not the tablet, and neither it nor `Mega Pad Pro KB`
has ever been a roster row in six weeks. The report has exactly one keyboard
row, `T1101` / "Keyboard Megapad 11", and it is fully accounted for (201 units
every week). Your own manual process reaches the same result: the lookup runs
from the roster into the pivot, so an item with no roster row is never looked
up at all.

You never edit these files by hand. When an item is unmatched, the app offers
four buttons — Alias, Fold, Drop, Ask me every time — and saves your answer.

---

## The refresher

If your template is last week's report, it still holds last week's DCR and QS
values. Two ways to clear them:

- **Step 1** — open the Refresher panel, choose which columns, download a
  blank template.
- **Step 5** — tick *"Zero out DCR and QS first"*. G/J/M and K/L are cleared
  before TW and AGT are written.

Always kept: Series, Launch Date, Item, Market Name, **SRP (the price list)**,
the Total formulas, and the Grand Total / DCR Inventory / Diff. qty. rows.
Those last ones are pure formulas (`=SUM(G5:G140)`, `=+G141`, `=G141-I141`),
so they fall to zero by themselves.

---

## If Excel offers to repair the file

The old template reaches the AGT file through an external link:

```
=IFERROR(SUMIFS([1]Sheet1!$C:$C, [1]Sheet1!$A:$A, D5), 0)
```

It points at the `Sheet1` pivot you build inside the AGT file, and it works
while that pivot exists. The risk is that `IFERROR` turns any failure into a
plain **0** — no `#REF!`, no warning. A missing pivot looks exactly like
"no stock".

The app removes the need for it: it reads the raw AGT sheet, aggregates, and
writes a value into column O. Verified against your own Aug 17 pivot — 115
items, 800 Intransit, 105,450 Onhand, zero differences.

Both writers also strip external links before saving and set
`fullCalcOnLoad`, so Excel recalculates the totals on open. If a repair prompt
still appears, click **Yes** — the data is fine — and send the file over so it
can be traced.

---

## Files

```
app.py                  Streamlit UI, five steps
run_local.bat           Windows launcher, binds to localhost only
soh/normalize.py        Text cleaning: whitespace, CHAR(160), LDU
soh/stock_files.py      Parser for the TW and AGT raw files
soh/reconcile.py        Aliases, equivalents, rules, findings
soh/report.py           Roster reader, scoped writer, refresher
soh/config.py           Alias and rule storage
config/aliases.json     What the app has learned
config/rules.json       Drop / fold / ask rules
output/                 Where finished reports appear
```

---

## Verification

Two independent weeks were checked end to end against completed reports:

| | TW Intransit | TW Onhand | AGT Intransit | AGT Onhand |
|---|---|---|---|---|
| **Aug 3** | 59,190 — match | 237,352 — match | no file | no file |
| **Aug 10** | 111,000 — match | 224,933 — match | 37,610 — match | 82,956 — match |

Both weeks ran with zero unmatched items and zero questions.

Structural checks on the output: 109 grouped rows intact, merged cells intact,
conditional formatting intact, 3,066 styled cells with no differences, Total
formulas alive, DCR and QS columns untouched, no external links.

**What is not yet verified.** Only one week has been checked for AGT. Four of
the six weeks have reports but no raw files, so only roster-level facts could
be confirmed there. And no Excel was available during development, so the
repair-prompt fix was validated structurally rather than by opening the file
in Excel.

For the first real report, run the app alongside your manual process and
compare TW Intransit, TW Onhand and AGT Onhand. That comparison is worth more
than any of the tests above.
