# ND SOH Report Builder

Streamlit app na pumapalit sa manual na Clean → Pivot → SUMIFS na proseso
para sa **TW at AGT** na bahagi ng National Distributors stocks-on-hand report.

Hindi nito binabago ang template. Kinokopya niya ito, pinupunan ang mga cell,
at ise-save bilang bagong file. Buo ang grouping, formatting, at merged cells.

## Sakop

| Column | Laman | Sino ang gumagawa |
|---|---|---|
| H, I | TW Intransit / Onhand | **app** |
| N, O | AGT Intransit / Onhand | **app** |
| G, J, M | DCR Inventory | ikaw (pivot + SUMIFS) |
| K, L | QS Intransit / Onhand | ikaw (galing sa picture) |
| P, Q, R | Totals | formula sa template |

**Hindi hinahawakan ng app ang G, J, M, K, at L.** Mananatili silang kung ano
ang nasa template, kaya hindi mabubura ang manual mong trabaho doon.

---

## Dalawang paraan ng paggamit

The app has a **mode switch** in step 1:

- **Browse a folder on this PC** — searchable file browser, no upload. Faster.
- **Upload files** — needed when deployed to Streamlit Community Cloud.

The browser lists every Excel file in the chosen folder (newest first) with a
search box that filters as you type. Shortcut buttons jump to Downloads,
Desktop, Documents and OneDrive Desktop. Tick "Include subfolders" to search
deeper, or paste a full path directly.

The interface is in English.

### A. Local (recommended)

**Easiest:** double-click `run_local.bat`.

Or from a terminal:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost
```

Opens at `http://localhost:8501`.

**Why `--server.address localhost` matters.** By default Streamlit also
serves on your local network and prints a second "Network URL". On an office
WiFi, anyone on the same network could open that address and see your stock
data. Binding to localhost makes the app reachable from this PC only —
verified: the LAN address refuses the connection while localhost still works.
`run_local.bat` already includes the flag.

**Note on OneDrive.** If the app folder sits inside OneDrive, the `output/`
folder and `config/*.json` get synced to your OneDrive account. Usually fine,
but OneDrive can lock files while syncing. If you hit a "permission denied"
on save, move the folder to `C:\soh_app`.

### B. Streamlit Community Cloud (via GitHub)

1. Gumawa ng bagong GitHub repo (puwedeng private).
2. I-upload lahat ng laman ng folder na ito.
3. Pumunta sa share.streamlit.io → **New app** → piliin ang repo.
4. Main file path: `app.py`
5. Deploy. Aabutin ng ~1 minuto.

Pagkatapos, piliin ang **⬆ I-upload** mode sa app at i-upload ang tatlong file.

Ang mga file ay maliliit (TW ~7MB, AGT ~18KB, template ~32KB), kaya kaya
ito ng free tier. Walang IMEI-level na data na pumapasok dito.

Kung ide-deploy mo, aggregate stock levels pa rin ang laman — gamitin ang
private app setting kung kailangan.

Dahil TW at AGT lang ang sakop, maliliit na lang ang files (TW ~7MB, AGT ~18KB) —
kaya puwede rin itong i-deploy sa Streamlit Community Cloud kung gusto mo.
Wala nang IMEI-level na data na papasok dito.

Kung ide-deploy mo, tandaan na aggregate stock levels pa rin ang laman nito.
Gamitin ang private app setting kung kailangan.

---

## Ang anim na hakbang

| # | Hakbang | Ano ang nangyayari |
|---|---|---|
| 1 | Pumili ng files | Hahanapin sa Downloads folder ang TW raw, AGT raw, at template |
| 2 | TW cleaning preview | LDU merge, whitespace fix, at Grand Total check |
| 3 | AGT preview | Lookup laban sa roster, may flag kung may hindi tumugma |
| 4 | Validation | Totals, unmatched items, at delta kumpara sa nakaraan |
| 5 | Generate | Isusulat ang H/I at N/O lang, tapos i-download |

Pagkatapos nito, buksan mo ang output sa Excel at gawin ang DCR pivot at
SUMIFS, saka ang QS entry — gaya ng nakasanayan.

---

## Refresher — pag-zero ng data

Kung last week's report ang gagamitin mong template, may lumang DCR at QS
values pa rin doon. Dalawang paraan para linisin:

**A. Sa Hakbang 1** — buksan ang `Refresher` expander, piliin kung aling
columns, tapos i-download ang blankong template.

**B. Sa Hakbang 5** — i-check ang *"I-zero muna ang DCR at QS"*. Awtomatikong
lilinisin ang G/J/M at K/L bago isulat ang TW at AGT.

Itinatago (hindi ginagalaw):

| Column | Laman |
|---|---|
| B | Series |
| C | Launch Date |
| D | Item |
| E | Market Name |
| **F** | **SRP — ang price list** |
| P, Q, R | Total formulas (`=G+J+M`) |
| Grand Total, DCR Inventory, Diff. qty. rows | `=SUM(G5:G140)`, `=+G141`, `=G141-I141` |

Ang total rows ay puro formula, kaya sila na ang bahalang mag-update
papuntang 0. Hindi sila hinahawakan.

Bonus: kapag na-clear ang column O, napapalitan ng plain `0` ang sirang
external link doon — kaya wala nang formula na puwedeng tahimik na mabigo.

---

## Ano ang inaayos nito nang automatic

Apat na klase ng sira ang nakita sa Aug 20 data:

**Klase A — whitespace.** Trailing space (`'T1002 128+4 '`), double space
(`'Mega Pad Pro  KB'`), at non-breaking space CHAR(160). Nasa DCR customer
names din ang NBSP (`BOBBITECH\xa0ENTERPRISES\xa0INC`).

**Klase B — LDU prefix.** 21 rows sa TW raw. Isinasama sa base item.
Halimbawa: `LDU CN7c 256+8` (50 units) → `CN7c 256+8`, kaya 11,296 + 50 = 11,346.

**Klase C — baliktad na column.** Sa TW raw, `BUDS 4 AIR` ang nasa Item column
at `BD04 Air` ang nasa Market Name — kabaligtaran ng report. Hindi ito kayang
ayusin ng find-and-replace dahil tama naman ang spelling; mali lang ang column.

**Klase D — maling item code.** `TSP-W03A` sa TW raw, `TSP-WP03A` sa report.
Kulang ng letrang P.

Ang Klase A at B ay **automatic** — mekanikal at ligtas.
Ang Klase C at D ay **hindi hinuhulaan** — itatanong sa'yo, tapos matatandaan.

---

## Ang alias map at rules

Nasa `config/` ang dalawang file na ito. Sila ang "memory" ng app.

**`aliases.json`** — kapag iba ang pagkakasulat ng parehong item:

```json
{"aliases": {"BUDS 4 AIR": "BD04 Air", "TSP-W03A": "TSP-WP03A"}}
```

**`rules.json`** — dalawang klase ng manual na desisyon:

```json
{
  "drop": ["Mega Pad Pro KB"],
  "fold": {"MegaPad 2": "T1103 256+8"}
}
```

- **drop** — hindi isinasama sa report (dating manual mong binubura sa raw)
- **fold** — isinasama ang stock sa ibang item code

Hindi mo kailangang i-edit ang mga file na ito nang manu-mano. Kapag may
unmatched item, itatanong ng app at ise-save niya ang sagot mo. Sa susunod
na report, tahimik na niyang tama-tama ito.

---


## Tungkol sa column O (AGT Actual Onhand)

Sa lumang paraan, ang column O ay may external link:

```
=IFERROR(SUMIFS([1]Sheet1!$C:$C, [1]Sheet1!$A:$A, D5), 0)
```

Naghahanap ito ng `Sheet1` sa AGT file — ang PivotTable output na
ginagawa mo (Row Labels / Sum of Intransit / Sum of Onhand). Gumagana
ito nang tama kapag nagawa mo na ang pivot.

Ang panganib: kapag hindi pa nagagawa ang pivot, hindi mahahanap ng
SUMIFS ang `Sheet1`. Sinasalo ito ng `IFERROR` at nagbabalik ng **0** —
walang `#REF!`, walang babala. Mukhang "walang stock" imbes na
"wala pang pivot".

**Hindi na kailangan ng pivot sa AGT.** Binabasa ng app ang raw SOH
sheet at siya na ang nag-a-aggregate, tapos isinusulat ang aktwal na
value sa column O. Na-verify laban sa Aug 17 pivot mo: 115 items,
Intransit 800, Onhand 105,450 — **zero differences**.

Kung may `Sheet1` pivot ang AGT file, nilalaktawan ito ng app at ang
raw SOH sheet ang binabasa.

---


## Mga file

```
app.py                  Streamlit UI (anim na hakbang)
soh/normalize.py        Text cleaning — NBSP, spaces, LDU
soh/stock_files.py      Parser para sa TW at AGT raw
soh/reconcile.py        Pinagsasama ang lahat, gumagawa ng findings
soh/report.py           Roster reader at template writer
soh/config.py           Aliases, rules, distributor map
config/aliases.json     Natutunan ng app
config/rules.json       Drop at fold rules
output/                 Dito lumalabas ang tapos na report
```

---

## Na-verify

Ang engine ay sinubukan laban sa aktwal na Aug 20 report:

| Column | Resulta |
|---|---|
| TW Intransit | 113,480 — **eksaktong tugma** |
| TW Onhand | 210,929 — **eksaktong tugma** |
| AGT Intransit | eksaktong tugma |
| DCR at QS columns | **0 cells na nagalaw** |
| Grouping (109 rows, 5 cols) | buo |
| Merged cells (6) | buo |
| Conditional formatting (5 rules) | buo |
| Cell styling (3,066 cells) | 0 differences |
| Total formulas (P/Q/R) | buhay pa rin |
