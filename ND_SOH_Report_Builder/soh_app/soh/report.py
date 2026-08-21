"""
Report reader/writer.

Dalawang trabaho:
  1. Basahin ang roster (listahan ng item) mula sa template.
  2. Isulat ang values pabalik sa template nang hindi nasisira
     ang formatting, grouping, at merged cells.

Na-verify ang round-trip laban sa Aug 20 report:
  109 grouped rows, 5 grouped columns, 6 merged cells,
  5 conditional formatting rules, 3,066 styled cells - lahat buo.

MAHALAGA: Ang column O (AGT Actual Onhand) ay may external link sa
template (=IFERROR(SUMIFS([1]Sheet1!...))). Sinasadya nating palitan
ito ng value. Ang external link ay nasisira kapag nag-save ang openpyxl,
at mas ligtas na ang value kaysa sa link na naka-turo sa file na wala
na sa Downloads folder.
"""

import shutil
from dataclasses import dataclass

from openpyxl import load_workbook

from .normalize import clean_text, item_key

# Mapping ng field papunta sa column sa ND SOH layout.
# D=Item, G..R = ang tatlong distributor + total.
COLUMN_FIELDS = {
    "TW_dcr": "G",
    "TW_intransit": "H",
    "TW_onhand": "I",
    "QS_dcr": "J",
    "QS_intransit": "K",
    "QS_onhand": "L",
    "AGT_dcr": "M",
    "AGT_intransit": "N",
    "AGT_onhand": "O",
}

# Nakagrupo ayon sa kung sino ang may hawak ng bawat column.
#
# Ang app ay nagsusulat lang sa TW at AGT. Ang DCR (G/J/M) at QS (K/L)
# ay manual - hindi sila hinahawakan, kaya hindi mabubura ang trabaho
# mo kapag ikaw ang gumawa ng pivot at SUMIFS doon.
FIELD_GROUPS = {
    "TW": ["TW_intransit", "TW_onhand"],
    "AGT": ["AGT_intransit", "AGT_onhand"],
    "QS": ["QS_intransit", "QS_onhand"],
    "DCR": ["TW_dcr", "QS_dcr", "AGT_dcr"],
}

# Ito ang default: TW at AGT lang.
DEFAULT_GROUPS = ("TW", "AGT")


def fields_for(groups) -> list:
    """Ibalik ang listahan ng field name para sa mga napiling grupo."""
    out = []
    for group in groups:
        out.extend(FIELD_GROUPS.get(group, []))
    return out

_ITEM_HEADERS = {"item", "item code"}
_SKIP_SERIES = ("GRAND TOTAL", "DCR INVENTORY", "TOTAL")


@dataclass
class Roster:
    """Ang listahan ng item sa report, kasama ang row number ng bawat isa."""
    items: list           # list[str] - canonical item codes, sunod sa template
    row_of: dict          # {item: excel_row}
    header_row: int
    item_col: int
    sheet_name: str
    skipped_rows: list    # (row, series) ng mga total row na nilaktawan


def read_roster(path: str, sheet_name: str = "SOH") -> Roster:
    """
    Basahin ang item roster mula sa template.

    Ang template ang laging masusunod: hindi tayo magdadagdag o
    magbabawas ng row. Kung may item sa raw na wala dito, ita-tag
    siyang unmatched sa validation - hindi basta idadagdag.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
        actual_sheet = ws.title

        header_row = None
        item_col = None
        series_col = None

        for row_idx, row in enumerate(ws.iter_rows(max_row=20, values_only=True), start=1):
            for col_idx, cell in enumerate(row):
                label = clean_text(cell).lower()
                if label in _ITEM_HEADERS:
                    header_row = row_idx
                    item_col = col_idx
                if label == "series":
                    series_col = col_idx
            if header_row:
                break

        if header_row is None:
            raise ValueError(
                f"Could not find an 'Item' header in sheet '{actual_sheet}'."
            )

        items = []
        row_of = {}
        skipped = []

        for offset, values in enumerate(
            ws.iter_rows(min_row=header_row + 1, values_only=True)
        ):
            excel_row = header_row + 1 + offset

            series = ""
            if series_col is not None and series_col < len(values):
                series = clean_text(values[series_col]).upper()

            if any(series.startswith(s) for s in _SKIP_SERIES):
                skipped.append((excel_row, series))
                continue

            if item_col >= len(values):
                continue
            item = item_key(values[item_col])
            if not item:
                continue

            # Kung may duplicate, ang unang row ang panalo.
            if item not in row_of:
                items.append(item)
                row_of[item] = excel_row

        return Roster(
            items=items,
            row_of=row_of,
            header_row=header_row,
            item_col=item_col,
            sheet_name=actual_sheet,
            skipped_rows=skipped,
        )
    finally:
        wb.close()


def clear_report(
    template_path: str,
    output_path: str,
    roster: Roster,
    groups=("TW", "QS", "AGT", "DCR"),
    blank_instead_of_zero: bool = False,
) -> dict:
    """
    Refresher: i-reset ang data columns papuntang 0.

    ITINATAGO (hindi ginagalaw):
      B  Series
      C  Launch Date
      D  Item
      E  Market Name
      F  SRP  <- ang price list
      P, Q, R  Total formulas (=G+J+M) - automatic silang magiging 0
      Grand Total, DCR Inventory, at Diff. qty. rows - puro formula sila
         (=SUM(G5:G140), =+G141, =G141-I141), kaya sila na ang bahala

    NILILINIS:
      G, H, I  TW
      J, K, L  QS
      M, N, O  AGT

    Ang column O ay may sirang external link sa lumang template
    (=IFERROR(SUMIFS([1]Sheet1!...),0)). Kapag nilinis, napapalitan
    ito ng plain 0 - kaya nawawala rin ang link na madalas mabigo
    nang tahimik.

    Data rows lang ang hinahawakan - ang mga nasa roster. Ang total
    rows ay nilalaktawan dahil formula sila at magsasariling mag-update.
    """
    shutil.copyfile(template_path, output_path)

    active_fields = fields_for(groups)

    wb = load_workbook(output_path)  # HINDI data_only - buhay ang formulas
    try:
        ws = wb[roster.sheet_name]

        cleared = 0
        for row in roster.row_of.values():
            for field in active_fields:
                cell = ws[f"{COLUMN_FIELDS[field]}{row}"]
                cell.value = None if blank_instead_of_zero else 0
                cleared += 1

        wb.save(output_path)

        return {
            "cells_cleared": cleared,
            "rows": len(roster.row_of),
            "columns_cleared": sorted(COLUMN_FIELDS[f] for f in active_fields),
            "output_path": output_path,
        }
    finally:
        wb.close()


def write_report(
    template_path: str,
    output_path: str,
    roster: Roster,
    values: dict,
    report_date=None,
    blank_zeros: bool = False,
    groups=DEFAULT_GROUPS,
) -> dict:
    """
    Kopyahin ang template, tapos isulat ang values.

    groups: alin sa "TW", "AGT", "QS", "DCR" ang isusulat.
            Ang default ay ("TW", "AGT") lang - ang DCR at QS ay
            manual, kaya hindi sila hinahawakan. Mahalaga ito:
            kung isusulat natin sila, mabubura ang pivot at SUMIFS
            na inilagay mo mismo sa template.

    Hindi hinahawakan ang formatting - values lang ang isinusulat.
    Ang columns P/Q/R (Total) ay may formula na sa template (=G+J+M),
    kaya hindi natin sila ginagalaw - sila na ang magko-compute.

    blank_zeros: kung True, iiwang blank ang mga zero imbes na maglagay
                 ng 0. Default ay False (0 ang ilalagay).
    """
    shutil.copyfile(template_path, output_path)

    active_fields = fields_for(groups)
    untouched = [
        COLUMN_FIELDS[f] for f in COLUMN_FIELDS if f not in active_fields
    ]

    wb = load_workbook(output_path)  # HINDI data_only - para buhay ang formulas
    try:
        ws = wb[roster.sheet_name]

        written = 0
        for item, row in roster.row_of.items():
            row_values = values.get(item)
            if row_values is None:
                continue
            for field in active_fields:
                column = COLUMN_FIELDS[field]
                amount = row_values.get(field, 0) or 0
                cell = ws[f"{column}{row}"]
                if blank_zeros and not amount:
                    cell.value = None
                else:
                    cell.value = amount
                written += 1

        # I-update ang report date kung mahahanap ang cell nito.
        date_cell = None
        if report_date is not None:
            for r in range(1, roster.header_row + 1):
                for c in range(1, 12):
                    text = clean_text(ws.cell(row=r, column=c).value).lower()
                    if "stocks on hand as of" in text:
                        # Ang petsa ay nasa cell sa kanan ng label.
                        for offset in range(1, 5):
                            target = ws.cell(row=r, column=c + offset)
                            if not isinstance(target, type(ws["A1"])):
                                continue
                            try:
                                target.value = report_date
                                date_cell = target.coordinate
                            except AttributeError:
                                continue  # merged cell - laktawan
                            break
                        break
                if date_cell:
                    break

        wb.save(output_path)

        return {
            "cells_written": written,
            "items_written": len([i for i in roster.row_of if i in values]),
            "date_cell": date_cell,
            "output_path": output_path,
            "columns_written": sorted(COLUMN_FIELDS[f] for f in active_fields),
            "columns_untouched": sorted(untouched),
        }
    finally:
        wb.close()
