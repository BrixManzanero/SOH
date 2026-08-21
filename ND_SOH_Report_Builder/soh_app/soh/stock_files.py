"""
Parser for the TW and AGT raw stock files.

Both files share the same layout (verified against the Aug 3, Aug 10
and Aug 20 files):

    Excel row 3   = header:  Series | Item | Market Name | Model | Intransit | Onhand
    Excel row 4+  = data
    a "Grand Total:" row at the bottom, which must be excluded

Column positions are not hard-coded. The header row is located by
name, so the parser survives a column being added or moved.
"""

from collections import defaultdict
from dataclasses import dataclass, field

from openpyxl import load_workbook

from .normalize import clean_text, item_key, is_ldu, describe_issues

# Accepted spellings for each column, lower-cased.
_HEADER_ALIASES = {
    "series": {"series"},
    "item": {"item", "item code", "itemcode"},
    "market": {"market name", "market", "marketname"},
    "model": {"model"},
    "intransit": {"intransit", "in transit", "in-transit"},
    "onhand": {"onhand", "on hand", "on-hand", "actual onhand"},
}


@dataclass
class StockRow:
    """One row from the raw file, before merging."""
    excel_row: int
    series: str
    raw_item: str
    item: str          # canonical key (cleaned, LDU stripped)
    market: str
    intransit: float
    onhand: float
    was_ldu: bool
    issues: list = field(default_factory=list)


@dataclass
class StockFile:
    """The result of parsing one raw stock file."""
    path: str
    sheet_name: str
    rows: list                     # list[StockRow] - before merging
    totals_row: tuple = None       # (intransit, onhand) from the Grand Total row
    header_row: int = None

    def merged(self) -> dict:
        """
        Merge every row under its canonical item code.

        This is where LDU rows fold into their base item and where
        duplicate rows for the same code are added together.

        Returns: {item_code: {"intransit": x, "onhand": y}}
        """
        out = defaultdict(lambda: {"intransit": 0.0, "onhand": 0.0})
        for row in self.rows:
            out[row.item]["intransit"] += row.intransit
            out[row.item]["onhand"] += row.onhand
        return dict(out)

    def item_totals(self) -> tuple:
        """Total Intransit and Onhand across item rows, excluding Grand Total."""
        return (
            sum(r.intransit for r in self.rows),
            sum(r.onhand for r in self.rows),
        )

    def totals_match(self) -> bool:
        """
        This is step 6 of the manual workflow: TW raw total must equal
        the pivot total.

        When the file carries a Grand Total row, the sum of the item
        rows is compared against it. A mismatch means a row is missing
        or has been counted twice, and the run should stop.
        """
        if self.totals_row is None:
            return True
        item_i, item_o = self.item_totals()
        gt_i, gt_o = self.totals_row
        return abs(item_i - gt_i) < 0.5 and abs(item_o - gt_o) < 0.5

    def rows_with_issues(self) -> list:
        return [r for r in self.rows if r.issues]


def _to_number(value) -> float:
    """
    Convert a cell value to a number.

    Handles the cases from step 4 of the manual workflow: blanks,
    dashes, and numbers stored as text.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)
    if text in ("", "-", "--", "N/A", "n/a"):
        return 0.0

    # Drop thousands separators and treat parentheses as a minus sign.
    text = text.replace(",", "").replace("(", "-").replace(")", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _find_header(ws, max_scan: int = 15):
    """
    Locate the header row and the column index of each field.

    Returns: (excel_row_number, {field: zero_based_col_index})
    """
    for row_idx, row in enumerate(ws.iter_rows(max_row=max_scan, values_only=True), start=1):
        labels = {}
        for col_idx, cell in enumerate(row):
            label = clean_text(cell).lower()
            if not label:
                continue
            for field_name, accepted in _HEADER_ALIASES.items():
                if label in accepted and field_name not in labels:
                    labels[field_name] = col_idx

        # An Item column plus at least one quantity column is enough.
        if "item" in labels and ("intransit" in labels or "onhand" in labels):
            return row_idx, labels

    raise ValueError(
        "Could not find the header row. Expected columns named "
        "'Item', 'Intransit' and 'Onhand' within the first 15 rows."
    )


def _pick_sheet(wb) -> str:
    """
    Pick the sheet that holds the real SOH data.

    The AGT file sometimes contains two sheets: the raw
    'SOH as of MM.DD.YYYY' and a pivot called 'Sheet1' (Row Labels /
    Sum of Intransit / Sum of Onhand). The pivot comes first, so
    taking wb.sheetnames[0] would read the wrong one.

    Instead the parser looks for a sheet with a genuine header row:
    Item, Intransit, Onhand. The pivot does not qualify because its
    first column is headed 'Row Labels' rather than 'Item'.
    """
    candidates = []
    for name in wb.sheetnames:
        try:
            _find_header(wb[name])
        except ValueError:
            continue
        # Prefer a sheet with "SOH" in its name when several qualify.
        score = 1 if "soh" in name.lower() else 0
        candidates.append((score, name))

    if not candidates:
        raise ValueError(
            "No sheet in this file has 'Item', 'Intransit' and 'Onhand' "
            "headers. If the file only contains a pivot, the raw SOH sheet "
            "is required."
        )

    candidates.sort(key=lambda x: -x[0])
    return candidates[0][1]


def parse_stock_file(path: str, sheet_name: str = None) -> StockFile:
    """
    Read a TW or AGT raw file.

    With no sheet_name the correct sheet is detected automatically,
    rather than defaulting to the first one, which is sometimes a
    pivot.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb[_pick_sheet(wb)]
        actual_sheet = ws.title

        header_row, cols = _find_header(ws)

        rows = []
        totals_row = None

        for offset, values in enumerate(
            ws.iter_rows(min_row=header_row + 1, values_only=True)
        ):
            excel_row = header_row + 1 + offset

            def get(field_name):
                idx = cols.get(field_name)
                if idx is None or idx >= len(values):
                    return None
                return values[idx]

            series = clean_text(get("series"))
            raw_item = get("item")
            item = item_key(raw_item)

            # Capture the Grand Total row: keep it out of the data,
            # but retain it for the totals check.
            joined = f"{series} {clean_text(raw_item)}".upper()
            if "GRAND TOTAL" in joined or joined.strip().startswith("TOTAL"):
                totals_row = (
                    _to_number(get("intransit")),
                    _to_number(get("onhand")),
                )
                continue

            if not item:
                continue

            intransit = _to_number(get("intransit"))
            onhand = _to_number(get("onhand"))

            rows.append(
                StockRow(
                    excel_row=excel_row,
                    series=series,
                    raw_item=clean_text(raw_item),
                    item=item,
                    market=clean_text(get("market")),
                    intransit=intransit,
                    onhand=onhand,
                    was_ldu=is_ldu(raw_item),
                    issues=describe_issues(raw_item),
                )
            )

        return StockFile(
            path=path,
            sheet_name=actual_sheet,
            rows=rows,
            totals_row=totals_row,
            header_row=header_row,
        )
    finally:
        wb.close()
