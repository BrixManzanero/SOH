"""
Reconciliation engine - dito pinagsasama ang lahat ng source.

Ang roster (listahan ng item) ay galing sa template mismo, hindi sa
raw files. Ibig sabihin, ang template ang laging masusunod - hindi
tayo magdadagdag ng row na wala doon.

Ang mahalagang output dito ay hindi lang ang numbers, kundi ang
findings: ano ang hindi tumugma, ano ang nawawala, ano ang dapat
mong tignan bago i-submit.
"""

from collections import defaultdict
from dataclasses import dataclass, field

from .normalize import match_key
from .config import load_aliases, load_rules


@dataclass
class Finding:
    """Isang bagay na dapat malaman ng user."""
    level: str        # "error" | "warning" | "info"
    category: str     # "unmatched" | "totals" | "cleaning" | "rule" | "dcr"
    item: str
    message: str
    intransit: float = 0.0
    onhand: float = 0.0
    suggestions: list = field(default_factory=list)


@dataclass
class ReconResult:
    values: dict          # {item_code: {"TW_intransit": x, ...}}
    findings: list
    stats: dict

    def errors(self):
        return [f for f in self.findings if f.level == "error"]

    def warnings(self):
        return [f for f in self.findings if f.level == "warning"]


def _apply_aliases_and_rules(merged: dict, roster_keys: dict):
    """
    I-apply ang alias map at manual rules sa merged raw data.

    Ibinabalik: (resolved, dropped, folded, unmatched)
    """
    aliases = load_aliases()
    rules = load_rules()
    drop_set = set(rules["drop"])
    fold_map = rules["fold"]

    resolved = defaultdict(lambda: {"intransit": 0.0, "onhand": 0.0})
    dropped = {}
    folded = {}
    unmatched = {}

    for item, qty in merged.items():
        # 1. Rule: drop
        if item in drop_set:
            if qty["intransit"] or qty["onhand"]:
                dropped[item] = qty
            continue

        # 2. Rule: fold papunta sa ibang item
        target = fold_map.get(item)
        if target:
            folded[item] = (target, qty)
            resolved[target]["intransit"] += qty["intransit"]
            resolved[target]["onhand"] += qty["onhand"]
            continue

        # 3. Alias: iba ang pangalan sa raw kaysa sa report
        target = aliases.get(item, item)

        # 4. Nasa roster ba?
        if target in roster_keys:
            resolved[target]["intransit"] += qty["intransit"]
            resolved[target]["onhand"] += qty["onhand"]
        else:
            if qty["intransit"] or qty["onhand"]:
                unmatched[item] = qty

    return dict(resolved), dropped, folded, unmatched


def _suggest_matches(item: str, roster: list, limit: int = 3) -> list:
    """
    Maghanap ng malapit na item code sa roster.

    Ginagamit ito para tulungan ang user kapag may unmatched item -
    hindi ito automatic na pinipili, mungkahi lang.
    """
    target = match_key(item)
    scored = []
    for candidate in roster:
        cand_key = match_key(candidate)
        if not cand_key or not target:
            continue
        if cand_key == target:
            score = 100
        elif cand_key.startswith(target) or target.startswith(cand_key):
            score = 80
        elif cand_key in target or target in cand_key:
            score = 60
        else:
            # Bilangin ang magkatulad na character prefix.
            common = 0
            for a, b in zip(cand_key, target):
                if a != b:
                    break
                common += 1
            score = common * 4 if common >= 3 else 0
        if score:
            scored.append((score, candidate))
    scored.sort(reverse=True)
    return [name for _, name in scored[:limit]]


def reconcile(roster, tw_file, agt_file, dcr_result, qs_manual=None) -> ReconResult:
    """
    Pagsamahin ang lahat ng source papunta sa isang values dict.

    roster      : list ng item code mula sa template (ito ang masusunod)
    tw_file     : StockFile para sa TW
    agt_file    : StockFile para sa AGT (puwedeng None)
    dcr_result  : DcrResult (puwedeng None)
    qs_manual   : {item: {"intransit": x, "onhand": y}} - manual entry
    """
    roster_keys = {item: True for item in roster}
    findings = []
    values = {
        item: {
            "TW_dcr": 0.0, "TW_intransit": 0.0, "TW_onhand": 0.0,
            "QS_dcr": 0.0, "QS_intransit": 0.0, "QS_onhand": 0.0,
            "AGT_dcr": 0.0, "AGT_intransit": 0.0, "AGT_onhand": 0.0,
        }
        for item in roster
    }

    stats = {}

    # ---------------------------------------------------------------- TW
    tw_merged = tw_file.merged()
    tw_resolved, tw_dropped, tw_folded, tw_unmatched = _apply_aliases_and_rules(
        tw_merged, roster_keys
    )
    for item, qty in tw_resolved.items():
        values[item]["TW_intransit"] = qty["intransit"]
        values[item]["TW_onhand"] = qty["onhand"]

    if not tw_file.totals_match():
        item_i, item_o = tw_file.item_totals()
        gt_i, gt_o = tw_file.totals_row
        findings.append(Finding(
            level="error", category="totals", item="",
            message=(
                f"TW raw item sum does not match the Grand Total row. "
                f"Item sum: Intransit {item_i:,.0f} / Onhand {item_o:,.0f}. "
                f"Grand Total: Intransit {gt_i:,.0f} / Onhand {gt_o:,.0f}. "
                f"Do not proceed - a row is missing or counted twice."
            ),
        ))

    for item, qty in tw_unmatched.items():
        findings.append(Finding(
            level="error", category="unmatched", item=item,
            message=f"'{item}' has stock in the TW raw file but no matching roster row.",
            intransit=qty["intransit"], onhand=qty["onhand"],
            suggestions=_suggest_matches(item, roster),
        ))

    for item, qty in tw_dropped.items():
        findings.append(Finding(
            level="info", category="rule", item=item,
            message=f"'{item}' excluded (drop rule).",
            intransit=qty["intransit"], onhand=qty["onhand"],
        ))

    for item, (target, qty) in tw_folded.items():
        findings.append(Finding(
            level="info", category="rule", item=item,
            message=f"'{item}' folded into '{target}' (fold rule).",
            intransit=qty["intransit"], onhand=qty["onhand"],
        ))

    for row in tw_file.rows_with_issues():
        non_ldu = [i for i in row.issues if i != "LDU prefix"]
        if non_ldu:
            findings.append(Finding(
                level="warning", category="cleaning", item=row.item,
                message=(
                    f"TW row {row.excel_row}: cleaned '{row.raw_item}' "
                    f"({', '.join(non_ldu)})."
                ),
                intransit=row.intransit, onhand=row.onhand,
            ))

    stats["tw_rows"] = len(tw_file.rows)
    stats["tw_ldu_merged"] = sum(1 for r in tw_file.rows if r.was_ldu)
    stats["tw_intransit"] = sum(v["TW_intransit"] for v in values.values())
    stats["tw_onhand"] = sum(v["TW_onhand"] for v in values.values())

    # --------------------------------------------------------------- AGT
    if agt_file is not None:
        agt_merged = agt_file.merged()
        agt_resolved, _, _, agt_unmatched = _apply_aliases_and_rules(
            agt_merged, roster_keys
        )
        for item, qty in agt_resolved.items():
            values[item]["AGT_intransit"] = qty["intransit"]
            values[item]["AGT_onhand"] = qty["onhand"]

        for item, qty in agt_unmatched.items():
            findings.append(Finding(
                level="error", category="unmatched", item=item,
                message=f"'{item}' has stock in the AGT raw file but no matching roster row.",
                intransit=qty["intransit"], onhand=qty["onhand"],
                suggestions=_suggest_matches(item, roster),
            ))

        stats["agt_rows"] = len(agt_file.rows)
        stats["agt_onhand"] = sum(v["AGT_onhand"] for v in values.values())

    # --------------------------------------------------------------- DCR
    if dcr_result is not None:
        for group, items in dcr_result.by_distributor.items():
            unmatched_qty = 0
            for item, qty in items.items():
                if item in values:
                    values[item][f"{group}_dcr"] = float(qty)
                else:
                    unmatched_qty += qty
            if unmatched_qty:
                findings.append(Finding(
                    level="warning", category="dcr", item="",
                    message=(
                        f"{unmatched_qty:,} DCR units for {group} sit on items "
                        f"outside the roster - they are not in the report."
                    ),
                ))

        for customer, qty in dcr_result.unknown_customers.items():
            findings.append(Finding(
                level="warning", category="dcr", item="",
                message=(
                    f"Distributor '{customer}' ({qty:,} units) is not in "
                    f"DISTRIBUTOR_MAP - excluded from the report."
                ),
            ))

        stats["dcr_totals"] = dcr_result.totals()
        stats["dcr_rows_scanned"] = dcr_result.rows_scanned
        stats["dcr_from_cache"] = dcr_result.from_cache

    # ---------------------------------------------------------------- QS
    if qs_manual:
        for item, qty in qs_manual.items():
            if item in values:
                values[item]["QS_intransit"] = float(qty.get("intransit") or 0)
                values[item]["QS_onhand"] = float(qty.get("onhand") or 0)
        stats["qs_intransit"] = sum(v["QS_intransit"] for v in values.values())
        stats["qs_onhand"] = sum(v["QS_onhand"] for v in values.values())
    else:
        stats["qs_intransit"] = 0.0
        stats["qs_onhand"] = 0.0

    return ReconResult(values=values, findings=findings, stats=stats)
