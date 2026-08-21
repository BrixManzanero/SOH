"""
Reconciliation engine - where every source is brought together.

The roster (the list of items) comes from the template, never from
the raw files. The template is therefore always authoritative: the
app will not invent a row that is not already there.

The important output is not only the numbers but the findings: what
failed to match, what was excluded, and what deserves a second look
before the report is submitted.
"""

from collections import defaultdict
from dataclasses import dataclass, field

from .normalize import match_key
from .config import load_aliases, load_rules, load_equivalents


@dataclass
class Finding:
    """One thing the user needs to know about."""
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


def _build_equivalence_resolver(roster_keys: dict) -> dict:
    """
    Turn equivalence groups into a concrete {code: roster_code} map,
    decided by the roster actually in front of us.

    For each group, find which member the roster uses, then point every
    other member at it. If the roster uses none of them (or more than
    one), the group is skipped and the items fall through to the normal
    unmatched handling - better to ask than to guess.
    """
    resolver = {}
    for group in load_equivalents():
        present = [code for code in group if code in roster_keys]
        if len(present) != 1:
            continue
        target = present[0]
        for code in group:
            if code != target:
                resolver[code] = target
    return resolver


def _apply_aliases_and_rules(merged: dict, roster_keys: dict):
    """
    Apply the alias map, equivalence groups and manual rules.

    Returns: (resolved, dropped, folded, unmatched)
    """
    aliases = load_aliases()
    equivalents = _build_equivalence_resolver(roster_keys)
    rules = load_rules()
    drop_set = set(rules["drop"])
    fold_map = rules["fold"]
    ask_set = set(rules.get("ask", []))

    resolved = defaultdict(lambda: {"intransit": 0.0, "onhand": 0.0})
    dropped = {}
    folded = {}
    unmatched = {}

    for item, qty in merged.items():
        # 0. Rule: always ask - decision varies week to week
        if item in ask_set:
            if qty["intransit"] or qty["onhand"]:
                unmatched[item] = qty
            continue

        # 1. Rule: drop
        if item in drop_set:
            if qty["intransit"] or qty["onhand"]:
                dropped[item] = qty
            continue

        # 2. Rule: fold into another item
        target = fold_map.get(item)
        if target:
            target = equivalents.get(target, target)
            folded[item] = (target, qty)
            resolved[target]["intransit"] += qty["intransit"]
            resolved[target]["onhand"] += qty["onhand"]
            continue

        # 3. One-way alias (raw name -> report code)
        target = aliases.get(item, item)

        # 4. Equivalence: settle on the spelling THIS roster uses
        target = equivalents.get(target, target)

        # 5. Is it in the roster?
        if target in roster_keys:
            resolved[target]["intransit"] += qty["intransit"]
            resolved[target]["onhand"] += qty["onhand"]
        else:
            if qty["intransit"] or qty["onhand"]:
                unmatched[item] = qty

    return dict(resolved), dropped, folded, unmatched


def _suggest_matches(item: str, roster: list, limit: int = 3) -> list:
    """
    Find roster codes that look close to an unmatched item.

    These are suggestions to speed up the user's choice. Nothing is
    selected automatically: a wrong silent match is worse than a
    question on screen.
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
            # Score on the length of the shared leading run.
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
    Combine every source into a single values dict.

    roster      : item codes from the template - authoritative
    tw_file     : StockFile for TW
    agt_file    : StockFile for AGT (may be None)
    dcr_result  : reserved; DCR is filled in by hand, so normally None
    qs_manual   : reserved; QS is filled in by hand, so normally None
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
