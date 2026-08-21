"""
Aliases, equivalence groups and manual rules - the app's memory.

This is the answer to problem classes C and D (a market name landing
in the Item column, and an item code that does not match the roster).
A computer cannot infer either of them, but it can remember them.

The point is that a case answered once is never asked again. Each
report should ask fewer questions than the last.
"""

import json
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
ALIASES_PATH = CONFIG_DIR / "aliases.json"
RULES_PATH = CONFIG_DIR / "rules.json"


# ---------------------------------------------------------------- aliases

def load_aliases() -> dict:
    """
    Load the one-way alias map: {name_in_raw: item_code_in_report}

    Use this only when the raw name could never itself be a roster
    code. Examples:

      "BUDS 4 AIR" -> "BD04 Air"     (Item and Market columns swapped)
      "T1002"      -> "T1002 256+4"  (variant suffix missing)
    """
    if not ALIASES_PATH.exists():
        return {}
    with open(ALIASES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data.get("aliases", {})


def save_alias(raw_item: str, report_item: str) -> None:
    """Add or replace one alias and write it to disk immediately."""
    aliases = load_aliases()
    aliases[raw_item] = report_item
    _write_aliases(aliases)


def delete_alias(raw_item: str) -> None:
    aliases = load_aliases()
    aliases.pop(raw_item, None)
    _write_aliases(aliases)


def load_equivalents() -> list:
    """
    Equivalence groups: item codes that mean the SAME product but are
    written differently depending on the file and the week.

    Why this exists (found by comparing four weeks of reports):

        roster Aug 3, 6, 10 : TSP-W03A
        roster Aug 20       : TSP-WP03A
        TW raw writes       : TSP-W03A
        AGT raw writes      : TSP-WP03A

    A one-directional alias cannot handle this. If it points at
    TSP-WP03A it breaks the Aug 10 roster; if it points at TSP-W03A
    it breaks the Aug 20 roster.

    An equivalence group says only "these are the same thing". At run
    time the app looks at the roster in front of it and maps every
    other member onto whichever form that roster actually uses. It
    works in both directions and survives the roster being renamed.

    Returns a list of lists, e.g. [["TSP-W03A", "TSP-WP03A"]]
    """
    if not ALIASES_PATH.exists():
        return []
    with open(ALIASES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data.get("equivalents", [])


def save_equivalent_group(group: list) -> None:
    """Add an equivalence group, merging it with any that overlap."""
    groups = load_equivalents()
    incoming = set(group)

    merged = []
    for existing in groups:
        if incoming & set(existing):
            incoming |= set(existing)
        else:
            merged.append(existing)
    merged.append(sorted(incoming))

    aliases = load_aliases()
    _write_aliases(aliases, merged)


def _write_aliases(aliases: dict, equivalents: list = None) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if equivalents is None:
        equivalents = load_equivalents()

    existing_notes = {}
    if ALIASES_PATH.exists():
        try:
            with open(ALIASES_PATH, "r", encoding="utf-8") as fh:
                existing_notes = json.load(fh).get("_notes", {})
        except (json.JSONDecodeError, OSError):
            pass

    payload = {
        "_comment": (
            "aliases  = one-way mapping from a name in the raw file to the "
            "item code in the report. Use when the raw name will never be a "
            "roster code (e.g. a market name landing in the Item column).  "
            "equivalents = groups of codes that mean the same product. The "
            "app maps them onto whichever form the current roster uses, so "
            "they survive the roster being renamed."
        ),
        "aliases": dict(sorted(aliases.items())),
        "equivalents": sorted(equivalents),
        "_notes": existing_notes,
    }
    with open(ALIASES_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)


# ------------------------------------------------------------------ rules

def load_rules() -> dict:
    """
    Manual rules. Three kinds:

      drop - never include this item in the report
      fold - always roll this item's stock into another item
      ask  - ALWAYS flag this item, never decide automatically

    "ask" exists because some decisions genuinely change week to week.
    MegaPad 2 is the example: on Aug 10 its 1,135 units were NOT folded
    into T1103 256+8 (raw 582 -> report 582), but on Aug 20 they WERE
    (raw 332 -> report 867). A fixed rule is wrong either way, so the
    app stops and asks instead of silently guessing.
    """
    if not RULES_PATH.exists():
        return {"drop": [], "fold": {}, "ask": []}
    with open(RULES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return {
        "drop": data.get("drop", []),
        "fold": data.get("fold", {}),
        "ask": data.get("ask", []),
    }


def save_rules(rules: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "_comment": (
            "drop = never included in the report. "
            "fold = stock always rolls into another item code. "
            "ask  = always flagged for a decision, never automatic."
        ),
        "drop": sorted(set(rules.get("drop", []))),
        "fold": dict(sorted(rules.get("fold", {}).items())),
        "ask": sorted(set(rules.get("ask", []))),
    }
    with open(RULES_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)


def add_drop(item: str) -> None:
    rules = load_rules()
    if item not in rules["drop"]:
        rules["drop"].append(item)
    rules["fold"].pop(item, None)
    rules["ask"] = [a for a in rules["ask"] if a != item]
    save_rules(rules)


def add_fold(source_item: str, target_item: str) -> None:
    rules = load_rules()
    rules["fold"][source_item] = target_item
    rules["drop"] = [d for d in rules["drop"] if d != source_item]
    rules["ask"] = [a for a in rules["ask"] if a != source_item]
    save_rules(rules)


def add_ask(item: str) -> None:
    """Flag this item every time instead of deciding automatically."""
    rules = load_rules()
    if item not in rules["ask"]:
        rules["ask"].append(item)
    rules["drop"] = [d for d in rules["drop"] if d != item]
    rules["fold"].pop(item, None)
    save_rules(rules)


def clear_rule(item: str) -> None:
    rules = load_rules()
    rules["drop"] = [d for d in rules["drop"] if d != item]
    rules["fold"].pop(item, None)
    rules["ask"] = [a for a in rules["ask"] if a != item]
    save_rules(rules)


# --------------------------------------------------- distributor mapping

# Maps a Customer Name in the DCR export to a column group in the
# ND SOH report. Verified against the Aug 20 report.
#
# VSTECS PHILS INC is deliberately absent: it is a distributor in the
# DCR data but is not part of the ND SOH report (812 units on Aug 17).
#
# Kept for reference only. DCR is filled in by hand, so the app does
# not read the DCR export.
DISTRIBUTOR_MAP = {
    "BOBBITECH ENTERPRISES INC": "TW",
    "Qstar Mobile Inc": "QS",
    "AGT": "AGT",
}

# The DCR pivot is filtered on these two conditions before counting.
DCR_STATUS_FILTER = "Available"
DCR_CUSTOMER_TYPE_FILTER = "Distributor"
