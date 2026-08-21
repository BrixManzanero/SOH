"""
Alias map at manual rules - ang "memory" ng app.

Ito ang sagot sa Klase C at D na problema (baliktad na Item/Market,
at maling item code). Hindi kayang hulaan ng computer ang mga ito,
pero kayang tandaan.

Ang punto: kapag sinagot mo na ang isang kaso, hindi ka na tatanungin ulit.
Bawat report, paunti nang paunti ang tanong.
"""

import json
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
ALIASES_PATH = CONFIG_DIR / "aliases.json"
RULES_PATH = CONFIG_DIR / "rules.json"


# ---------------------------------------------------------------- aliases

def load_aliases() -> dict:
    """
    Basahin ang alias map: {raw_item_sa_TW: item_code_sa_report}

    Halimbawa:
      "BUDS 4 AIR" -> "BD04 Air"     (baliktad ang Item at Market sa raw)
      "TSP-W03A"   -> "TSP-WP03A"    (kulang ng letrang P)
    """
    if not ALIASES_PATH.exists():
        return {}
    with open(ALIASES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data.get("aliases", {})


def save_alias(raw_item: str, report_item: str) -> None:
    """Idagdag o palitan ang isang alias, tapos i-save agad sa disk."""
    aliases = load_aliases()
    aliases[raw_item] = report_item
    _write_aliases(aliases)


def delete_alias(raw_item: str) -> None:
    aliases = load_aliases()
    aliases.pop(raw_item, None)
    _write_aliases(aliases)


def _write_aliases(aliases: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "_comment": (
            "Mapping mula sa item name sa TW/AGT raw file papunta sa "
            "item code na ginagamit sa ND SOH report. Idinadagdag ito "
            "ng app kapag may sinagot kang unmatched item."
        ),
        "aliases": dict(sorted(aliases.items())),
    }
    with open(ALIASES_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)


# ------------------------------------------------------------------ rules

def load_rules() -> dict:
    """
    Basahin ang manual rules.

    Dalawang klase:
      drop  - listahan ng item na hindi isinasama sa report
      fold  - {galing: papunta} - isinasama ang stock sa ibang item
    """
    if not RULES_PATH.exists():
        return {"drop": [], "fold": {}}
    with open(RULES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return {
        "drop": data.get("drop", []),
        "fold": data.get("fold", {}),
    }


def save_rules(rules: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "_comment": (
            "drop = item na hindi isinasama sa report (binubura sa raw). "
            "fold = item na isinasama ang stock sa ibang item code."
        ),
        "drop": sorted(set(rules.get("drop", []))),
        "fold": dict(sorted(rules.get("fold", {}).items())),
    }
    with open(RULES_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)


def add_drop(item: str) -> None:
    rules = load_rules()
    if item not in rules["drop"]:
        rules["drop"].append(item)
    rules["fold"].pop(item, None)
    save_rules(rules)


def add_fold(source_item: str, target_item: str) -> None:
    rules = load_rules()
    rules["fold"][source_item] = target_item
    rules["drop"] = [d for d in rules["drop"] if d != source_item]
    save_rules(rules)


def clear_rule(item: str) -> None:
    rules = load_rules()
    rules["drop"] = [d for d in rules["drop"] if d != item]
    rules["fold"].pop(item, None)
    save_rules(rules)


# --------------------------------------------------- distributor mapping

# Mapping mula sa Customer Name sa DCR export papunta sa column group
# sa ND SOH report. Na-verify laban sa Aug 20 report.
#
# Ang VSTECS PHILS INC ay sinasadyang wala dito - hindi siya kasama
# sa ND SOH report (812 units noong Aug 17).
DISTRIBUTOR_MAP = {
    "BOBBITECH ENTERPRISES INC": "TW",
    "Qstar Mobile Inc": "QS",
    "AGT": "AGT",
}

# Sinasala natin ang DCR sa dalawang kondisyon na ito bago magbilang.
DCR_STATUS_FILTER = "Available"
DCR_CUSTOMER_TYPE_FILTER = "Distributor"
