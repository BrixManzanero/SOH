"""
Text normalization for SOH item codes.

Ito ang pinaka-core na module. Lahat ng item code na papasok sa system
ay dadaan dito, para tiyak na pare-pareho ang porma bago mag-match.

Apat na klase ng sira ang inaayos nito (Klase A at B sa analysis):
  A. Whitespace  - NBSP (CHAR 160), double space, trailing/leading space
  B. LDU prefix  - "LDU CN7c 256+8" -> "CN7c 256+8"

Ang Klase C (baliktad na Item/Market) at Klase D (maling code) ay
HINDI dito inaayos - nasa aliases.py sila, kasi hindi sila mahuhulaan.
"""

import re
import unicodedata

# Mga karaniwang invisible characters na nakikita sa platform exports.
_INVISIBLE = {
    "\xa0": " ",   # non-breaking space (CHAR 160) - nasa DCR customer names
    "\u200b": "",  # zero-width space
    "\u200c": "",  # zero-width non-joiner
    "\u200d": "",  # zero-width joiner
    "\ufeff": "",  # byte-order mark
    "\u2007": " ",  # figure space
    "\u202f": " ",  # narrow no-break space
}

_LDU_RE = re.compile(r"^LDU\s+", re.IGNORECASE)


def clean_text(value) -> str:
    """
    Linisin ang anumang text value.

    Katumbas ito ng =TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," "))) sa Excel,
    pero mas kumpleto - kasama ang zero-width characters na hindi
    nahuhuli ng CLEAN().

    >>> clean_text("T1002 128+4 ")
    'T1002 128+4'
    >>> clean_text("Mega Pad Pro  KB")
    'Mega Pad Pro KB'
    >>> clean_text("BOBBITECH\\xa0ENTERPRISES\\xa0INC")
    'BOBBITECH ENTERPRISES INC'
    """
    if value is None:
        return ""

    text = str(value)

    # Palitan ang mga invisible characters.
    for bad, good in _INVISIBLE.items():
        text = text.replace(bad, good)

    # Alisin ang control characters (katumbas ng CLEAN()).
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C")

    # I-collapse ang lahat ng whitespace runs papuntang isang space,
    # tapos i-trim. Katumbas ng TRIM() pero pati tabs at newlines kasama.
    text = re.sub(r"\s+", " ", text).strip()

    return text


def strip_ldu(item: str) -> str:
    """
    Alisin ang LDU prefix.

    Ang LDU units ay display/demo stock na dapat isama sa base item.
    Hindi sila hiwalay na SKU.

    >>> strip_ldu("LDU CN7c 256+8")
    'CN7c 256+8'
    >>> strip_ldu("CN7c 256+8")
    'CN7c 256+8'
    """
    return _LDU_RE.sub("", item).strip()


def is_ldu(raw_item) -> bool:
    """Totoo kung ang raw item ay LDU row."""
    return bool(_LDU_RE.match(clean_text(raw_item)))


def item_key(value) -> str:
    """
    Ang canonical form ng isang item code: linis + tanggal ng LDU.

    Ito ang ginagamit na join key sa lahat ng source.
    """
    return strip_ldu(clean_text(value))


def match_key(value) -> str:
    """
    Case-insensitive na key para sa fuzzy matching.

    Ginagamit lang sa paghahanap ng suggestion para sa unmatched items -
    HINDI sa aktwal na pag-sum, para hindi tahimik na magkamali kapag
    may item na magkaiba lang sa capitalization (hal. CN7C vs CN7c).
    """
    return item_key(value).upper().replace(" ", "")


def describe_issues(raw_value) -> list:
    """
    Ilista ang mga natuklasang problema sa isang raw value.

    Ginagamit ito sa cleaning preview (Step 2) para makita ng user
    kung ano talaga ang inayos - hindi lang basta "nalinis na".
    """
    if raw_value is None:
        return []

    text = str(raw_value)
    issues = []

    if "\xa0" in text:
        issues.append("non-breaking space (CHAR 160)")
    if any(ch in text for ch in ("\u200b", "\u200c", "\u200d", "\ufeff")):
        issues.append("zero-width character")
    if "  " in text:
        issues.append("double space")
    if text != text.strip():
        if text != text.lstrip():
            issues.append("leading space")
        if text != text.rstrip():
            issues.append("trailing space")
    if any(ord(ch) < 32 for ch in text):
        issues.append("control character")
    if _LDU_RE.match(text.strip()):
        issues.append("LDU prefix")

    return issues
