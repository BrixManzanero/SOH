"""
Text normalization for SOH item codes.

This is the core module. Every item code that enters the system passes
through here, so that everything is in the same shape before matching.

It fixes two of the four classes of problem found in the raw files:

  A. Whitespace  - non-breaking space (CHAR 160), double spaces,
                   leading and trailing spaces
  B. LDU prefix  - "LDU CN7c 256+8" becomes "CN7c 256+8"

Classes C and D are NOT handled here:

  C. Swapped columns - the market name lands in the Item column
                       (e.g. "BUDS 4 AIR" instead of "BD04 Air")
  D. Wrong item code - "TSP-W03A" where the roster says "TSP-WP03A"

Those cannot be inferred from the text alone, so they live in
config.py as aliases and equivalence groups.
"""

import re
import unicodedata

# Invisible characters that turn up in platform exports.
# The non-breaking space is the common one: it appears in DCR customer
# names such as "BOBBITECH\xa0ENTERPRISES\xa0INC" and silently breaks
# any Excel lookup, because it looks identical to a normal space.
_INVISIBLE = {
    "\xa0": " ",   # non-breaking space (CHAR 160)
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
    Clean any text value.

    This is the equivalent of the Excel formula
    =TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," "))) but more thorough:
    CLEAN() does not remove zero-width characters, and this does.

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

    # Replace the invisible characters listed above.
    for bad, good in _INVISIBLE.items():
        text = text.replace(bad, good)

    # Strip control characters. This is what Excel's CLEAN() does.
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C")

    # Collapse every run of whitespace to a single space, then trim.
    # Equivalent to TRIM(), but this also catches tabs and newlines.
    text = re.sub(r"\s+", " ", text).strip()

    return text


def strip_ldu(item: str) -> str:
    """
    Remove the LDU prefix.

    LDU rows are display or demo units. They belong to the base item
    rather than being a separate SKU, so their stock is merged in.

    >>> strip_ldu("LDU CN7c 256+8")
    'CN7c 256+8'
    >>> strip_ldu("CN7c 256+8")
    'CN7c 256+8'
    """
    return _LDU_RE.sub("", item).strip()


def is_ldu(raw_item) -> bool:
    """True when the raw item is an LDU row."""
    return bool(_LDU_RE.match(clean_text(raw_item)))


def item_key(value) -> str:
    """
    The canonical form of an item code: cleaned, with LDU removed.

    This is the join key used across every source.
    """
    return strip_ldu(clean_text(value))


def match_key(value) -> str:
    """
    A loose key for finding near-matches.

    Used only to suggest candidates for an unmatched item. It is
    deliberately NOT used for summing, because ignoring case and spaces
    would quietly merge codes that differ only in capitalisation
    (CN7C versus CN7c), and a silent wrong merge is worse than a
    question on screen.
    """
    return item_key(value).upper().replace(" ", "")


def describe_issues(raw_value) -> list:
    """
    List the problems found in a raw value.

    Feeds the cleaning preview, so the user can see exactly what was
    changed rather than just being told the data was "cleaned".
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
