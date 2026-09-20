"""Normalise identifiers and build mailbox search queries."""
from __future__ import annotations
import re

_WS = re.compile(r"\s+")

def clean(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.upper() in {"NA", "N/A", "NONE"}:
        return None
    # "RFQ 6100175731" -> "6100175731"
    s = re.sub(r"^(RFQ|PO|REF)[\s:#-]*", "", s, flags=re.I)
    return _WS.sub(" ", s)

def is_valid_job_no(job_no: str | None) -> bool:
    """Job numbers are 7 digits. Flags typos such as 77372045."""
    return bool(job_no and re.fullmatch(r"\d{7}", job_no))

def search_terms(item: dict) -> list[tuple[str, str]]:
    """Ordered (label, term) pairs. Job No first because it is unique."""
    terms = []
    for key in ("job_no", "rfq", "order_ref"):
        v = clean(item.get(key))
        if v and len(v) >= 5:
            terms.append((key, v))
    # dedupe while keeping order
    seen, out = set(), []
    for k, v in terms:
        if v not in seen:
            seen.add(v); out.append((k, v))
    return out

def graph_search_string(term: str) -> str:
    # Graph $search uses KQL; quote the term so hyphens/slashes are literal.
    return f'"{term}"'
