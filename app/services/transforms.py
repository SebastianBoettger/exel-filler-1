import re
from typing import Optional, Tuple

def split_street_house(value: str) -> Tuple[str, str]:
    if value is None:
        return "", ""
    s = str(value).strip()
    if not s:
        return "", ""
    m = re.match(r"^(.*?)(?:\s+)(\d+[a-zA-Z]?(?:[-/]\d+[a-zA-Z]?)?)\s*$", s)
    if not m:
        return s, ""
    return m.group(1).strip(), m.group(2).strip()

def normalize_phone(value: str) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    plus = s.startswith("+")
    digits = re.sub(r"\D+", "", s)
    return ("+" + digits) if plus else digits

def state_from_zip_de(zip_code: str) -> Optional[str]:
    try:
        import pgeocode
    except Exception:
        return None

    if not zip_code:
        return None
    z = re.sub(r"\D+", "", str(zip_code))
    if len(z) != 5:
        return None

    nomi = pgeocode.Nominatim("de")
    r = nomi.query_postal_code(z)

    state = getattr(r, "state_name", None)
    if isinstance(state, str) and state.strip():
        return state.strip()

    state2 = getattr(r, "state", None)
    if isinstance(state2, str) and state2.strip():
        return state2.strip()

    return None
