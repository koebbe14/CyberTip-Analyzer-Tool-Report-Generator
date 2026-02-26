import ipaddress
from typing import Any, Optional

from catrg.utils.logger import get_logger

log = get_logger(__name__)


def is_valid_ip(ip: str) -> bool:
    """Return True if *ip* is a valid IPv4 or IPv6 address."""
    try:
        ipaddress.ip_address(ip)
        return True
    except (ValueError, TypeError):
        return False


def validate_cybertip_json(data: Any) -> tuple[bool, str]:
    """Check that *data* looks like a valid NCMEC CyberTip JSON structure.

    Returns (valid, reason).
    """
    if not isinstance(data, dict):
        return False, "Top-level JSON value is not an object"

    if "reportId" not in data:
        return False, "Missing 'reportId' field"

    ri = data.get("reportedInformation")
    if not isinstance(ri, dict):
        return False, "Missing or invalid 'reportedInformation' section"

    if "reportingEsp" not in ri:
        return False, "Missing 'reportingEsp' in reportedInformation"

    return True, ""


def safe_get(value: Any, default: str = "N/A") -> Optional[str]:
    """Normalise a field value: return *default* for None or literal 'N/A'."""
    if value is None or value == "N/A":
        return None
    return value
