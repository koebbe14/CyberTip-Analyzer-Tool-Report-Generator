"""Statement management with safe condition evaluation, formatting metadata,
placeholder substitution, ordering, and import/export.
"""

import json
import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Set

from catrg.utils.logger import get_logger
from catrg.utils.date_utils import get_data_path

log = get_logger(__name__)

PLACEMENT_PREFIXES = {
    "At Beginning of Report": "at_beginning:",
    "After introduction paragraph": "after_intro:",
    "Before Incident Summary": "before_incident:",
    "After Incident Summary": "after_incident:",
    "Before Suspect Information": "before_suspect:",
    "After Suspect Information": "after_suspect:",
    "Before Evidence Summary": "before_evidence:",
    "After Evidence Summary": "after_evidence:",
    "Before IP Address Analysis": "before_ip:",
    "After IP Address Analysis": "after_ip:",
    "After main sections (before final custom block)": "after_all_sections:",
    "At End of Report": "",
}

PREFIX_TO_PLACEMENT = {v: k for k, v in PLACEMENT_PREFIXES.items()}


def _custom_section_id(name: str) -> str:
    safe = "".join(c if c.isalnum() or c == "_" else "_" for c in name.lower())
    return f"custom_{safe}"


def get_all_placement_prefixes(template=None) -> Dict[str, str]:
    """Return PLACEMENT_PREFIXES extended with Before/After entries for
    each custom section in *template*."""
    result = dict(PLACEMENT_PREFIXES)
    if template:
        for cs in getattr(template, "custom_sections", []):
            name = cs.get("name", "")
            if not name:
                continue
            sid = _custom_section_id(name)
            result[f"Before {name}"] = f"before_{sid}:"
            result[f"After {name}"] = f"after_{sid}:"
    return result

DEFAULT_FORMATTING = {
    "font_size": 12,
    "bold": False,
    "italic": False,
    "indent": 0.0,
    "highlight": "",
}

DEFAULT_INTRO = (
    "On [CURRENT_DATE], I, [INVESTIGATOR_TITLE] [INVESTIGATOR_NAME], "
    "reviewed Cybertip #[CYBERTIP_NUMBER], which was received by the National Center "
    "for Missing and Exploited Children (NCMEC) on [REPORT_DATE_RECEIVED]. "
    "I observed the following information regarding this CyberTip:"
)

DEFAULT_STATEMENTS: Dict[str, str] = {
    "intro": DEFAULT_INTRO,
    "meta": (
        "\nWhen Meta responds \"Yes\" it means the contents of the file were viewed by an "
        "employee or contractor at Meta concurrently with or immediately before the file was "
        "submitted to NCMEC. When Meta responds \"No\" it means that while the contents of the "
        "file were not reviewed concurrently with or immediately before the file was submitted "
        "to NCMEC, historically at least one employee or contractor at Meta viewed a file whose "
        "hash matched the hash of the reported content and determined it contained apparent "
        "child pornography.\n\n"
        "For video files, when Meta responds \"Yes\" it means the entire contents of the file "
        "were viewed by an employee or contractor at Meta concurrently with or immediately "
        "before the file was submitted to NCMEC. When Meta responds \"No\" it means that while "
        "the contents of the file were not reviewed concurrently with or immediately before the "
        "file was submitted to NCMEC, historically at least one employee or contractor at Meta "
        "viewed a file and determined it contained apparent child pornography, and that file's "
        "hash matched a violating portion or the entirety of the reported content.\n"
    ),
    "ip_intro": (
        "\nThese following IP addresses were reported in the Cybertip. Each IP address was "
        "queried through the American Registry for Internet Numbers (ARIN) and Maxmind.com. "
        "ARIN is responsible for managing and distributing Internet number resources, like IP "
        "addresses, in North America. It is one of the five Regional Internet Registries (RIRs) "
        "worldwide, working under the global Internet Assigned Numbers Authority (IANA). ARIN "
        "maintains a public database (WHOIS) which tracks who holds what IP address. Maxmind is "
        "a company that provides a webservice IP geolocation tool. It should be noted that the "
        "estimated geographical location obtained from Maxmind's geolocation database are not "
        "always accurate, particularly when resolving IP addresses utilized by cellular "
        "providers, as a mobile user's location is constantly changing. The exact location of "
        "where an IP addresses geographically resolves to, along with the subscriber details, "
        "can only be obtained through legal process served to the provider.\n"
    ),
    "bingimage": (
        "\n\"BingImage\" (referred to as Visual Search) is a service of Microsoft's Bing search "
        "engine that provides similar images to an image provided by the user. This image can be "
        "provided either via upload or as a URL. The date/time provided indicates the time at "
        "which the image was received and evaluated by the BingImage service.\n"
    ),
    "xcorp": (
        "\n\n X retains different types of information for different time periods. Given X's "
        "real-time nature, some information may only be stored for a very brief period of time. "
        "\n\nFor accounts reported to NCMEC, X provides a copy of the preserved files, within "
        "the CyberTip report in the form of a .zip file which may be uploaded in multiple parts. "
        "\n\nAll times reported by X are in UTC. \n\nThe incident date/time is the timestamp from "
        "the most recent reported Post; however, if a Post is not reported, then the incident "
        "date/time will represent the account creation timestamp. \n\nX logs IPs in connection "
        "with user authentications to X (i.e., sessions, which may span multiple days) rather "
        "than individual Post postings; as a result, X is unable to provide insight into which "
        "IP address a specific Post was posted from. While X does not capture IPs for individual "
        "Posts, X provided a log of IPs for the timeframe relevant to the report. \n\nAll IP "
        "addresses in this report are associated with log-ins to the specific user account "
        "identified in the report.\n"
    ),
}


def evaluate_condition(condition: str, data: dict) -> bool:
    """Safely evaluate a condition string against report data.

    Supported formats (constructed by the UI):
      - esp_name == "Facebook"
      - esp_name in ["Facebook", "Instagram"]
      - has_evidence
      - has_ips
      - is_multi_tip
      - suspect_count > 1
      - Multiple data checks combined with && (AND), e.g. has_evidence&&has_ips

    No eval() is used.
    """
    if not condition:
        return True

    condition = condition.strip()
    if "&&" in condition:
        parts = [p.strip() for p in condition.split("&&") if p.strip()]
        if not parts:
            return True
        return all(evaluate_condition(p, data) for p in parts)

    esp_name = (
        data.get("reportedInformation", {})
        .get("reportingEsp", {})
        .get("espName", "N/A")
    )

    reported_info = data.get("reportedInformation") or {}
    evidence_files = (reported_info.get("fileDetails") or {}).get("uploadedFiles") or []
    ip_list = (reported_info.get("internetDetails") or {}).get("ipCaptureEvents") or []
    persons = (reported_info.get("reportedPeople") or [])
    recipients = (reported_info.get("intendedRecipients") or [])

    try:
        if condition == "has_evidence":
            return len(evidence_files) > 0
        if condition == "has_ips":
            return len(ip_list) > 0
        if condition == "is_multi_tip":
            return data.get("_is_multi_tip", False)
        if condition.startswith("suspect_count"):
            total = len(persons) + len(recipients)
            m = re.match(r"suspect_count\s*(>|>=|==|<|<=)\s*(\d+)", condition)
            if m:
                op, val = m.group(1), int(m.group(2))
                ops = {">": total > val, ">=": total >= val, "==": total == val,
                       "<": total < val, "<=": total <= val}
                return ops.get(op, False)
            return False

        if " in " in condition and "[" in condition:
            field, values_str = condition.split(" in ", 1)
            field = field.strip()
            values_str = values_str.strip()
            if values_str.startswith("[") and values_str.endswith("]"):
                values = re.findall(r'"([^"]*)"', values_str)
                if field == "esp_name":
                    return any(v.lower() in esp_name.lower() for v in values)
            return False

        if "==" in condition:
            field, value = condition.split("==", 1)
            field = field.strip()
            value = value.strip().strip('"')
            if field == "esp_name":
                return value.lower() in esp_name.lower()
            return False

    except Exception as e:
        log.error("Error evaluating condition '%s': %s", condition, e)

    return False


def substitute_placeholders(text: str, context: Dict[str, str]) -> str:
    """Replace [PLACEHOLDER] tokens with values from *context*."""
    for key, val in context.items():
        text = text.replace(f"[{key}]", str(val))
    return text


def build_placeholder_context(
    investigator_name: str = "",
    investigator_title: str = "",
    tip_id: str = "",
    esp_name: str = "",
    date_received: str = "",
    suspect_name: str = "",
    total_files: int = 0,
    total_ips: int = 0,
    agency_name: str = "",
    case_number: str = "",
    suspect_email: str = "",
    suspect_phone: str = "",
    suspect_screen_name: str = "",
    incident_date: str = "",
    incident_type: str = "",
) -> Dict[str, str]:
    """Build the placeholder-substitution context dict."""
    return {
        "CURRENT_DATE": datetime.now().strftime("%m-%d-%Y"),
        "INVESTIGATOR_NAME": investigator_name,
        "INVESTIGATOR_TITLE": investigator_title,
        "CYBERTIP_NUMBER": tip_id,
        "ESP_NAME": esp_name,
        "REPORT_DATE_RECEIVED": date_received or "N/A",
        "SUSPECT_NAME": suspect_name or "N/A",
        "TOTAL_FILES": str(total_files),
        "TOTAL_IPS": str(total_ips),
        "AGENCY_NAME": agency_name or "[AGENCY_NAME]",
        "CASE_NUMBER": case_number or "[CASE_NUMBER]",
        "SUSPECT_EMAIL": suspect_email or "N/A",
        "SUSPECT_PHONE": suspect_phone or "N/A",
        "SUSPECT_SCREEN_NAME": suspect_screen_name or "N/A",
        "INCIDENT_DATE": incident_date or "N/A",
        "INCIDENT_TYPE": incident_type or "N/A",
        "EVIDENCE_COUNT": str(total_files),
    }


def get_formatting(stmt_value: Any) -> dict:
    """Extract formatting metadata from a statement value, with defaults."""
    fmt = dict(DEFAULT_FORMATTING)
    if isinstance(stmt_value, dict) and "formatting" in stmt_value:
        fmt.update(stmt_value["formatting"])
    return fmt


class StatementManager:
    """Load, save, and query customisable report statements."""

    def __init__(self, base_path: Optional[str] = None):
        self._base = base_path or str(get_data_path())
        self._file = os.path.join(self._base, "custom_statements.json")
        self.statements: Dict[str, Any] = {}
        self.selected: Set[str] = set()

    def load(self) -> None:
        try:
            if os.path.exists(self._file):
                with open(self._file, "r", encoding="utf-8") as f:
                    self.statements = json.load(f)
            else:
                self.statements = dict(DEFAULT_STATEMENTS)
        except Exception as e:
            log.error("Error loading statements: %s", e)
            self.statements = dict(DEFAULT_STATEMENTS)

        for k, v in DEFAULT_STATEMENTS.items():
            if k not in self.statements:
                self.statements[k] = v

        self.selected = set(self.statements.keys())

    def save(self) -> None:
        try:
            with open(self._file, "w", encoding="utf-8") as f:
                json.dump(self.statements, f, indent=4)
        except Exception as e:
            log.error("Could not save custom statements: %s", e)

    def get_text(self, key: str) -> str:
        """Return the text for a statement key, falling back to defaults."""
        val = self.statements.get(key, DEFAULT_STATEMENTS.get(key, ""))
        if isinstance(val, dict):
            return val.get("text", "")
        return val

    def get_order(self, key: str) -> int:
        """Return the sort-order for a statement (lower = earlier)."""
        val = self.statements.get(key)
        if isinstance(val, dict):
            return val.get("order", 999)
        return 999

    def _sorted_keys_for_prefix(self, prefix: str) -> List[str]:
        """Return statement keys matching *prefix*, sorted alphabetically by key."""
        keys = [
            k for k in self.statements
            if k not in DEFAULT_STATEMENTS and k.startswith(prefix) and k in self.selected
        ]
        return sorted(keys, key=lambda k: k.lower())

    def get_for_prefix(self, prefix: str, data: dict, context: Optional[Dict[str, str]] = None) -> str:
        """Return concatenated custom statements matching *prefix*."""
        parts = []
        for key in self._sorted_keys_for_prefix(prefix):
            value = self.statements[key]
            if isinstance(value, dict):
                if not evaluate_condition(value.get("condition", ""), data):
                    continue
                sub_key = key[len(prefix):].strip()
                text = value["text"]
            else:
                sub_key = key[len(prefix):].strip()
                text = value

            if context:
                text = substitute_placeholders(text, context)

            if sub_key:
                parts.append(f"{sub_key.upper()}: {text}")
            else:
                parts.append(text)

        if parts:
            return "\n\n" + "\n\n".join(parts) + "\n\n"
        return ""

    def get_end_statements(self, data: dict, context: Optional[Dict[str, str]] = None,
                            template=None) -> str:
        """Return custom statements without a specific section prefix."""
        all_pfx = get_all_placement_prefixes(template)
        active_prefixes = [v for v in all_pfx.values() if v]
        keys = [
            k for k in self.statements
            if k not in DEFAULT_STATEMENTS
            and not any(k.startswith(p) for p in active_prefixes)
            and k in self.selected
        ]
        keys.sort(key=lambda k: k.lower())

        parts = []
        for key in keys:
            value = self.statements[key]
            if isinstance(value, dict):
                if not evaluate_condition(value.get("condition", ""), data):
                    continue
                text = value["text"]
            else:
                text = value
            if context:
                text = substitute_placeholders(text, context)
            parts.append(f"{key.upper()}: {text}")

        if parts:
            return "\n\nCUSTOM STATEMENTS:\n" + "\n\n".join(parts)
        return ""

    def get_placement_label(self, key: str, template=None) -> str:
        if key in DEFAULT_STATEMENTS:
            return "Default Statement"
        all_prefixes = get_all_placement_prefixes(template)
        rev = {v: k for k, v in all_prefixes.items()}
        for prefix, label in rev.items():
            if prefix and key.startswith(prefix):
                return label
        return "At End of Report"

    # ── Import / Export ───────────────────────────────────────────

    def export_to_file(self, filepath: str) -> None:
        """Export all custom (non-default) statements to a JSON file."""
        custom = {k: v for k, v in self.statements.items() if k not in DEFAULT_STATEMENTS}
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump({"catrg_statements_export": True, "statements": custom}, f, indent=4)
        log.info("Exported %d statements to %s", len(custom), filepath)

    def import_from_file(self, filepath: str, mode: str = "merge") -> int:
        """Import statements from a JSON file.

        *mode*: 'merge' keeps existing, 'overwrite' replaces on conflict.
        Returns the number of statements imported.
        """
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)

        stmts = data.get("statements", data)
        if not isinstance(stmts, dict):
            raise ValueError("Invalid statement export file format")

        count = 0
        for key, val in stmts.items():
            if key in DEFAULT_STATEMENTS:
                continue
            if mode == "merge" and key in self.statements:
                continue
            self.statements[key] = val
            self.selected.add(key)
            count += 1
        return count
