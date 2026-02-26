"""Police report and IP report generation from parsed CyberTip data."""

from __future__ import annotations

import json
import os
from typing import Callable, Dict, List, Optional

from catrg.core.parser import (
    ParsedCyberTip,
    IpOccurrence,
    parse_cybertip,
)
from catrg.core.ip_lookup import IpLookupService, IpLookupResult
from catrg.models.statements import (
    StatementManager,
    DEFAULT_STATEMENTS,
    build_placeholder_context,
    substitute_placeholders,
)
from catrg.utils.date_utils import format_datetime, get_data_path
from catrg.utils.logger import get_logger

log = get_logger(__name__)


# ── Report template (section visibility / ordering / naming) ──────

DEFAULT_SECTIONS = [
    "incident_summary",
    "suspect_information",
    "evidence_summary",
    "ip_analysis",
]

DEFAULT_SECTION_NAMES = {
    "incident_summary": "INCIDENT SUMMARY",
    "suspect_information": "SUSPECT INFORMATION",
    "evidence_summary": "EVIDENCE SUMMARY",
    "ip_analysis": "IP ADDRESS ANALYSIS",
}


class ReportTemplate:
    """Controls which sections appear, their order, custom names,
    and user-defined custom sections.  Persists to disk as named profiles.
    """

    _profiles_dir: str = ""

    def __init__(self, name: str = "Default"):
        self.name: str = name
        self.section_order: List[str] = list(DEFAULT_SECTIONS)
        self.section_visible: Dict[str, bool] = {s: True for s in DEFAULT_SECTIONS}
        self.section_names: Dict[str, str] = dict(DEFAULT_SECTION_NAMES)
        self.custom_sections: List[Dict[str, str]] = []

    def is_visible(self, section: str) -> bool:
        return self.section_visible.get(section, True)

    def get_section_name(self, section_id: str) -> str:
        return self.section_names.get(section_id, section_id.upper().replace("_", " "))

    # ── Serialisation ─────────────────────────────────────────────

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "section_order": self.section_order,
            "section_visible": self.section_visible,
            "section_names": self.section_names,
            "custom_sections": self.custom_sections,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "ReportTemplate":
        t = cls(name=d.get("name", "Default"))
        t.section_order = d.get("section_order", list(DEFAULT_SECTIONS))
        t.section_visible = d.get("section_visible", {s: True for s in DEFAULT_SECTIONS})
        t.section_names = d.get("section_names", dict(DEFAULT_SECTION_NAMES))
        t.custom_sections = d.get("custom_sections", [])
        return t

    # ── Profile persistence ───────────────────────────────────────

    @classmethod
    def _ensure_dir(cls) -> str:
        if not cls._profiles_dir:
            cls._profiles_dir = os.path.join(str(get_data_path()), "report_templates")
        os.makedirs(cls._profiles_dir, exist_ok=True)
        return cls._profiles_dir

    def save_profile(self) -> None:
        d = self._ensure_dir()
        safe = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in self.name)
        path = os.path.join(d, f"{safe}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=4)
        log.info("Saved template profile: %s", path)

    @classmethod
    def load_profile(cls, name: str) -> "ReportTemplate":
        d = cls._ensure_dir()
        safe = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in name)
        path = os.path.join(d, f"{safe}.json")
        with open(path, "r", encoding="utf-8") as f:
            return cls.from_dict(json.load(f))

    @classmethod
    def list_profiles(cls) -> List[str]:
        d = cls._ensure_dir()
        profiles = []
        for fn in sorted(os.listdir(d)):
            if fn.endswith(".json"):
                profiles.append(fn[:-5])
        return profiles

    @classmethod
    def delete_profile(cls, name: str) -> None:
        d = cls._ensure_dir()
        safe = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in name)
        path = os.path.join(d, f"{safe}.json")
        if os.path.exists(path):
            os.remove(path)
            log.info("Deleted template profile: %s", name)

    # ── Import / Export ───────────────────────────────────────────

    def export_to_file(self, filepath: str) -> None:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump({"catrg_template_export": True, **self.to_dict()}, f, indent=4)

    @classmethod
    def import_from_file(cls, filepath: str) -> "ReportTemplate":
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "section_order" not in data:
            raise ValueError("Invalid template export file")
        return cls.from_dict(data)


# ── Summary statistics ────────────────────────────────────────────

def generate_summary_stats(tip: ParsedCyberTip, ip_results: Dict[str, IpLookupResult]) -> str:
    lines = ["SUMMARY STATISTICS:", "=" * 40]
    lines.append(f"CyberTip ID: {tip.report_id}")
    lines.append(f"ESP: {tip.esp_name}")
    lines.append(f"Date Received: {tip.date_received}")
    lines.append(f"Total Suspect Records: {len(tip.persons)}")
    lines.append(f"Total Evidence Files: {len(tip.evidence_files)}")
    lines.append(f"Total Unique IPs: {len(tip.all_ip_data)}")

    countries = {r.geo.country for r in ip_results.values() if r.geo.country and r.geo.country != "N/A"}
    if countries:
        lines.append(f"Countries Identified: {', '.join(sorted(countries))}")

    isps = {r.whois.organization for r in ip_results.values() if r.whois.organization and r.whois.organization != "N/A"}
    if isps:
        lines.append(f"ISPs/Orgs Identified: {len(isps)}")

    lines.append("=" * 40 + "\n")
    return "\n".join(lines)


# ── Police report ─────────────────────────────────────────────────

def generate_police_report(
    tip: ParsedCyberTip,
    data: dict,
    stmts: StatementManager,
    investigator_title: str,
    investigator_name: str,
    template: Optional[ReportTemplate] = None,
    ip_service: Optional[IpLookupService] = None,
) -> str:
    tmpl = template or ReportTemplate()

    ctx = build_placeholder_context(
        investigator_name=investigator_name,
        investigator_title=investigator_title,
        tip_id=tip.report_id,
        esp_name=tip.esp_name,
        date_received=tip.date_received or "",
        suspect_name=_first_suspect_name(tip),
        total_files=len(tip.evidence_files),
        total_ips=len(tip.all_ip_data),
    )

    intro_raw = stmts.get_text("intro")
    intro = substitute_placeholders(intro_raw, ctx) + "\n\n"

    report = stmts.get_for_prefix("at_beginning:", data, ctx) + intro

    section_builders = {
        "incident_summary": lambda: _build_incident_section(tip, data, stmts, tmpl),
        "suspect_information": lambda: _build_suspect_section(tip, tmpl),
        "evidence_summary": lambda: _build_evidence_section(tip, data, stmts, tmpl),
        "ip_analysis": lambda: _build_meetme_ip_section(tip, ip_service) if "MeetMe" in tip.esp_name and ip_service else "",
    }

    for section in tmpl.section_order:
        if not tmpl.is_visible(section):
            continue
        prefix_before = _section_prefix_map(section, "before")
        prefix_after = _section_prefix_map(section, "after")
        if prefix_before:
            report += stmts.get_for_prefix(prefix_before, data, ctx)
        builder = section_builders.get(section)
        if builder:
            report += builder()
        if prefix_after:
            report += stmts.get_for_prefix(prefix_after, data, ctx)
        report += "\n"

    for cs in tmpl.custom_sections:
        report += f"\n{cs.get('name', 'CUSTOM SECTION').upper()}:\n"
        body = cs.get("body", "")
        report += substitute_placeholders(body, ctx) + "\n"

    report += stmts.get_end_statements(data, ctx)
    return report


def _first_suspect_name(tip: ParsedCyberTip) -> str:
    if tip.persons:
        p = tip.persons[0]
        for label in ("Display Name", "First Name", "Name"):
            val = p.fields.get(label)
            if val and val != "N/A":
                return val
    return ""


def _section_prefix_map(section: str, direction: str) -> str:
    mapping = {
        "incident_summary": {"before": "before_incident:", "after": "after_incident:"},
        "suspect_information": {"before": "before_suspect:", "after": "after_suspect:"},
        "evidence_summary": {"before": "before_evidence:", "after": "after_evidence:"},
        "ip_analysis": {"before": "before_ip:", "after": "after_ip:"},
    }
    return mapping.get(section, {}).get(direction, "")


def _build_incident_section(tip: ParsedCyberTip, data: dict, stmts: StatementManager, tmpl: ReportTemplate) -> str:
    header = tmpl.get_section_name("incident_summary")
    lines = [f"{header}:\n"]
    if tip.incident_type:
        lines.append(f"Incident Type: {tip.incident_type}")
    if tip.incident_datetime:
        lines.append(f"Incident Date/Time: {tip.incident_datetime}")
    if tip.reported_by and tip.reported_by != "N/A":
        lines.append(f"Reported By: {tip.reported_by}")
    if tip.incident_desc:
        lines.append(f"Incident Date/Time Description: {tip.incident_desc}")

    if "Reddit" in tip.esp_name and tip.additional_informations:
        for info in tip.additional_informations:
            lines.append(f"Additional Information: {info}")

    if tip.is_bingimage and "bingimage" in stmts.selected:
        lines.append(stmts.get_text("bingimage"))

    if "X Corp" in tip.esp_name and "xcorp" in stmts.selected:
        lines.append(stmts.get_text("xcorp"))

    return "\n".join(lines) + "\n"


def _build_suspect_section(tip: ParsedCyberTip, tmpl: ReportTemplate) -> str:
    if not tip.persons:
        return ""
    header = tmpl.get_section_name("suspect_information")
    lines = [f"{header}:\n"]
    esp = tip.esp_name

    for person in tip.persons:
        for label, val in person.fields.items():
            lines.append(f"{label}: {val}")

        if person.screen_name:
            lines.append(f"Screen Name: {person.screen_name}")

        for em in person.emails:
            line = f"Email: {em['value']}"
            if "MeetMe" in esp:
                line += " (Note: This email was voluntarily provided by the user and may not be verified by MeetMe)"
            elif em.get("verified"):
                line += f" (Verified: {em['verified']})"
            lines.append(line)

        for addr in person.addresses:
            lines.append(f"Address: {addr}")

        for ph in person.phones:
            line = f"Phone: {ph['value']}"
            if ph.get("verified"):
                line += f" (Verified: {ph['verified']})"
            lines.append(line)

        if person.additional_contact:
            lines.append(f"Additional Contact Information: {person.additional_contact}")
        if person.languages:
            lines.append(f"Languages: {', '.join(person.languages)}")
        if person.races:
            lines.append(f"Races: {', '.join(person.races)}")
        if person.disabilities:
            lines.append(f"Disabilities: {', '.join(person.disabilities)}")
        if person.additional_disability_info:
            lines.append(f"Additional Disability Information: {person.additional_disability_info}")

        for login in person.ip_logins:
            lines.append(f"IP Address (Login): {login['ip']}")
            lines.append(f"Login Date/Time: {login['datetime']}")
            lines.append(f"Event: {login['event']}")

        if person.esp_extra.get("imgur_additional"):
            lines.append(f"Additional Information from ESP:\n{person.esp_extra['imgur_additional']}")

        if "X Corp" in esp:
            for key, label in [("full_name", "Full Name"), ("location", "Location"),
                               ("description", "Description"), ("profile_url", "Profile URL")]:
                if person.esp_extra.get(key):
                    lines.append(f"{label}: {person.esp_extra[key]}")

        if "MeetMe" in esp and person.esp_extra.get("meetme_profile"):
            mp = person.esp_extra["meetme_profile"]
            lines.append("Registration details from Suspect's MeetMe profile (provided by visitor, and is NOT verified):")
            if mp.get("profile_name"):
                lines.append(f"MeetMe Profile Name: {mp['profile_name']}")
            if mp.get("user_id"):
                lines.append(f"MeetMe UserID: {mp['user_id']}")
            if mp.get("dob"):
                lines.append(f"Date of Birth: {mp['dob']}")
            if mp.get("age"):
                lines.append(f"Approximate Age: {mp['age']}")
            addr_parts = []
            if mp.get("city"):
                addr_parts.append(f"{mp['city']},")
            if mp.get("state"):
                addr_parts.append(mp["state"])
            if mp.get("zip"):
                addr_parts.append(mp["zip"])
            if addr_parts:
                lines.append(f"Address: {' '.join(addr_parts)}")
            if mp.get("email"):
                lines.append(
                    f"Email: {mp['email']} "
                    "(Note: This email was voluntarily provided by the user and may not be verified by MeetMe)"
                )
            if mp.get("date_joined"):
                lines.append(f"Date Joined MeetMe: {mp['date_joined']}")
            if mp.get("registration_ip"):
                lines.append(f"Registration IP: {mp['registration_ip']}")
            if mp.get("phone") and mp["phone"] != "N/A":
                lines.append(f"Phone: {mp['phone']}")
            if mp.get("gps"):
                lines.append(f"Recent GPS Data: Lat./Long.: {mp['gps']}")
                lines.append(
                    "Note: This GPS data is used for business purposes and may or may not "
                    "be indicative of a user's true geographic location."
                )

        lines.append("")

    return "\n".join(lines)


def _build_evidence_section(tip: ParsedCyberTip, data: dict, stmts: StatementManager, tmpl: ReportTemplate) -> str:
    if not tip.evidence_files and not ("Reddit" in tip.esp_name and tip.webpages):
        return ""

    header = tmpl.get_section_name("evidence_summary")
    lines = [f"{header}:\n"]
    esp = tip.esp_name

    if "X Corp" in esp and tip.webpages:
        lines.append("Webpage/URL Information:")
        for wp in tip.webpages:
            lines.append(f"  Webpage {wp.number}:")
            if wp.url:
                lines.append(f"    URL: {wp.url}")
            if wp.type_info:
                lines.append(f"    Type: {wp.type_info}")
            if wp.text:
                lines.append(f"    Text: {wp.text}")
        lines.append("")

    if "Reddit" in esp and tip.webpages:
        lines.append("Reddit Chat Information:")
        for wp in tip.webpages:
            if wp.additional_info:
                lines.append(f"Additional Information: {wp.additional_info}")
        lines.append("")

    meta_stmt = stmts.get_text("meta") if "meta" in stmts.selected else ""

    for ef in tip.evidence_files:
        lines.append(f"FILE NUMBER {ef.file_number}:\n")
        if ef.filename:
            lines.append(f"File Name: {ef.filename}")
        if ef.submittal_id:
            lines.append(f"NCMEC Identifier: {ef.submittal_id}")
        if ef.verification_hash:
            lines.append(f"MD5 Hash: {ef.verification_hash}")
        if ef.sent_date:
            lines.append(f"Sent Date: {ef.sent_date}")
        if ef.from_email:
            lines.append(f"Emailed From: {ef.from_email}")
        if ef.to_email:
            lines.append(f"Emailed To: {ef.to_email}")
        if ef.original_filename:
            lines.append(f"Original Filename: {ef.original_filename}")

        if ef.additional_info and ef.additional_info != "X does not capture IPs for individual Posts":
            lines.append(f"Additional Information from ESP:\n{ef.additional_info}")

        if "MeetMe" in esp and ef.file_number == 1 and tip.meetme_messages:
            lines.append("Private Message Correspondence:")
            prev_subject = None
            for msg in tip.meetme_messages:
                lines.append(f"- To: {msg.get('To', 'N/A')}")
                lines.append(f"- From: {msg.get('From', 'N/A')}")
                lines.append(f"- Sent: {msg.get('Sent', 'N/A')}")
                subj = msg.get("Subject", "N/A")
                if subj != prev_subject and subj != "N/A":
                    lines.append(f"- Subject: {subj}")
                    prev_subject = subj
                lines.append(f"- Message: {msg.get('Message', 'N/A')}\n")

        if ef.viewed_by_esp is True:
            lines.append("Viewed by ESP: Yes")
            if esp in ("Instagram, Inc.", "Facebook") and meta_stmt:
                lines.append(meta_stmt + "\n")
        elif ef.viewed_by_esp is False:
            lines.append("Viewed by ESP: No")
            if esp in ("Instagram, Inc.", "Facebook") and meta_stmt:
                lines.append(meta_stmt + "\n")
        else:
            lines.append("Viewed by ESP: Unknown\n")

        if ef.upload_time:
            formatted = format_datetime(ef.upload_time)
            lines.append(f"Upload Date/Time: {formatted or ef.upload_time}")
        if ef.ncmec_tags:
            lines.append(f"NCMEC Tags: {ef.ncmec_tags}")

        if ef.ip_addresses:
            lines.append("Upload IP Address:")
            if "X Corp" in esp:
                lines.append(
                    "X does not capture IPs for individual Posts, but information from the "
                    "log of IPs provided by X for the timeframe relevant to the upload date/time "
                    "will be documented below."
                )
            else:
                for ip in ef.ip_addresses:
                    lines.append(ip)

        lines.append("\nThis file was viewed by the reporting Investigator\n")
        lines.append("Investigator's Description:\n")
        lines.append("=" * 50 + "\n")

    return "\n".join(lines)


def _build_meetme_ip_section(tip: ParsedCyberTip, svc: IpLookupService) -> str:
    lines = ["IP ADDRESS ANALYSIS:\n"]
    if not tip.meetme_ip_logs:
        lines.append("No IP address login history found in the provided data.\n")
        return "\n".join(lines)

    lines.append(
        "The following IP addresses and login timestamps were extracted from the "
        "MeetMe login history, with geolocation and ownership details:\n"
    )
    for entry in tip.meetme_ip_logs:
        lines.append(f"Login Date/Time: {entry.timestamp}")
        lines.append(f"IP Address: {entry.ip}")
        if entry.device != "N/A":
            lines.append(f"Device: {entry.device}")

        result = svc.lookup(entry.ip)
        lines.append("MaxMind Geolocation Data:")
        if result.geo.error:
            lines.append(f"  Error: {result.geo.error}")
        else:
            lines.append(f"  Country: {result.geo.country}")
            lines.append(f"  City: {result.geo.city}")
        lines.append("ARIN WHOIS Data:")
        if result.whois.error:
            lines.append(f"  Error: {result.whois.error}")
        else:
            lines.append(f"  Organization: {result.whois.organization}")
            if result.whois.registry != "ARIN":
                lines.append(f"  Registry: {result.whois.registry}")
        lines.append("")

    return "\n".join(lines) + "\n"


# ── IP address analysis report ────────────────────────────────────

def generate_ip_report(
    all_ip_data: Dict[str, List[IpOccurrence]],
    ip_service: IpLookupService,
    queried_ips: Optional[set] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> str:
    report_lines = ["IP ADDRESS ANALYSIS:\n"]
    total_unique = len(all_ip_data)
    report_lines.append(f"Total Unique IP Addresses: {total_unique}\n")

    ips_to_query = list(queried_ips) if queried_ips else list(all_ip_data.keys())
    ip_results = ip_service.lookup_batch(ips_to_query, progress_callback=progress_callback)

    for ip, occurrences in all_ip_data.items():
        report_lines.append(f"IP Address: {ip}")
        report_lines.append("Occurrences:")
        for occ in occurrences:
            formatted = format_datetime(occ.datetime)
            report_lines.append(f"      - Date/Time: {formatted or occ.datetime or 'N/A'}")
            if occ.port:
                report_lines.append(f"        Port: {occ.port}")
            report_lines.append(f"        IP Event: {occ.event or 'N/A'}")

        result = ip_results.get(ip)
        if result:
            report_lines.append("\nMaxMind Geolocation Data:")
            if result.geo.error:
                report_lines.append(f"        Error: {result.geo.error}")
            else:
                report_lines.append(f"        Country: {result.geo.country}")
                report_lines.append(f"        City: {result.geo.city}")
            report_lines.append("\nARIN WHOIS Data:")
            if result.whois.error:
                report_lines.append(f"        Error: {result.whois.error}")
            else:
                report_lines.append(f"        Organization: {result.whois.organization}")
                if result.whois.registry and result.whois.registry not in ("ARIN", "N/A"):
                    report_lines.append(f"        Registry: {result.whois.registry}")
        else:
            report_lines.append("\nMaxMind Geolocation Data:\n        Not queried (IP limit applied)")
            report_lines.append("\nARIN WHOIS Data:\n        Not queried (IP limit applied)")

        report_lines.append("\n" + "=" * 50 + "\n")

    return "\n".join(report_lines)
