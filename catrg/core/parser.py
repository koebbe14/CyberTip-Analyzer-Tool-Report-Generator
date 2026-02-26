"""Shared CyberTip JSON parsing logic.

This module centralises data extraction so that report generation,
Excel export, and comparison all use the same code paths.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Tuple

from catrg.utils.logger import get_logger
from catrg.utils.validators import is_valid_ip, safe_get
from catrg.utils.date_utils import format_datetime, looks_like_date_line

log = get_logger(__name__)


@dataclass
class IpOccurrence:
    datetime: Optional[str] = None
    port: Optional[str] = None
    event: Optional[str] = None


@dataclass
class EvidenceFile:
    file_number: int = 0
    filename: Optional[str] = None
    original_filename: Optional[str] = None
    submittal_id: Optional[str] = None
    verification_hash: Optional[str] = None
    viewed_by_esp: Optional[bool] = None
    upload_time: Optional[str] = None
    ncmec_tags: Optional[str] = None
    ip_addresses: List[str] = field(default_factory=list)
    additional_info: str = ""
    sent_date: Optional[str] = None
    from_email: Optional[str] = None
    to_email: Optional[str] = None


@dataclass
class WebpageInfo:
    number: int = 0
    url: Optional[str] = None
    type_info: Optional[str] = None
    text: Optional[str] = None
    additional_info: str = ""


@dataclass
class PersonInfo:
    fields: Dict[str, str] = field(default_factory=dict)
    screen_name: Optional[str] = None
    emails: List[dict] = field(default_factory=list)
    addresses: List[str] = field(default_factory=list)
    phones: List[dict] = field(default_factory=list)
    additional_contact: Optional[str] = None
    languages: List[str] = field(default_factory=list)
    races: List[str] = field(default_factory=list)
    disabilities: List[str] = field(default_factory=list)
    additional_disability_info: Optional[str] = None
    ip_logins: List[dict] = field(default_factory=list)
    esp_extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class MeetMeIpLog:
    timestamp: str = ""
    ip: str = ""
    device: str = "N/A"


@dataclass
class ParsedCyberTip:
    report_id: str = "N/A"
    date_received: Optional[str] = None
    esp_name: str = "N/A"
    incident_type: Optional[str] = None
    incident_datetime: Optional[str] = None
    incident_desc: Optional[str] = None
    reported_by: str = "N/A"
    is_bingimage: bool = False
    persons: List[PersonInfo] = field(default_factory=list)
    evidence_files: List[EvidenceFile] = field(default_factory=list)
    webpages: List[WebpageInfo] = field(default_factory=list)
    additional_informations: List[str] = field(default_factory=list)
    email_data_by_msg_id: Dict[str, dict] = field(default_factory=dict)
    meetme_messages: List[dict] = field(default_factory=list)
    meetme_ip_logs: List[MeetMeIpLog] = field(default_factory=list)
    all_ip_data: Dict[str, List[IpOccurrence]] = field(default_factory=dict)


def load_json(file_path: str) -> Optional[dict]:
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log.error("Error loading JSON file: %s", e)
        return None


def parse_email_incidents(data: dict) -> Dict[str, dict]:
    """Extract email metadata keyed by message ID."""
    result: Dict[str, dict] = {}
    for email in data.get("reportedInformation", {}).get("incidentDetails", {}).get("emailIncident", []):
        contents = email.get("contents", [{}])[0].get("value", "")
        sent_date = from_email = to_email = msg_id = None
        for line in contents.split("\n"):
            if line.startswith("Sent Date:"):
                sent_date = line[len("Sent Date:"):].strip()
            elif line.startswith("From:"):
                from_email = line[len("From:"):].strip()
            elif line.startswith("To:"):
                to_email = line[len("To:"):].strip()
            elif line.startswith("X-Ymail-Msg-Id:"):
                msg_id = line[len("X-Ymail-Msg-Id:"):].strip()
        if msg_id:
            result[msg_id] = {"sent_date": sent_date, "from": from_email, "to": to_email}
    return result


def _extract_upload_time(file_data: dict, esp_name: str) -> Optional[str]:
    """Return the upload time string for a file, handling ESP-specific locations."""
    source_info = file_data.get("sourceInformation")
    esp_metadata = file_data.get("espMetadata")

    if esp_name == "Dropbox, Inc." and esp_metadata and "metadatas" in esp_metadata:
        for md in esp_metadata.get("metadatas", []):
            if md.get("name") == "Upload Date and Time":
                return md.get("value")
        return None

    if "Imgur" in esp_name and source_info:
        for capture in source_info.get("sourceCaptures", []):
            if capture.get("captureType") == "IP Address" and capture.get("eventName") == "UPLOAD":
                return capture.get("dateTime")
        return None

    if source_info:
        captures = source_info.get("sourceCaptures", [])
        if captures:
            return captures[0].get("dateTime")
    return None


def _extract_file_ips(file_data: dict, esp_name: str) -> Tuple[List[str], str]:
    """Return (list_of_ip_strings, special_note) for a file."""
    ips: List[str] = []
    note = ""
    source_info = file_data.get("sourceInformation")
    if source_info and "sourceCaptures" in source_info:
        if "X Corp" in esp_name:
            note = "X does not capture IPs for individual Posts"
        else:
            for capture in source_info.get("sourceCaptures", []):
                if capture.get("captureType") == "IP Address":
                    val = safe_get(capture.get("value"))
                    if val:
                        ips.append(val)
    return ips, note


def _filter_additional_info(raw: str, esp_name: str) -> str:
    """Clean up ESP additional-info text by removing boilerplate."""
    if esp_name in ("Facebook", "Instagram, Inc."):
        lines = raw.split("\n")
        filtered = []
        for line in lines:
            line = line.strip()
            if any(phrase in line for phrase in [
                'With respect to the section "Was File Reviewed by Company?',
                'When Meta responds "Yes"',
                'When Meta responds "No"',
                "File's unique ESP Identifier:",
                "Messenger Thread ID:",
                "Uploaded ",
            ]):
                continue
            if line and line != "The content can be found in this report.":
                filtered.append(line)
        return "\n".join(filtered).strip()

    if esp_name == "WhatsApp Inc.":
        paragraphs = raw.split("\n\n")
        filtered = [p.replace("\n", " ") for p in paragraphs if '{"report_surface":' not in p]
        return "\n\n".join(filtered).strip()

    lines = raw.split("\n")
    filtered = [l for l in lines if "The content can be found in this report" not in l]
    return "\n".join(filtered).strip()


def _extract_ncmec_tags(file_data: dict) -> Optional[str]:
    ncmec_tags = file_data.get("ncmecTags")
    if ncmec_tags is None:
        return None
    try:
        tags = ", ".join(
            tag.get("value") for tag in ncmec_tags.get("groups", [{}])[0].get("tags", [])
        )
        return tags if tags else None
    except (IndexError, TypeError):
        return None


def extract_ip_addresses(data: dict) -> Tuple[Dict[str, List[IpOccurrence]], Dict[str, List[IpOccurrence]]]:
    """Extract IPs from persons and files.

    Returns (ip_data_for_querying, all_ip_data).
    """
    seen: Set[tuple] = set()
    entries: List[Tuple[str, IpOccurrence]] = []
    unique_ips: Set[str] = set()

    def collect(ip: str, occ: IpOccurrence) -> None:
        key = (ip, occ.datetime, occ.port, occ.event)
        if key not in seen:
            seen.add(key)
            entries.append((ip, occ))
            unique_ips.add(ip)

    for person in data.get("reportedInformation", {}).get("reportedPeople", {}).get("reportedPersons", []):
        si = person.get("sourceInformation")
        if si:
            for cap in si.get("sourceCaptures", []):
                ip = cap.get("value")
                if ip and is_valid_ip(ip):
                    collect(ip, IpOccurrence(
                        datetime=cap.get("dateTime"),
                        port=cap.get("port"),
                        event=cap.get("eventName"),
                    ))

    for f in data.get("reportedInformation", {}).get("uploadedFiles", {}).get("uploadedFiles", []):
        si = f.get("sourceInformation")
        if si:
            for cap in si.get("sourceCaptures", []):
                ip = cap.get("value")
                if ip and is_valid_ip(ip):
                    collect(ip, IpOccurrence(
                        datetime=cap.get("dateTime"),
                        port=cap.get("port"),
                        event=cap.get("eventName"),
                    ))

    all_ip: Dict[str, List[IpOccurrence]] = {}
    for ip, occ in entries:
        all_ip.setdefault(ip, []).append(occ)

    return all_ip, all_ip


def parse_cybertip(data: dict) -> ParsedCyberTip:
    """Parse a full CyberTip JSON into a structured object."""
    tip = ParsedCyberTip()
    tip.report_id = data.get("reportId", "N/A")

    raw_date = data.get("dateReceived", "N/A")
    tip.date_received = format_datetime(raw_date, "%m/%d/%Y") or raw_date

    ri = data.get("reportedInformation", {})
    reporter = ri.get("reportingEsp", {})
    tip.esp_name = reporter.get("espName", "N/A")

    incident = ri.get("incidentSummary", {})
    tip.incident_type = safe_get(incident.get("incidentType"))
    tip.incident_datetime = format_datetime(safe_get(incident.get("incidentDateTime")))
    if not tip.incident_datetime:
        tip.incident_datetime = safe_get(incident.get("incidentDateTime"))
    tip.incident_desc = safe_get(incident.get("incidentDateTimeDescription"))

    reported_by = tip.esp_name
    if "Microsoft" in tip.esp_name:
        last_name = reporter.get("lastName", "N/A")
        if last_name == "Microsoft BingImage":
            reported_by += ", BingImage"
            tip.is_bingimage = True
    tip.reported_by = reported_by

    for info in ri.get("additionalInformations", []):
        val = safe_get(info.get("value"))
        if val:
            tip.additional_informations.append(val)

    tip.email_data_by_msg_id = parse_email_incidents(data)

    _parse_persons(data, tip)
    _parse_evidence(data, tip)
    _parse_webpages(data, tip)
    _parse_meetme_extras(data, tip)

    all_ip, _ = extract_ip_addresses(data)
    tip.all_ip_data = all_ip

    return tip


def _parse_persons(data: dict, tip: ParsedCyberTip) -> None:
    esp = tip.esp_name
    persons_raw = data.get("reportedInformation", {}).get("reportedPeople", {}).get("reportedPersons", [])
    for person in persons_raw:
        pi = PersonInfo()

        field_map = {
            "First Name": "firstName", "Middle Name": "middleName", "Last Name": "lastName",
            "Preferred Name": "preferredName", "Gender": "gender",
            "Preferred Pronouns": "preferredPronouns", "Date of Birth": "dateOfBirth",
            "Approximate Age": "approximateAge", "Physical Description": "physicalDescription",
            "Vehicle Description": "vehicleDescription", "Vehicle Tag Number": "vehicleTagNumber",
            "Occupation": "occupation", "ESP Service": "espService", "ESP User ID": "espUserId",
            "IP Address": "ipAddress", "Relationship to Reporter": "relationshipToReporter",
            "Relationship to Child Victim": "relationshipToChildVictim",
            "Access to Child Victim": "accessToChildVictim",
            "Access to Children": "accessToChildren", "Access to Firearms": "accessToFirearms",
            "Convicted Sex Offender": "convictedSexOffender",
            "Aware of Report": "awareOfReport", "Gang Affiliation": "gangAffiliation",
        }
        for label, key in field_map.items():
            val = safe_get(person.get(key))
            if val is not None:
                pi.fields[label] = str(val)

        sn = person.get("screenName") or {}
        pi.screen_name = safe_get(sn.get("value")) if isinstance(sn, dict) else None

        emails_data = person.get("emails") or {}
        for em in emails_data.get("emails", []):
            val = safe_get(em.get("value"))
            if val:
                verified = safe_get(em.get("verified"))
                pi.emails.append({"value": val, "verified": verified})

        addresses_data = person.get("addresses")
        if addresses_data:
            for addr in addresses_data.get("addresses", []):
                parts = []
                for k in ("street1", "street2", "city"):
                    v = safe_get(addr.get(k))
                    if v:
                        parts.append(v)
                state = safe_get(addr.get("state"))
                postal = safe_get(addr.get("postalCode"))
                country = safe_get(addr.get("country"))
                if state:
                    parts.append(state)
                if postal:
                    parts.append(postal)
                if country:
                    parts.append(country)
                if parts:
                    pi.addresses.append(" ".join(parts))

        phones_data = person.get("phones")
        if phones_data:
            for ph in phones_data.get("phones", []):
                val = safe_get(ph.get("value"))
                if val:
                    pi.phones.append({"value": val, "verified": safe_get(ph.get("verified"))})

        pi.additional_contact = safe_get(person.get("additionalContactInformation"))

        langs = person.get("languages")
        if langs and langs != "N/A":
            pi.languages = langs if isinstance(langs, list) else [langs]

        races = person.get("races")
        if races and races != "N/A":
            pi.races = races if isinstance(races, list) else [races]

        disab = person.get("disabilities")
        if disab and disab != "N/A":
            pi.disabilities = disab if isinstance(disab, list) else [disab]

        pi.additional_disability_info = safe_get(person.get("additionalDisabilityInformation"))

        if "TikTok" in esp and person.get("sourceInformation"):
            for cap in person.get("sourceInformation", {}).get("sourceCaptures", []):
                if cap.get("captureType") == "IP Address":
                    ip_val = safe_get(cap.get("value"))
                    if ip_val:
                        pi.ip_logins.append({
                            "ip": ip_val,
                            "datetime": safe_get(cap.get("dateTime")) or "N/A",
                            "event": safe_get(cap.get("eventName")) or "N/A",
                        })

        if "Imgur" in esp:
            for info in data.get("reportedInformation", {}).get("additionalInformations", []):
                val = safe_get(info.get("value"))
                if val:
                    pi.esp_extra["imgur_additional"] = val

        if "X Corp" in esp:
            _parse_xcorp_person(person, pi)

        if "MeetMe" in esp:
            _parse_meetme_person(data, person, pi)

        tip.persons.append(pi)


def _parse_xcorp_person(person: dict, pi: PersonInfo) -> None:
    additional_info = person.get("additionalInformations", [{}])[0].get("value", "")
    source_info = person.get("sourceInformation", {})
    profile_url = None
    if source_info and "sourceCaptures" in source_info:
        for cap in source_info.get("sourceCaptures", []):
            if cap.get("captureType") == "Profile URL":
                profile_url = cap.get("value")
                break
    full_name = location = description = None
    for line in additional_info.split("\n"):
        if line.startswith("Full Name:"):
            full_name = line[len("Full Name:"):].strip()
        elif line.startswith("Location:"):
            location = line[len("Location:"):].strip()
        elif line.startswith("Description:"):
            description = line[len("Description:"):].strip()
    if full_name and full_name != "N/A":
        pi.esp_extra["full_name"] = full_name
    if location and location != "N/A":
        pi.esp_extra["location"] = location
    if description and description != "N/A":
        pi.esp_extra["description"] = description
    if profile_url and profile_url != "N/A":
        pi.esp_extra["profile_url"] = profile_url


def _parse_meetme_person(data: dict, person: dict, pi: PersonInfo) -> None:
    for info in data.get("reportedInformation", {}).get("additionalInformations", []):
        val = safe_get(info.get("value"))
        if val and "Registration details from Suspect's MeetMe profile" in val:
            profile: Dict[str, str] = {}
            for line in val.split("\n"):
                mappings = [
                    ("MeetMe Profile Name:", "profile_name"),
                    ("MeetMe UserID:", "user_id"),
                    ("DOB:", "dob"),
                    ("Age:", "age"),
                    ("Zip:", "zip"),
                    ("City:", "city"),
                    ("State:", "state"),
                    ("Email:", "email"),
                    ("Date Joined meetme.com:", "date_joined"),
                    ("Registration IP:", "registration_ip"),
                    ("Phone number used to verify account:", "phone"),
                    ("Recent GPS Data:", "gps"),
                ]
                for prefix, key in mappings:
                    if line.startswith(prefix):
                        val_str = line[len(prefix):].strip()
                        if key == "gps" and "Lat./Long.:" in line:
                            val_str = line.split("Lat./Long.:")[1].strip() if "Lat./Long.:" in line else val_str
                        profile[key] = val_str
                        break
            pi.esp_extra["meetme_profile"] = profile

    if person.get("sourceInformation"):
        for cap in person.get("sourceInformation", {}).get("sourceCaptures", []):
            if cap.get("captureType") == "IP Address":
                ip_val = safe_get(cap.get("value"))
                if ip_val:
                    pi.ip_logins.append({"ip": ip_val, "datetime": "N/A", "event": "login"})


def _parse_evidence(data: dict, tip: ParsedCyberTip) -> None:
    files_raw = data.get("reportedInformation", {}).get("uploadedFiles", {}).get("uploadedFiles", [])
    for i, f in enumerate(files_raw, 1):
        ef = EvidenceFile(file_number=i)
        ef.filename = safe_get(f.get("filename"))
        ef.original_filename = safe_get(f.get("originalFilename"))

        if tip.esp_name in ("Facebook", "Instagram, Inc."):
            ef.submittal_id = safe_get(f.get("submittalId"))

        ef.verification_hash = safe_get(f.get("verificationHash"))
        ef.viewed_by_esp = f.get("viewedByEsp")
        ef.upload_time = _extract_upload_time(f, tip.esp_name)
        ef.ncmec_tags = _extract_ncmec_tags(f)

        ips, note = _extract_file_ips(f, tip.esp_name)
        ef.ip_addresses = ips
        if note:
            ef.additional_info = note

        additional_info = f.get("additionalInformations", [])
        msg_id = None
        if additional_info:
            for info in additional_info:
                value = info.get("value", "")
                if value.startswith("Message ID:"):
                    msg_id = value[len("Message ID:"):].strip()
                    break

        if msg_id and msg_id in tip.email_data_by_msg_id:
            ed = tip.email_data_by_msg_id[msg_id]
            ef.sent_date = ed.get("sent_date")
            ef.from_email = ed.get("from")
            ef.to_email = ed.get("to")

        raw_desc = additional_info[0].get("value", "") if additional_info else ""
        if raw_desc:
            ef.additional_info = _filter_additional_info(raw_desc, tip.esp_name)

        tip.evidence_files.append(ef)


def _parse_webpages(data: dict, tip: ParsedCyberTip) -> None:
    webpages = data.get("reportedInformation", {}).get("incidentDetails", {}).get("webpageIncident", [])
    for i, wp in enumerate(webpages, 1):
        wi = WebpageInfo(number=i)

        si = wp.get("sourceInformation", {})
        if si and "sourceCaptures" in si:
            for cap in si.get("sourceCaptures", []):
                if cap.get("captureType") == "URL":
                    wi.url = safe_get(cap.get("value"))
                    break

        addl = wp.get("additionalInformations", [])
        if addl:
            info_text = addl[0].get("value", "")
            for line in info_text.split("\n"):
                if line.startswith("Type:"):
                    wi.type_info = line[len("Type:"):].strip()
                elif line.startswith("Text:"):
                    wi.text = line[len("Text:"):].strip()
            if "Reddit" in tip.esp_name:
                wi.additional_info = info_text

        tip.webpages.append(wi)


def _parse_meetme_extras(data: dict, tip: ParsedCyberTip) -> None:
    if "MeetMe" not in tip.esp_name:
        return

    for info in data.get("reportedInformation", {}).get("additionalInformations", []):
        val = safe_get(info.get("value"))
        if not val:
            continue

        # Parse IP login history
        for line in val.split("\n"):
            if line.strip() and looks_like_date_line(line):
                parts = line.split()
                if len(parts) >= 3:
                    tip.meetme_ip_logs.append(MeetMeIpLog(
                        timestamp=f"{parts[0]} {parts[1]}",
                        ip=parts[2],
                        device=parts[3].strip("()") if len(parts) > 3 and parts[3].startswith("(") else "N/A",
                    ))

        # Parse private messages
        lines = val.split("\n")
        messages: List[dict] = []
        current_msg: dict = {}
        capture = False
        for line in lines:
            if "Complete private message correspondence" in line:
                capture = True
                continue
            if not capture:
                continue
            if looks_like_date_line(line):
                break
            if not line.strip():
                continue
            if line.startswith("To:"):
                if current_msg:
                    messages.append(current_msg)
                current_msg = {"To": line[len("To:"):].strip()}
            elif line.startswith("From:"):
                current_msg["From"] = line[len("From:"):].strip()
            elif line.startswith("Sent:"):
                current_msg["Sent"] = line[len("Sent:"):].strip()
            elif line.startswith("Subject:"):
                current_msg["Subject"] = line[len("Subject:"):].strip()
            elif line.startswith("Message:"):
                current_msg["Message"] = line[len("Message:"):].strip()
        if current_msg:
            messages.append(current_msg)
        tip.meetme_messages.extend(messages)


def extract_comparison_data(data: dict) -> dict:
    """Extract identifiers for cross-tip comparison."""
    tip = parse_cybertip(data)
    result: Dict[str, set] = {
        "ips": set(),
        "emails": set(),
        "phones": set(),
        "screen_names": set(),
        "user_ids": set(),
        "hashes": set(),
    }

    for ip in tip.all_ip_data:
        result["ips"].add(ip)

    for person in tip.persons:
        for em in person.emails:
            result["emails"].add(em["value"].lower())
        for ph in person.phones:
            result["phones"].add(ph["value"])
        if person.screen_name:
            result["screen_names"].add(person.screen_name.lower())
        uid = person.fields.get("ESP User ID")
        if uid:
            result["user_ids"].add(uid)

    for ef in tip.evidence_files:
        if ef.verification_hash:
            result["hashes"].add(ef.verification_hash)
        for ip in ef.ip_addresses:
            result["ips"].add(ip)

    return {k: v for k, v in result.items() if v}
