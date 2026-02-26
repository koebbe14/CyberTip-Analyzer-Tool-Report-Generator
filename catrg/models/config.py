"""Credential, investigator-info, and recent-file management.

Credentials are stored via the OS keyring when available, falling back
to plain-JSON files with a logged warning.
"""

import json
import os
from dataclasses import dataclass, field
from typing import List, Optional

from catrg.utils.logger import get_logger
from catrg.utils.date_utils import get_data_path

log = get_logger(__name__)

_KEYRING_AVAILABLE = False
try:
    import keyring
    _KEYRING_AVAILABLE = True
except ImportError:
    log.warning(
        "keyring package not installed; credentials will be stored in plain-text JSON. "
        "Install it with:  pip install keyring"
    )

SERVICE_MAXMIND = "catrg-maxmind"
SERVICE_ARIN = "catrg-arin"


@dataclass
class InvestigatorInfo:
    name: str = ""
    title: str = ""


@dataclass
class MaxMindCredentials:
    account_id: str = ""
    license_key: str = ""

    @property
    def is_configured(self) -> bool:
        return bool(self.account_id and self.license_key)


@dataclass
class ArinCredentials:
    api_key: str = ""

    @property
    def is_configured(self) -> bool:
        return bool(self.api_key)


class ConfigManager:
    """Manages all persistent configuration for CAT-RG."""

    def __init__(self, base_path: Optional[str] = None):
        self._base = base_path or str(get_data_path())
        self.investigator = InvestigatorInfo()
        self.maxmind = MaxMindCredentials()
        self.arin = ArinCredentials()
        self.recent_files: List[str] = []

        self._investigator_file = os.path.join(self._base, "investigator_info.json")
        self._recent_files_file = os.path.join(self._base, "recent_files.json")
        self._maxmind_file = os.path.join(self._base, "maxmind_credentials.json")
        self._arin_file = os.path.join(self._base, "arin_credentials.json")

    # ── Investigator ──────────────────────────────────────────────

    def load_investigator(self) -> InvestigatorInfo:
        try:
            if os.path.exists(self._investigator_file):
                with open(self._investigator_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.investigator = InvestigatorInfo(
                    name=data.get("name", ""),
                    title=data.get("title", ""),
                )
        except Exception as e:
            log.error("Error loading investigator info: %s", e)
        return self.investigator

    def save_investigator(self) -> None:
        try:
            with open(self._investigator_file, "w", encoding="utf-8") as f:
                json.dump({"name": self.investigator.name, "title": self.investigator.title}, f)
        except Exception as e:
            log.error("Could not save investigator info: %s", e)

    # ── MaxMind ───────────────────────────────────────────────────

    def load_maxmind(self) -> MaxMindCredentials:
        if _KEYRING_AVAILABLE:
            try:
                aid = keyring.get_password(SERVICE_MAXMIND, "account_id") or ""
                lk = keyring.get_password(SERVICE_MAXMIND, "license_key") or ""
                if aid and lk:
                    self.maxmind = MaxMindCredentials(account_id=aid, license_key=lk)
                    return self.maxmind
            except Exception as e:
                log.warning("keyring read failed, falling back to file: %s", e)

        try:
            if os.path.exists(self._maxmind_file):
                with open(self._maxmind_file, "r", encoding="utf-8") as f:
                    creds = json.load(f)
                self.maxmind = MaxMindCredentials(
                    account_id=creds.get("account_id", ""),
                    license_key=creds.get("license_key", ""),
                )
        except Exception as e:
            log.error("Could not load MaxMind credentials: %s", e)
        return self.maxmind

    def save_maxmind(self) -> None:
        if _KEYRING_AVAILABLE:
            try:
                keyring.set_password(SERVICE_MAXMIND, "account_id", self.maxmind.account_id)
                keyring.set_password(SERVICE_MAXMIND, "license_key", self.maxmind.license_key)
                if os.path.exists(self._maxmind_file):
                    os.remove(self._maxmind_file)
                    log.info("Migrated MaxMind credentials to OS keyring; removed plain-text file")
                return
            except Exception as e:
                log.warning("keyring write failed, falling back to file: %s", e)

        try:
            with open(self._maxmind_file, "w", encoding="utf-8") as f:
                json.dump({"account_id": self.maxmind.account_id,
                           "license_key": self.maxmind.license_key}, f)
            log.warning("MaxMind credentials saved as plain-text JSON (install keyring for secure storage)")
        except Exception as e:
            log.error("Could not save MaxMind credentials: %s", e)

    # ── ARIN ──────────────────────────────────────────────────────

    def load_arin(self) -> ArinCredentials:
        if _KEYRING_AVAILABLE:
            try:
                ak = keyring.get_password(SERVICE_ARIN, "api_key") or ""
                if ak:
                    self.arin = ArinCredentials(api_key=ak)
                    return self.arin
            except Exception as e:
                log.warning("keyring read failed for ARIN, falling back to file: %s", e)

        try:
            if os.path.exists(self._arin_file):
                with open(self._arin_file, "r", encoding="utf-8") as f:
                    creds = json.load(f)
                self.arin = ArinCredentials(api_key=creds.get("api_key", ""))
        except Exception as e:
            log.error("Could not load ARIN credentials: %s", e)
        return self.arin

    def save_arin(self) -> None:
        if _KEYRING_AVAILABLE:
            try:
                keyring.set_password(SERVICE_ARIN, "api_key", self.arin.api_key)
                if os.path.exists(self._arin_file):
                    os.remove(self._arin_file)
                    log.info("Migrated ARIN API key to OS keyring; removed plain-text file")
                return
            except Exception as e:
                log.warning("keyring write failed for ARIN, falling back to file: %s", e)

        try:
            with open(self._arin_file, "w", encoding="utf-8") as f:
                json.dump({"api_key": self.arin.api_key}, f)
            log.warning("ARIN API key saved as plain-text JSON (install keyring for secure storage)")
        except Exception as e:
            log.error("Could not save ARIN credentials: %s", e)

    # ── Recent files ──────────────────────────────────────────────

    def load_recent_files(self) -> List[str]:
        try:
            if os.path.exists(self._recent_files_file):
                with open(self._recent_files_file, "r", encoding="utf-8") as f:
                    self.recent_files = json.load(f)
        except Exception as e:
            log.error("Error loading recent files: %s", e)
            self.recent_files = []
        return self.recent_files

    def save_recent_files(self) -> None:
        try:
            with open(self._recent_files_file, "w", encoding="utf-8") as f:
                json.dump(self.recent_files, f)
        except Exception as e:
            log.error("Error saving recent files: %s", e)

    def add_recent_file(self, path: str) -> None:
        if path in self.recent_files:
            self.recent_files.remove(path)
        self.recent_files.insert(0, path)
        self.recent_files = self.recent_files[:5]
        self.save_recent_files()
