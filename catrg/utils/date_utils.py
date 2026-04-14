import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

_DATE_LINE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def get_base_path() -> Path:
    """Return the application root directory, works both frozen and from source.

    Use this for read-only bundled assets (e.g. logo.jpg).
    For user-writable data, use :func:`get_data_path` instead.

    When frozen with PyInstaller (especially one-file), assets live under
    ``sys._MEIPASS``, not next to the executable.
    """
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent


def get_data_path() -> Path:
    """Return the user-data directory for configs, logs, and templates.

    On Windows this resolves to ``%LOCALAPPDATA%\\CATRG``.
    On other platforms it falls back to ``~/.local/share/CATRG``.
    The directory is created automatically if it does not exist.
    """
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path.home() / ".local" / "share"
    data_dir = base / "CATRG"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def format_datetime(raw: Optional[str], fmt: str = "%m/%d/%Y %H:%M:%S UTC") -> Optional[str]:
    """Parse an ISO-8601 datetime string and return it in *fmt*.

    Returns None on failure or if *raw* is None / 'N/A'.
    """
    if not raw or raw == "N/A":
        return None
    try:
        dt = datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
        return dt.strftime(fmt)
    except ValueError:
        return raw


def looks_like_date_line(line: str) -> bool:
    """Return True if *line* starts with a YYYY-MM-DD pattern (any year)."""
    return bool(_DATE_LINE_RE.match(line.strip()))
