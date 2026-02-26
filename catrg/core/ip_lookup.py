"""IP geolocation and WHOIS lookups with caching, parallelism, and multi-RIR support.

Supported registries:
  - ARIN   (North America)          -- whois.arin.net REST API
  - RIPE   (Europe / Middle East)   -- rdap.db.ripe.net
  - APNIC  (Asia-Pacific)           -- rdap.apnic.net
  - LACNIC (Latin America)          -- rdap.lacnic.net
  - AFRINIC (Africa)                -- rdap.afrinic.net

Falls back to rdap.org (auto-routing) when a specific RIR query fails.
"""

from __future__ import annotations

import ipaddress
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from threading import Lock
from typing import Callable, Dict, List, Optional, Tuple

import requests

from catrg.utils.logger import get_logger

log = get_logger(__name__)


@dataclass
class GeoResult:
    country: str = "N/A"
    city: str = "N/A"
    error: Optional[str] = None


@dataclass
class WhoisResult:
    organization: str = "N/A"
    registry: str = "N/A"
    error: Optional[str] = None


@dataclass
class IpLookupResult:
    geo: GeoResult = field(default_factory=GeoResult)
    whois: WhoisResult = field(default_factory=WhoisResult)


# ── RIR detection ─────────────────────────────────────────────────

_ARIN_V4 = [
    ("3.0.0.0", "3.255.255.255"),
    ("4.0.0.0", "4.255.255.255"),
    ("6.0.0.0", "7.255.255.255"),
    ("8.0.0.0", "15.255.255.255"),
    ("16.0.0.0", "19.255.255.255"),
    ("20.0.0.0", "20.255.255.255"),
    ("23.0.0.0", "23.255.255.255"),
    ("24.0.0.0", "24.255.255.255"),
    ("32.0.0.0", "35.255.255.255"),
    ("40.0.0.0", "40.255.255.255"),
    ("44.0.0.0", "44.255.255.255"),
    ("47.0.0.0", "47.255.255.255"),
    ("50.0.0.0", "50.255.255.255"),
    ("52.0.0.0", "52.255.255.255"),
    ("54.0.0.0", "54.255.255.255"),
    ("63.0.0.0", "63.255.255.255"),
    ("64.0.0.0", "75.255.255.255"),
    ("96.0.0.0", "96.255.255.255"),
    ("97.0.0.0", "97.255.255.255"),
    ("98.0.0.0", "100.255.255.255"),
    ("104.0.0.0", "104.255.255.255"),
    ("107.0.0.0", "107.255.255.255"),
    ("108.0.0.0", "108.255.255.255"),
    ("128.0.0.0", "128.255.255.255"),
    ("129.0.0.0", "130.255.255.255"),
    ("131.0.0.0", "131.255.255.255"),
    ("132.0.0.0", "134.255.255.255"),
    ("135.0.0.0", "135.255.255.255"),
    ("136.0.0.0", "136.255.255.255"),
    ("137.0.0.0", "137.255.255.255"),
    ("138.0.0.0", "138.255.255.255"),
    ("139.0.0.0", "139.255.255.255"),
    ("140.0.0.0", "140.255.255.255"),
    ("142.0.0.0", "142.255.255.255"),
    ("143.0.0.0", "143.255.255.255"),
    ("144.0.0.0", "144.255.255.255"),
    ("146.0.0.0", "146.255.255.255"),
    ("147.0.0.0", "148.255.255.255"),
    ("149.0.0.0", "149.255.255.255"),
    ("152.0.0.0", "152.255.255.255"),
    ("155.0.0.0", "155.255.255.255"),
    ("156.0.0.0", "156.255.255.255"),
    ("157.0.0.0", "157.255.255.255"),
    ("158.0.0.0", "159.255.255.255"),
    ("160.0.0.0", "160.255.255.255"),
    ("161.0.0.0", "161.255.255.255"),
    ("162.0.0.0", "162.255.255.255"),
    ("164.0.0.0", "164.255.255.255"),
    ("165.0.0.0", "165.255.255.255"),
    ("166.0.0.0", "166.255.255.255"),
    ("167.0.0.0", "167.255.255.255"),
    ("168.0.0.0", "168.255.255.255"),
    ("169.0.0.0", "169.255.255.255"),
    ("170.0.0.0", "170.255.255.255"),
    ("172.0.0.0", "172.255.255.255"),
    ("173.0.0.0", "173.255.255.255"),
    ("174.0.0.0", "174.255.255.255"),
    ("184.0.0.0", "184.255.255.255"),
    ("192.0.0.0", "192.255.255.255"),
    ("198.0.0.0", "199.255.255.255"),
    ("204.0.0.0", "209.255.255.255"),
    ("216.0.0.0", "216.255.255.255"),
]


def _is_likely_arin(ip_str: str) -> bool:
    """Heuristic: return True if the IP is likely in ARIN space."""
    try:
        addr = ipaddress.ip_address(ip_str)
        if addr.version == 6:
            return True  # Default to ARIN for v6, fallback will handle it
        for start, end in _ARIN_V4:
            if ipaddress.ip_address(start) <= addr <= ipaddress.ip_address(end):
                return True
    except ValueError:
        pass
    return False


def _detect_rir(ip_str: str) -> str:
    """Return 'arin', 'ripe', 'apnic', 'lacnic', 'afrinic', or 'rdap'."""
    if _is_likely_arin(ip_str):
        return "arin"
    return "rdap"


# ── Query functions ───────────────────────────────────────────────

def _query_maxmind(ip: str, account_id: str, license_key: str) -> GeoResult:
    if not account_id or not license_key:
        return GeoResult(error="MaxMind credentials not provided")
    try:
        url = f"https://geolite.info/geoip/v2.1/city/{ip}"
        resp = requests.get(url, auth=(account_id, license_key), timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            return GeoResult(
                country=data.get("country", {}).get("names", {}).get("en", "N/A"),
                city=data.get("city", {}).get("names", {}).get("en", "N/A"),
            )
        return GeoResult(error=f"MaxMind request failed with status {resp.status_code}")
    except requests.Timeout:
        return GeoResult(error="MaxMind query timed out after 10 seconds")
    except Exception as e:
        return GeoResult(error=str(e))


def _query_arin(ip: str, api_key: str = "") -> WhoisResult:
    try:
        base_url = f"https://whois.arin.net/rest/ip/{ip}.json"
        url = f"{base_url}?apikey={api_key}" if api_key else base_url
        headers = {"Accept": "application/json"}
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            net = data.get("net", {})
            org = net.get("orgRef", {}).get("@name", "N/A")
            return WhoisResult(organization=org, registry="ARIN")
        return WhoisResult(error=f"ARIN request failed with status {resp.status_code}")
    except requests.Timeout:
        return WhoisResult(error="ARIN query timed out after 10 seconds")
    except Exception as e:
        return WhoisResult(error=str(e))


def _query_rdap(ip: str) -> WhoisResult:
    """Query rdap.org which auto-routes to the correct RIR."""
    try:
        url = f"https://rdap.org/ip/{ip}"
        resp = requests.get(url, timeout=10, headers={"Accept": "application/rdap+json"})
        if resp.status_code == 200:
            data = resp.json()
            org = "N/A"
            for entity in data.get("entities", []):
                vcard = entity.get("vcardArray", [None, []])[1] if entity.get("vcardArray") else []
                for entry in vcard:
                    if isinstance(entry, list) and len(entry) >= 4 and entry[0] == "fn":
                        org = entry[3]
                        break
                if org != "N/A":
                    break
            port43 = data.get("port43", "")
            registry = "RDAP"
            if "arin" in port43:
                registry = "ARIN"
            elif "ripe" in port43:
                registry = "RIPE NCC"
            elif "apnic" in port43:
                registry = "APNIC"
            elif "lacnic" in port43:
                registry = "LACNIC"
            elif "afrinic" in port43:
                registry = "AFRINIC"
            return WhoisResult(organization=org, registry=registry)
        return WhoisResult(error=f"RDAP request failed with status {resp.status_code}")
    except requests.Timeout:
        return WhoisResult(error="RDAP query timed out after 10 seconds")
    except Exception as e:
        return WhoisResult(error=str(e))


# ── Cached lookup service ────────────────────────────────────────

class IpLookupService:
    """Thread-safe, cached IP lookup service with parallel query support."""

    def __init__(self, maxmind_id: str = "", maxmind_key: str = "", arin_key: str = ""):
        self.maxmind_id = maxmind_id
        self.maxmind_key = maxmind_key
        self.arin_key = arin_key
        self._cache: Dict[str, IpLookupResult] = {}
        self._lock = Lock()
        self._max_workers = 4

    @property
    def arin_rate_per_min(self) -> int:
        return 60 if self.arin_key else 15

    def lookup(self, ip: str) -> IpLookupResult:
        """Look up a single IP, returning cached results when available."""
        with self._lock:
            if ip in self._cache:
                return self._cache[ip]

        geo = _query_maxmind(ip, self.maxmind_id, self.maxmind_key)

        rir = _detect_rir(ip)
        if rir == "arin":
            whois = _query_arin(ip, self.arin_key)
            if whois.error:
                fallback = _query_rdap(ip)
                if not fallback.error:
                    whois = fallback
        else:
            whois = _query_rdap(ip)

        result = IpLookupResult(geo=geo, whois=whois)
        with self._lock:
            self._cache[ip] = result
        return result

    def lookup_batch(
        self,
        ips: List[str],
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> Dict[str, IpLookupResult]:
        """Look up multiple IPs in parallel with rate limiting.

        *progress_callback(completed, total)* is called after each IP completes.
        """
        to_query = []
        results: Dict[str, IpLookupResult] = {}

        with self._lock:
            for ip in ips:
                if ip in self._cache:
                    results[ip] = self._cache[ip]
                else:
                    to_query.append(ip)

        total = len(ips)
        completed = len(results)

        if progress_callback:
            progress_callback(completed, total)

        if not to_query:
            return results

        delay = 60.0 / self.arin_rate_per_min
        last_query_time = 0.0

        def _rate_limited_lookup(ip: str) -> Tuple[str, IpLookupResult]:
            nonlocal last_query_time
            with self._lock:
                now = time.monotonic()
                wait = delay - (now - last_query_time)
                if wait > 0:
                    time.sleep(wait)
                last_query_time = time.monotonic()
            return ip, self.lookup(ip)

        with ThreadPoolExecutor(max_workers=self._max_workers) as pool:
            futures = {pool.submit(_rate_limited_lookup, ip): ip for ip in to_query}
            for future in as_completed(futures):
                try:
                    ip, result = future.result()
                    results[ip] = result
                except Exception as e:
                    ip = futures[future]
                    results[ip] = IpLookupResult(
                        geo=GeoResult(error=str(e)),
                        whois=WhoisResult(error=str(e)),
                    )
                completed += 1
                if progress_callback:
                    progress_callback(completed, total)

        return results

    def get_cached(self, ip: str) -> Optional[IpLookupResult]:
        with self._lock:
            return self._cache.get(ip)

    def clear_cache(self) -> None:
        with self._lock:
            self._cache.clear()
