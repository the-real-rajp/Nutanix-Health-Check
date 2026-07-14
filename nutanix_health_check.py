#!/usr/bin/env python3
"""
Nutanix Health Check Script - APIv4  (Prism Central edition)
=============================================================
Connects to Prism Central, discovers all registered clusters,
and generates a separate health-check report (.docx) for each one.

Usage (interactive – recommended):
    python nutanix_health_check.py

Usage (non-interactive / scripted):
    python nutanix_health_check.py --host pc.example.com --user admin
    python nutanix_health_check.py --from-json cluster_raw.json --customer "Acme"

Requirements:
    pip install requests urllib3 matplotlib
    node  (Node.js 18+ must be installed; docx npm package is auto-installed locally on first run)

    matplotlib is optional but recommended — it generates the 7-day CPU usage chart in the report.
    Without it the chart section is skipped and a text note is shown instead.
"""

import argparse
import atexit
import csv
import getpass
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib3
from datetime import datetime, timedelta, timezone
from typing import Any, Optional

import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


OS_COMPAT_CSV_FILENAMES = [
    "OS_Compatibility_Matrix.csv",
    "OS Compatibility Matrix.csv",
    "os_compatibility_matrix.csv",
]

AOS_EOL_CSV_FILENAMES = [
    "NOS_EOL_information_list.csv",
    "NOS EOL information list.csv",
    "nos_eol_information_list.csv",
]


def _support_file_candidate_paths(filenames: list, explicit_path: str = "") -> list:
    """Build a deduplicated list of candidate support-file paths."""
    candidates = []
    if explicit_path:
        candidates.append(explicit_path)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    search_dirs = [
        os.getcwd(),
        os.path.join(os.getcwd(), "data"),
        script_dir,
        os.path.join(script_dir, "data"),
        "/mnt/data",
    ]
    for folder in search_dirs:
        for filename in filenames:
            candidates.append(os.path.join(folder, filename))

    return list(dict.fromkeys([c for c in candidates if c]))


def _find_optional_file(filename, explicit_path: str = "") -> str:
    """Return an optional file path if it exists near the script/current working directory."""
    filenames = filename if isinstance(filename, list) else [filename]
    for candidate in _support_file_candidate_paths(filenames, explicit_path):
        if os.path.exists(candidate):
            return candidate
    return ""


def _find_required_file(label: str, filename, explicit_path: str = "") -> str:
    """Return a required supporting file path or raise a clear, actionable error."""
    filenames = filename if isinstance(filename, list) else [filename]
    found = _find_optional_file(filenames, explicit_path)
    if found:
        return found

    expected = "\n  - ".join(filenames)
    raise FileNotFoundError(
        f"Missing required {label}.\n"
        f"Expected one of:\n  - {expected}"
    )


def _parse_iso_date(value: str) -> Optional[datetime]:
    """Parse Nutanix CSV ISO timestamp values into timezone-aware datetimes."""
    if not value:
        return None
    try:
        clean = value.strip().replace("Z", "+00:00")
        return datetime.fromisoformat(clean)
    except Exception:
        return None


def _version_tuple(value: str) -> tuple:
    nums = re.findall(r"\d+", str(value or ""))
    return tuple(int(n) for n in nums)


def _aos_minor_version(value: str) -> str:
    """Return the AOS minor train (for example, 7.5 from 7.5.1.6)."""
    m = re.search(r"(\d+\.\d+)", str(value or ""))
    return m.group(1) if m else str(value or "").strip()


def _normalise_os_name(value: str) -> str:
    """Normalise OS names for compatibility-matrix matching."""
    s = str(value or "").strip().lower()
    s = s.replace("red hat enterprise linux", "rhel")
    s = s.replace("rockylinux", "rocky linux")
    s = s.replace("amazonlinux", "amazon linux")
    s = s.replace("windowsserver", "windows server")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _load_os_compatibility_matrix(path: str, aos_version: str) -> dict:
    """Load OS compatibility entries for the current AOS minor train."""
    matrix = {}
    if not path or not os.path.exists(path):
        return matrix
    aos_minor = _aos_minor_version(aos_version)
    target = f"AOS {aos_minor}".strip()
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        header_idx = None
        for i, row in enumerate(rows):
            if row and row[0].strip() == "AOS Version":
                header_idx = i
                break
        if header_idx is None:
            return matrix
        headers = [h.strip() for h in rows[header_idx]]
        idx = {h: n for n, h in enumerate(headers)}
        for row in rows[header_idx + 1:]:
            if len(row) <= max(idx.get("AOS Version", 0), idx.get("OS Family", 0), idx.get("OS Status", 0)):
                continue
            if row[idx["AOS Version"]].strip() != target:
                continue
            os_family = row[idx["OS Family"]].strip()
            os_status = row[idx["OS Status"]].strip() or "-"
            if not os_family:
                continue
            report_status = "Legacy Support" if os_status.lower().startswith("legacy") else "Supported"
            matrix[_normalise_os_name(os_family)] = {
                "os_family": os_family,
                "csv_status": os_status,
                "support": report_status,
            }
    except Exception as exc:
        print(f"    [WARN] Could not read OS compatibility matrix {path}: {exc}")
    return matrix


def _lookup_os_support(os_name: str, matrix: dict) -> dict:
    """Return OS support classification for a VM operating system."""
    norm = _normalise_os_name(os_name)
    if not norm or norm == "unknown":
        return {"support": "Unsupported", "matched_os": "N/A"}
    if norm in matrix:
        return {"support": matrix[norm]["support"], "matched_os": matrix[norm]["os_family"]}
    # Prefix matching handles strings like "Rocky Linux 10.1" matching "Rocky Linux 10".
    best = None
    for key, entry in matrix.items():
        if norm == key or norm.startswith(key + " ") or key.startswith(norm + " "):
            if best is None or len(key) > len(best[0]):
                best = (key, entry)
    if best:
        entry = best[1]
        return {"support": entry["support"], "matched_os": entry["os_family"]}
    return {"support": "Unsupported", "matched_os": "N/A"}


def _load_aos_eol_info(path: str, aos_version: str) -> dict:
    """Load lifecycle/EOL information for the current AOS version/train."""
    result = {
        "current_version": aos_version or "N/A",
        "matched_version": "N/A",
        "latest_version": "N/A",
        "end_of_maintenance": "N/A",
        "end_of_support_life": "N/A",
        "lifecycle_status": "Unknown",
        "report_status": "Recommended",
        "note": "AOS lifecycle data was not available.",
    }
    if not path or not os.path.exists(path):
        return result
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        header_idx = None
        for i, row in enumerate(rows):
            if row and row[0].strip() == "AOS Version":
                header_idx = i
                break
        if header_idx is None:
            return result
        headers = [h.strip() for h in rows[header_idx]]
        idx = {h: n for n, h in enumerate(headers)}
        entries = []
        for row in rows[header_idx + 1:]:
            if not row or not row[0].strip():
                continue
            def get(col):
                pos = idx.get(col)
                return row[pos].strip() if pos is not None and pos < len(row) else ""
            entries.append({
                "version": get("AOS Version"),
                "ga_date": get("GA Date"),
                "eom": get("End of Maintenance"),
                "eos": get("End of Support Life"),
                "latest": get("Latest Version"),
                "note": get("Note"),
            })
        current = str(aos_version or "").strip()
        current_minor = _aos_minor_version(current)
        exact = [e for e in entries if e["version"] == current]
        train = [e for e in entries if _aos_minor_version(e["version"]) == current_minor]
        match = exact[0] if exact else (train[0] if train else None)
        if not match:
            result["note"] = f"AOS {current} was not found in the lifecycle matrix."
            result["lifecycle_status"] = "Unsupported / Unknown"
            result["report_status"] = "Critical"
            return result
        latest = match.get("latest") or ""
        if not latest and train:
            versions = [e["version"] for e in train if e.get("version")]
            latest = sorted(versions, key=_version_tuple)[-1] if versions else "N/A"
        eom_dt = _parse_iso_date(match.get("eom", ""))
        eos_dt = _parse_iso_date(match.get("eos", ""))
        now = datetime.now(timezone.utc)
        if eos_dt and now > eos_dt:
            lifecycle = "Past End of Support Life"
            report_status = "Critical"
        elif eom_dt and now > eom_dt:
            lifecycle = "Past End of Maintenance"
            report_status = "Recommended"
        elif latest and latest != "N/A" and _version_tuple(current) < _version_tuple(latest):
            lifecycle = "Supported - update available"
            report_status = "Recommended"
        else:
            lifecycle = "Supported"
            report_status = "Healthy"
        result.update({
            "matched_version": match.get("version") or "N/A",
            "latest_version": latest or "N/A",
            "end_of_maintenance": (eom_dt.date().isoformat() if eom_dt else "N/A"),
            "end_of_support_life": (eos_dt.date().isoformat() if eos_dt else "N/A"),
            "lifecycle_status": lifecycle,
            "report_status": report_status,
            "note": match.get("note") or "",
        })
    except Exception as exc:
        print(f"    [WARN] Could not read AOS lifecycle matrix {path}: {exc}")
    return result


# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------

BANNER = r"""
  _   _       _              _        _   _            _ _   _     
 | \ | |_   _| |_ __ _ _ _ (_)_  __ | | | | ___  __ _| | |_| |__  
 |  \| | | | | __/ _` | ' \| \ \/ / | |_| |/ _ \/ _` | | __| '_ \ 
 | |\  | |_| | || (_| | | | | |>  <  |  _  |  __/ (_| | | |_| | | |
 |_| \_|\__,_|\__\__,_|_| |_|_/_/\_\ |_| |_|\___|\__,_|_|\__|_| |_|
                                                                     
        Cluster Health Check  |  APIv4  |  Prism Central Edition
"""


# ---------------------------------------------------------------------------
# Interactive prompts
# ---------------------------------------------------------------------------

def prompt_connection() -> tuple[str, int, str, str, str]:
    """
    Ask the user for PC host, port, username, password, and customer name.
    Returns (host, port, username, password, customer).
    """
    print(BANNER)
    print("=" * 65)
    print("  Prism Central Connection Setup")
    print("=" * 65)
    print()

    # Host
    while True:
        host = input("  Prism Central IP or FQDN: ").strip()
        if host:
            break
        print("  [!] Host cannot be empty. Please try again.")

    # Port
    port_raw = input("  API port [9440]: ").strip()
    port = int(port_raw) if port_raw.isdigit() else 9440

    # Username
    while True:
        username = input("  Username: ").strip()
        if username:
            break
        print("  [!] Username cannot be empty.")

    # Password (hidden input)
    while True:
        password = getpass.getpass("  Password: ")
        if password:
            break
        print("  [!] Password cannot be empty.")

    # Customer name for the report
    customer = input("  Customer / Organization name [CUSTOMER_NAME]: ").strip()
    if not customer:
        customer = "CUSTOMER_NAME"

    print()
    return host, port, username, password, customer


# ---------------------------------------------------------------------------
# API Client  (Prism Central)
# ---------------------------------------------------------------------------

class PrismCentralClient:
    """
    APIv4 client scoped to a Prism Central instance.
    All cluster-level calls pass the target cluster's extId as a path
    segment where required by the v4 API.
    """

    def __init__(self, host: str, username: str, password: str, port: int = 9440):
        self.host     = host
        self.port     = port
        self.base     = f"https://{host}:{port}/api"
        self.session  = requests.Session()
        self.session.auth   = (username, password)
        self.session.verify = False
        self.session.headers.update({
            "Content-Type": "application/json",
            "Accept":       "application/json",
        })

    # ── low-level helpers ────────────────────────────────────────────────

    def get(self, path: str, params: Optional[dict] = None) -> Any:
        url  = f"{self.base}{path}"
        last_resp = None
        for attempt in range(3):
            resp = self.session.get(url, params=params, timeout=60)
            last_resp = resp
            if resp.status_code == 429 and attempt < 2:
                time.sleep(2 + attempt * 3)
                continue
            resp.raise_for_status()
            return resp.json()
        last_resp.raise_for_status()
        return last_resp.json()

    def post(self, path: str, body: Optional[dict] = None, params: Optional[dict] = None) -> Any:
        url  = f"{self.base}{path}"
        last_resp = None
        for attempt in range(3):
            resp = self.session.post(url, json=(body or {}), params=params, timeout=60)
            last_resp = resp
            if resp.status_code == 429 and attempt < 2:
                time.sleep(2 + attempt * 3)
                continue
            resp.raise_for_status()
            return resp.json()
        last_resp.raise_for_status()
        return last_resp.json()

    def paginate(self, path: str, limit: int = 100,
                 extra_params: Optional[dict] = None) -> list:
        items, offset = [], 0
        while True:
            params = {"$limit": limit, "$offset": offset}
            if extra_params:
                params.update(extra_params)
            data = self.get(path, params=params)
            page = data.get("data", [])
            if not page:
                break
            items.extend(page)
            if len(page) < limit:
                break
            offset += limit
        return items

    # ── PC-level: discover registered clusters ───────────────────────────

    def list_clusters(self) -> list[dict]:
        """
        Returns all clusters registered to this Prism Central.
        PC itself appears as a cluster of type PRISM_CENTRAL — we skip it
        and return only AOS / AHV clusters.
        """
        raw = self.paginate("/clustermgmt/v4.0/config/clusters")
        clusters = []
        for c in raw:
            cluster_type = (
                c.get("config", {}).get("clusterFunction", []) or
                c.get("clusterFunction", []) or
                []
            )
            # Skip Prism Central management cluster entries
            if "PRISM_CENTRAL" in [str(f).upper() for f in cluster_type]:
                continue
            clusters.append(c)
        return clusters

    def detect_api_version(self, namespace: str = "clustermgmt") -> str:
        """
        Detect the real API version by probing common versions.
        Returns the highest working version string, e.g. 'v4.2'.
        Falls back to 'v4.0'.
        """
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            try:
                self.get(f"/{namespace}/{ver}/config/clusters", {"$limit": 1})
                return ver
            except Exception:
                continue
        return "v4.0"

    def detect_api_version_from_link(self, link: str) -> str:
        """Extract API version from a self-link href, e.g. '.../v4.2/config/...' -> 'v4.2'."""
        import re as _re
        m = _re.search(r"/(v[\d]+\.[\d]+(?:\.[a-z0-9]+)?)/", link)
        return m.group(1) if m else "v4.0"

    def test_connection(self) -> bool:
        try:
            self.get("/clustermgmt/v4.0/config/clusters", {"$limit": 1})
            return True
        except Exception:
            return False


# ---------------------------------------------------------------------------
# Per-Cluster Data Collector
# ---------------------------------------------------------------------------

class ClusterDataCollector:
    """
    Collects all health-check data for ONE cluster via Prism Central.
    The cluster extId is used to scope API calls where the v4 API requires it.
    """

    def __init__(self, client: PrismCentralClient, cluster_ext_id: str,
                 cluster_name: str):
        self.client          = client
        self.cluster_ext_id  = cluster_ext_id
        self.cluster_name    = cluster_name
        self.data            = {}
        self._api_ver        = None   # lazily detected from self-links

    def _api_version(self) -> str:
        """
        Return the detected API version for this PC.
        Reads the self-link from the first storage container if available,
        otherwise probes the PC for the highest working version.
        """
        if self._api_ver:
            return self._api_ver
        # Try to read from already-collected storage container links
        containers = self.data.get("storage_containers", [])
        if isinstance(containers, list) and containers:
            for link in containers[0].get("links", []):
                href = link.get("href", "")
                if href:
                    ver = self.client.detect_api_version_from_link(href)
                    if ver != "v4.0" or "v4.0" in href:
                        self._api_ver = ver
                        return self._api_ver
        # Fall back to probing
        self._api_ver = self.client.detect_api_version("clustermgmt")
        print(f"      Detected API version: {self._api_ver}")
        return self._api_ver

    def _safe_get(self, path: str, params: Optional[dict] = None) -> Any:
        try:
            return self.client.get(path, params)
        except Exception as exc:
            raise exc

    def _safe_paginate(self, path: str, extra_params: Optional[dict] = None) -> list:
        try:
            return self.client.paginate(path, extra_params=extra_params)
        except Exception:
            return []

    def collect_all(self) -> dict:
        cid = self.cluster_ext_id
        steps = [
            ("cluster_info",        self._cluster_info),
            ("security_hardening",  self._security_hardening),
            ("storage_containers",  self._storage_containers),  # collect early to detect API ver
            ("container_stats",     self._container_stats),     # per-container detail stats
            ("nodes",               self._nodes),
            ("virtual_machines",    self._virtual_machines),    # collect before ahv_version
            ("cvm_virtual_machines", self._cvm_virtual_machines), # optional PE CVM VM details
            ("ahv_version",         self._ahv_version),         # uses VM host UUIDs as fallback
            ("alerts",              self._alerts),
            ("protection_policies", self._protection_policies),
            ("recovery_plans",      self._recovery_plans),
            ("storage_stats",       self._storage_stats),
            ("cluster_stats",       self._cluster_stats),
            ("cluster_stats_7d",    self._cluster_stats_7d),
            ("networks",            self._networks),
            ("network_details",     self._network_details),
            ("licensing",           self._licensing),
            ("ncc_checks",          self._ncc_checks),
        ]
        for name, fn in steps:
            print(f"      • {name}...")
            try:
                self.data[name] = fn()
            except Exception as exc:
                print(f"        WARNING: {exc}")
                self.data[name] = None
        return self.data

    # ── collection methods ───────────────────────────────────────────────

    def _cluster_info(self):
        # Try multiple API versions for the cluster detail endpoint
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            try:
                result = self.client.get(
                    f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}"
                )
                if result:
                    return result
            except Exception:
                continue
        # Fall back to list and filter
        for ver in ["v4.2", "v4.0"]:
            try:
                raw = self.client.paginate(f"/clustermgmt/{ver}/config/clusters")
                for c in raw:
                    if c.get("extId") == self.cluster_ext_id:
                        return {"data": c}
            except Exception:
                continue
        return {}

    def _security_hardening(self):
        """Collect cluster security-hardening state from PC and PE.

        The PC Security Summary API exposes the dashboard-level controls such
        as cluster lockdown, log forwarding, consent banner, and host Secure
        Boot.  PE v2 cluster inventory supplies the CVM/AHV consent-banner
        settings when they are available.
        Both sources are optional because Security Dashboard availability and
        legacy PE field coverage vary by software version and licensing.
        """
        result = {}

        # Security Dashboard API. Try the current version first and retain the
        # cluster-specific record only.
        for ver in ["v4.1", "v4.0"]:
            path = f"/security/{ver}/report/security-summaries"
            attempts = [
                {"$filter": f"clusterExtId eq '{self.cluster_ext_id}'", "$limit": 100},
                {"$limit": 100},
            ]
            for params in attempts:
                try:
                    response = self.client.get(path, params)
                    records = response.get("data", []) if isinstance(response, dict) else []
                    if isinstance(records, dict):
                        records = [records]
                    match = next(
                        (r for r in records if isinstance(r, dict) and r.get("clusterExtId") == self.cluster_ext_id),
                        None,
                    )
                    if match:
                        result["pc_security_summary"] = match
                        break
                except Exception:
                    continue
            if result.get("pc_security_summary"):
                break

        # PE v2 returns security_compliance_config and
        # hypervisor_security_compliance_config on the cluster object.
        pe_cluster = self._pe_v2_get("/cluster/") or self._pe_v2_get("/cluster")
        if isinstance(pe_cluster, dict) and pe_cluster:
            result["pe_cluster"] = pe_cluster

        return result

    def _nodes(self):
        """Fetch hosts/host-nodes using all known API version + path patterns."""
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            # Cluster-scoped sub-resource. The cluster self-link commonly exposes:
            # /clustermgmt/<ver>/config/clusters/<cluster_ext_id>/hosts
            for sub in ["hosts", "host-nodes", "nodes"]:
                try:
                    result = self.client.paginate(
                        f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/{sub}"
                    )
                    if result:
                        return result
                except Exception:
                    continue

            # Global lists with common cluster filter field names.
            for endpoint in ["hosts", "host-nodes"]:
                for filt in [
                    f"clusterExtId eq '{self.cluster_ext_id}'",
                    f"cluster/extId eq '{self.cluster_ext_id}'",
                    f"clusterReference eq '{self.cluster_ext_id}'",
                ]:
                    try:
                        result = self.client.paginate(
                            f"/clustermgmt/{ver}/config/{endpoint}",
                            extra_params={"$filter": filt}
                        )
                        if result:
                            return result
                    except Exception:
                        continue
        return []

    def _ahv_version(self):
        """
        Return the cluster-wide AHV version from Prism Central.

        The previous code only tried /clustermgmt/v4.0/config/hosts with one
        filter shape. In many PC builds, the filter property is clusterExtId,
        and in others the AHV value is on host-nodes rather than hosts. This
        version tries all common PC v4 paths/filter shapes and falls back to
        Prism Element v2 using the cluster VIP.
        """

        def add_version(value, versions: set[str]) -> None:
            if not value:
                return
            text = str(value).strip()
            if not text or text.upper() in {"N/A", "NONE", "UNKNOWN"}:
                return
            # Examples: "AHV 11.0.1.2", "Nutanix AHV 11.0.1", "11.0.1.2"
            m = re.search(r"(\d+\.\d+(?:\.\d+){0,3})", text)
            versions.add(m.group(1) if m else text)

        def walk(obj):
            """Yield every dict recursively so we can survive small API-shape changes."""
            if isinstance(obj, dict):
                yield obj
                for v in obj.values():
                    yield from walk(v)
            elif isinstance(obj, list):
                for item in obj:
                    yield from walk(item)

        def extract_versions(hosts) -> set[str]:
            versions: set[str] = set()
            if isinstance(hosts, dict):
                hosts = hosts.get("data", hosts)
            if isinstance(hosts, dict):
                hosts = [hosts]
            if not isinstance(hosts, list):
                return versions

            for host in hosts:
                if not isinstance(host, dict):
                    continue

                # Known Prism Central v4 / host-node shapes.
                for d in walk(host):
                    for key in (
                        "fullName", "hypervisorFullName", "hypervisor_full_name",
                        "version", "hypervisorVersion", "hypervisor_version",
                        "hypervisorFullVersion", "hypervisor_full_version",
                    ):
                        val = d.get(key)
                        if val and ("AHV" in str(val).upper() or re.search(r"\d+\.\d+", str(val))):
                            add_version(val, versions)

                # Avoid accidentally returning AOS/NCC versions from nested software maps.
                versions = {v for v in versions if not str(v).lower().startswith("ncc-")}
            return versions

        versions: set[str] = set()
        versions_to_try = ["v4.2", "v4.1", "v4.0.b1", "v4.0"]

        # 1) Prism Central host inventory endpoint. Try multiple filter names.
        host_filters = [
            f"clusterExtId eq '{self.cluster_ext_id}'",
            f"cluster/extId eq '{self.cluster_ext_id}'",
            f"clusterReference eq '{self.cluster_ext_id}'",
        ]
        for ver in versions_to_try:
            for filt in host_filters:
                try:
                    resp = self.client.get(
                        f"/clustermgmt/{ver}/config/hosts",
                        params={"$filter": filt, "$limit": 100},
                    )
                    versions.update(extract_versions(resp))
                    if versions:
                        return ", ".join(sorted(versions))
                except Exception:
                    continue

        # 2) Prism Central cluster-scoped host inventory endpoint.
        # Your raw cluster_info self-link advertises this exact relation:
        # /clustermgmt/v4.2/config/clusters/<cluster_uuid>/hosts
        node_filters = [
            f"clusterExtId eq '{self.cluster_ext_id}'",
            f"cluster/extId eq '{self.cluster_ext_id}'",
            f"clusterReference eq '{self.cluster_ext_id}'",
        ]
        for ver in versions_to_try:
            # Cluster-scoped subresources. Include hosts first; this was missing.
            for sub in ("hosts", "host-nodes", "nodes"):
                try:
                    hosts = self.client.paginate(
                        f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/{sub}"
                    )
                    versions.update(extract_versions(hosts))
                    if versions:
                        return ", ".join(sorted(versions))
                except Exception:
                    continue

            # Global host and host-node lists with filters.
            for endpoint in ("hosts", "host-nodes"):
                for filt in node_filters:
                    try:
                        hosts = self.client.paginate(
                            f"/clustermgmt/{ver}/config/{endpoint}",
                            extra_params={"$filter": filt},
                        )
                        versions.update(extract_versions(hosts))
                        if versions:
                            return ", ".join(sorted(versions))
                    except Exception:
                        continue

        # 3) Use nodes already collected earlier in collect_all().
        versions.update(extract_versions(self.data.get("nodes", [])))
        if versions:
            return ", ".join(sorted(versions))

        # 4) Final fallback: call Prism Element v2 /hosts through the cluster VIP.
        pe_hosts = self._pe_v2_get("/hosts") or self._pe_v2_get("/hosts/")
        if pe_hosts:
            entities = pe_hosts.get("entities", pe_hosts)
            versions.update(extract_versions(entities))
            if versions:
                return ", ".join(sorted(versions))

        print("        AHV version not found. Try this endpoint manually:")
        print(f"        GET https://{self.client.host}:{self.client.port}/api/clustermgmt/v4.2/config/clusters/{self.cluster_ext_id}/hosts")
        return None

    def _virtual_machines(self):
        """
        Fetch VMs for this cluster. Tries multiple filter field names
        because PC API versions differ on the exact OData filter path.
        Falls back to fetching all VMs and filtering by cluster UUID client-side.
        """
        for filt in [
            f"cluster/extId eq '{self.cluster_ext_id}'",
            f"clusterExtId eq '{self.cluster_ext_id}'",
            f"clusterUuid eq '{self.cluster_ext_id}'",
        ]:
            try:
                result = self.client.paginate(
                    "/vmm/v4.0/ahv/config/vms",
                    extra_params={"$filter": filt}
                )
                if result:
                    return result
            except Exception:
                continue

        # Final fallback: fetch all VMs and filter by cluster UUID client-side
        try:
            all_vms = self.client.paginate("/vmm/v4.0/ahv/config/vms")
            return [
                vm for vm in all_vms
                if (vm.get("cluster", {}).get("extId") == self.cluster_ext_id or
                    vm.get("clusterExtId") == self.cluster_ext_id or
                    vm.get("clusterUuid") == self.cluster_ext_id)
            ]
        except Exception:
            return []

    def _cvm_virtual_machines(self):
        """
        Collect CVM configuration from the authoritative Cluster Management
        CVM API. The list endpoint is queried first, then each CVM is queried
        individually when an extId is available so CPU and memory fields are
        not lost in abbreviated list responses.

        Falls back to the earlier PE/v3 discovery logic when the Cluster
        Management CVM endpoint is unavailable.
        """
        candidates = []
        seen = set()

        def add_candidate(obj):
            if not isinstance(obj, dict):
                return
            key = obj.get("extId") or obj.get("uuid") or obj.get("name") or json.dumps(obj, sort_keys=True, default=str)
            if key in seen:
                return
            seen.add(key)
            candidates.append(obj)

        # Preferred source: /clusters/{clusterExtId}/cvms and the per-CVM
        # detail endpoint documented in clustermgmt v4.2.
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            list_path = f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/cvms"
            try:
                cvm_list = self.client.paginate(list_path)
            except Exception:
                cvm_list = []

            if not cvm_list:
                continue

            for summary in cvm_list:
                if not isinstance(summary, dict):
                    continue
                ext_id = summary.get("extId") or summary.get("uuid")
                detail = None
                if ext_id:
                    try:
                        response = self.client.get(f"{list_path}/{ext_id}")
                        if isinstance(response, dict):
                            detail = response.get("data", response)
                    except Exception:
                        detail = None

                # Preserve useful fields from both responses. Detail wins, but
                # list-only fields remain available if the detail response is
                # sparse on a particular release.
                merged = dict(summary)
                if isinstance(detail, dict):
                    merged.update(detail)
                add_candidate(merged)

            if candidates:
                return candidates

        # Fallback 1: common Prism Element VM endpoints.
        for path in [
            "/vms?include_vm_disk_config=false&include_vm_nic_config=true",
            "/vms/?include_vm_disk_config=false&include_vm_nic_config=true",
            "/vms",
            "/vms/",
        ]:
            data = self._pe_v2_get(path)
            if not data:
                continue
            entities = data.get("entities", data if isinstance(data, list) else [])
            if isinstance(entities, list):
                for entity in entities:
                    add_candidate(entity)

        # Fallback 2: v3 VM inventory.
        try:
            v3 = self.client.post(
                "/nutanix/v3/vms/list",
                body={"kind": "vm", "length": 500, "offset": 0},
            )
            entities = v3.get("entities", []) if isinstance(v3, dict) else []
            if isinstance(entities, list):
                for entity in entities:
                    add_candidate(entity)
        except Exception:
            pass

        # Keep only likely CVMs for fallback sources. The authoritative
        # clustermgmt path returned above does not need this filtering.
        cvm_ips = set()
        for n in self._safe_list("nodes"):
            ip = (n.get("controllerVm", {})
                    .get("externalAddress", {})
                    .get("ipv4", {})
                    .get("value"))
            if ip:
                cvm_ips.add(str(ip))

        out = []
        for vm in candidates:
            text_blob = json.dumps(vm, default=str).lower()
            name = str(vm.get("name") or vm.get("vm_name") or vm.get("vmName") or vm.get("uuid") or "")
            has_cvm_name = (
                name.lower().endswith("-cvm")
                or "-cvm" in name.lower()
                or "controller" in name.lower()
            )
            has_cvm_ip = any(ip in text_blob for ip in cvm_ips)
            if has_cvm_name or has_cvm_ip:
                out.append(vm)
        return out

    def _alerts(self):
        """
        Fetch active alerts from Prism Central v3 /alerts/list and scope them to
        this cluster using status.resources.originating_cluster_uuid.

        Important PC v3 alert fields:
          - status.resources.originating_cluster_uuid
          - status.resources.resolution_status.is_true
          - status.resources.severity
          - status.resources.title / default_message
          - status.resources.parameters
        """

        def _param_value(v):
            if not isinstance(v, dict):
                return str(v) if v is not None else ""
            for key in ("string_value", "int_value", "integer_value", "bool_value", "value"):
                if key in v and v.get(key) is not None:
                    return str(v.get(key))
            return ""

        def _render_template(template, params):
            """Replace Nutanix alert placeholders like {cvm_ip} with parameter values."""
            if not template:
                return "Alert"
            if not isinstance(params, dict):
                return template
            rendered = str(template)
            for key, value in params.items():
                if not key:
                    continue
                rendered = rendered.replace("{" + key + "}", _param_value(value))
            return rendered

        def _normalise_pc_v3_alert(entity):
            resources = entity.get("status", {}).get("resources", {}) if isinstance(entity, dict) else {}
            params = resources.get("parameters", {}) or {}
            severity = str(resources.get("severity", "unknown")).upper()
            title = _render_template(resources.get("title") or resources.get("default_message") or "Alert", params)
            message = _render_template(resources.get("default_message") or title, params)
            source_entity = resources.get("source_entity", {}).get("entity", {}) or {}
            affected = resources.get("affected_entity_list", []) or []

            # Nutanix places the suggested fix under possible_cause_list[*].resolution_list.
            # Join unique recommendations so the report is useful without dumping raw JSON.
            resolutions = []
            for cause in resources.get("possible_cause_list", []) or []:
                for item in cause.get("resolution_list", []) or []:
                    if item and item not in resolutions:
                        resolutions.append(str(item))
            recommended_resolution = " ".join(resolutions) if resolutions else "Review the alert in Prism Central and remediate per Nutanix guidance."

            classifications = resources.get("classification_list", []) or []
            impact_types = resources.get("impact_type_list", []) or []

            # Prefer the source entity name. If missing, fall back to the first affected entity.
            source_name = source_entity.get("name", "")
            if not source_name and affected and isinstance(affected[0], dict):
                source_name = affected[0].get("name", "")

            return {
                "extId": entity.get("metadata", {}).get("uuid", ""),
                "title": title,
                "message": message,
                "severity": severity,
                "creationTime": resources.get("creation_time") or entity.get("metadata", {}).get("creation_time", ""),
                "lastOccurred": resources.get("latest_occurrence_time") or resources.get("creation_time") or entity.get("metadata", {}).get("creation_time", ""),
                "lastUpdateTime": resources.get("last_update_time") or entity.get("metadata", {}).get("last_update_time", ""),
                "clusterUuid": resources.get("originating_cluster_uuid", ""),
                "sourceHost": source_name,
                "sourceEntity": source_name,
                "sourceType": source_entity.get("type", ""),
                "recommendedResolution": recommended_resolution,
                "classification": ", ".join(str(x) for x in classifications) if classifications else "N/A",
                "impactType": ", ".join(str(x) for x in impact_types) if impact_types else "N/A",
            }

        # Primary source: Prism Central v3 alerts/list. This is the endpoint that
        # returned the active alerts during testing.
        try:
            all_active = []
            offset = 0
            length = 500
            url = f"https://{self.client.host}:{self.client.port}/api/nutanix/v3/alerts/list"

            while True:
                payload = {"kind": "alert", "length": length, "offset": offset}
                resp = self.client.session.post(url, json=payload, timeout=60)
                resp.raise_for_status()
                data = resp.json()
                entities = data.get("entities", []) or []

                for e in entities:
                    resources = e.get("status", {}).get("resources", {}) if isinstance(e, dict) else {}
                    cluster_uuid = resources.get("originating_cluster_uuid")
                    resolved = resources.get("resolution_status", {}).get("is_true", False)
                    if cluster_uuid == self.cluster_ext_id and not resolved:
                        all_active.append(_normalise_pc_v3_alert(e))

                total = data.get("metadata", {}).get("total_matches")
                offset += length
                if not entities or (isinstance(total, int) and offset >= total) or len(entities) < length:
                    break

            return all_active
        except Exception as exc:
            print(f"WARNING: PC v3 alerts/list failed for {self.cluster_name}: {exc}")

        # Fallback: Prism Element v2 /alerts, if the cluster VIP is reachable.
        # Keep this as a fallback only because PC v3 provides all registered clusters.
        try:
            pe_data = self._pe_v2_get("/alerts?resolved=false")
            if pe_data is not None:
                entities = pe_data.get("entities", [])
                normalised = []
                sev_map = {
                    "kCritical": "CRITICAL", "kWarning": "WARNING",
                    "kInfo": "INFO", "kUnknown": "INFO",
                }
                for e in entities:
                    sev_raw = e.get("severity", "kWarning")
                    ts_usec = e.get("creation_time_stamp_in_usecs", 0)
                    ts_str = ""
                    if ts_usec:
                        dt = datetime.fromtimestamp(ts_usec / 1_000_000, tz=timezone.utc)
                        ts_str = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
                    normalised.append({
                        "extId": e.get("id", ""),
                        "title": e.get("alert_title", e.get("message", "Alert")),
                        "message": e.get("message", ""),
                        "severity": sev_map.get(sev_raw, str(sev_raw).replace("k", "").upper()),
                        "creationTime": ts_str,
                        "clusterUuid": self.cluster_ext_id,
                        "sourceHost": e.get("node_name") or e.get("host_name") or e.get("source_entity_name", "N/A"),
                        "recommendedResolution": e.get("resolution", "Review the alert in Prism Element and remediate per Nutanix guidance."),
                        "classification": e.get("classification", "N/A"),
                        "impactType": e.get("impact_type", "N/A"),
                    })
                return normalised
        except Exception as exc:
            print(f"WARNING: PE v2 alerts fallback failed for {self.cluster_name}: {exc}")

        return []

    def _protection_policies(self):
        # Current Prism Central Data Policies API.
        for ver in ["v4.2", "v4.1"]:
            try:
                response = self.client.get(
                    f"/datapolicies/{ver}/config/protection-policies",
                    {"$limit": 100},
                )
                policies = response.get("data", []) if isinstance(response, dict) else []
                if isinstance(policies, list) and policies:
                    return policies
            except Exception:
                continue

        # Compatibility fallbacks for older environments.
        for path in [
            "/dataprotection/v4.0/config/protection-policies",
            "/dataprotection/v4.0/config/protection-rules",
        ]:
            try:
                result = self.client.paginate(path)
                if result is not None:
                    return result
            except Exception:
                continue
        return []

    def _recovery_plans(self):
        """Collect Recovery Plans from v4, then fall back to legacy v3."""
        for ver in ["v4.2", "v4.1"]:
            try:
                response = self.client.get(
                    f"/datapolicies/{ver}/config/recovery-plans",
                    {"$limit": 100},
                )
                plans = response.get("data", []) if isinstance(response, dict) else []
                if not isinstance(plans, list):
                    continue

                enriched = []
                for plan in plans:
                    if not isinstance(plan, dict):
                        continue
                    item = dict(plan)
                    item["_api_generation"] = "v4"
                    ext_id = item.get("extId")
                    if ext_id:
                        for key, child_path in [
                            ("_stages", "stages"),
                            ("_network_mappings", "network-mappings"),
                        ]:
                            try:
                                child = self.client.get(
                                    f"/datapolicies/{ver}/config/recovery-plans/{ext_id}/{child_path}",
                                    {"$limit": 100},
                                )
                                rows = child.get("data", []) if isinstance(child, dict) else []
                                item[key] = rows if isinstance(rows, list) else []
                            except Exception:
                                item[key] = []
                    enriched.append(item)
                if enriched:
                    return enriched
            except Exception:
                continue

        # Recovery Plans created through the legacy Prism Central v3 API are
        # not returned by the v4 Data Policies list endpoint. Prism Central's
        # UI can therefore show plans while the v4 endpoint reports zero.
        plans = []
        offset = 0
        page_length = 100
        try:
            while True:
                response = self.client.post(
                    "/nutanix/v3/recovery_plans/list",
                    body={
                        "kind": "recovery_plan",
                        "length": page_length,
                        "offset": offset,
                    },
                )
                page = response.get("entities", []) if isinstance(response, dict) else []
                if not isinstance(page, list) or not page:
                    break
                for plan in page:
                    if isinstance(plan, dict):
                        item = dict(plan)
                        item["_api_generation"] = "v3"
                        plans.append(item)

                metadata = response.get("metadata", {}) if isinstance(response, dict) else {}
                total = metadata.get("total_matches")
                offset += len(page)
                if len(page) < page_length or (isinstance(total, int) and offset >= total):
                    break
        except Exception:
            return []
        return plans

    def _storage_containers(self):
        """Fetch storage containers; also used early to detect API version."""
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            try:
                result = self.client.paginate(
                    f"/clustermgmt/{ver}/config/storage-containers",
                    extra_params={"$filter": f"clusterExtId eq '{self.cluster_ext_id}'"}
                )
                if result:
                    # Extract and cache API version from the self-link
                    for link in result[0].get("links", []):
                        href = link.get("href", "")
                        if href:
                            self._api_ver = self.client.detect_api_version_from_link(href)
                            break
                    return result
            except Exception:
                continue
        return []

    def _container_stats(self) -> list:
        """
        Fetch per-container stats for each non-system container by following
        the storage-container-stats link embedded in each container object.
        Returns a list of {"name": ..., "extId": ..., "stats": {...}} dicts.
        """
        SYSTEM = {"NutanixManagementShare", "SelfServiceContainer",
                  "NutanixMetadataContainer"}
        # Also skip Nutanix Objects service containers (name starts with "objects")
        results = []
        for c in self.data.get("storage_containers", []):
            name = c.get("name", "")
            if name in SYSTEM or name.startswith("objects") or c.get("isInternal", False):
                continue
            ext_id = c.get("containerExtId") or c.get("extId")
            # Follow the stats self-link from the container object
            stats_path = None
            for link in c.get("links", []):
                if link.get("rel") == "storage-container-stats":
                    href = link.get("href", "")
                    if href and "/api/" in href:
                        stats_path = href.split("/api", 1)[1]
                    break
            # Fallback: build path from detected API version
            if not stats_path and ext_id:
                ver = self._api_version()
                stats_path = f"/clustermgmt/{ver}/stats/storage-containers/{ext_id}"

            if not stats_path:
                continue
            try:
                resp = self.client.get(stats_path)
                if resp:
                    results.append({
                        "name":   name,
                        "extId":  ext_id,
                        "stats":  resp,
                    })
            except Exception:
                continue
        return results

    def _pe_vip(self) -> Optional[str]:
        """Extract the Prism Element VIP from cluster_info.network.externalAddress."""
        try:
            ci = self.data.get("cluster_info", {})
            ci_data = ci.get("data", ci) if isinstance(ci, dict) else ci
            if isinstance(ci_data, list): ci_data = ci_data[0] if ci_data else {}
            net = ci_data.get("network", {}) if isinstance(ci_data, dict) else {}
            ip  = net.get("externalAddress", {}).get("ipv4", {}).get("value")
            return ip
        except Exception:
            return None


    def _find_numeric_values(self, obj: Any, wanted_keys: set[str]) -> list[float]:
        """Recursively find Nutanix numeric stat values.

        This helper must exist on ClusterDataCollector because _cluster_stats()
        and _pe_v2_stats() use it while data is being collected. A previous
        version only had this method on HealthAnalyser, so collection failed and
        cluster_stats was saved as null/None.
        """
        def norm(x: Any) -> str:
            return str(x).replace("-", "_").lower()

        wanted_norm = {norm(k) for k in wanted_keys}

        def add_numeric(v: Any, found: list[float]) -> None:
            if v is None:
                return
            if isinstance(v, list):
                for point in v:
                    if isinstance(point, (list, tuple)) and point:
                        add_numeric(point[-1], found)
                    elif isinstance(point, dict):
                        for vk in ("value", "avg", "average", "last", "max", "min"):
                            if vk in point:
                                add_numeric(point.get(vk), found)
                    else:
                        add_numeric(point, found)
                return
            try:
                found.append(float(v))
            except (TypeError, ValueError):
                pass

        found: list[float] = []

        if isinstance(obj, dict):
            metric_name = obj.get("metric") or obj.get("name") or obj.get("metricName")
            if metric_name is not None and norm(metric_name) in wanted_norm:
                for value_key in (
                    "values", "data", "series", "timeSeries", "time_series", "stats",
                    "value", "avg", "average", "last", "max", "min"
                ):
                    if value_key in obj:
                        add_numeric(obj.get(value_key), found)

            for k, v in obj.items():
                if norm(k) in wanted_norm:
                    add_numeric(v, found)
                found.extend(self._find_numeric_values(v, wanted_keys))

        elif isinstance(obj, list):
            for item in obj:
                found.extend(self._find_numeric_values(item, wanted_keys))

        return found

    def _pe_v2_get(self, path: str) -> Optional[dict]:
        """
        Call the Prism Element v2 REST API directly on the cluster VIP.
        Uses the same credentials as the PC session.
        """
        pe_ip = self._pe_vip()
        if not pe_ip:
            return None
        import requests as _req
        url  = f"https://{pe_ip}:{self.client.port}/PrismGateway/services/rest/v2.0{path}"
        try:
            resp = _req.get(
                url,
                auth=self.client.session.auth,
                verify=False,
                timeout=20,
            )
            if resp.status_code == 200:
                return resp.json()
        except Exception:
            pass
        return None

    def _pe_v2_stats(self) -> Optional[dict]:
        """Fall back to Prism Element stats directly on the cluster VIP.

        For this health check, the most reliable CPU metric is usually exposed
        by Prism Element legacy stats as hypervisor_cpu_usage_ppm. Prism Central
        v4 config inventory gives CPU capacity/core counts, but often does not
        return real-time utilisation.
        """
        pe_ip = self._pe_vip()
        if not pe_ip:
            return None

        import time as _time
        import requests as _req

        auth = self.client.session.auth
        end_usecs = int(_time.time() * 1000000)
        start_usecs = end_usecs - (15 * 60 * 1000000)

        attempts = []

        def add(label, base, path, params=None):
            attempts.append((label, f"https://{pe_ip}:{self.client.port}{base}{path}", params or {}))

        def add_cluster_stats(label_prefix, base):
            # Nutanix stats endpoints vary by API generation. Some expect
            # start_time_usecs/end_time_usecs, while older PE endpoints expect
            # startTimeInUsecs/endTimeInUsecs. Also, some builds require a
            # trailing slash after /stats/. Try all common forms.
            metric_sets = [
                "hypervisor_cpu_usage_ppm,controller_vm_cpu_usage_ppm",
                "hypervisor_cpu_usage_ppm",
                "cpu_usage_ppm",
            ]
            for metrics in metric_sets:
                for path in ["/cluster/stats", "/cluster/stats/"]:
                    add(f"{label_prefix} stats snake {metrics}", base, path, {
                        "metrics": metrics,
                        "start_time_usecs": start_usecs,
                        "end_time_usecs": end_usecs,
                        "interval_in_secs": 60,
                    })
                    add(f"{label_prefix} stats camel {metrics}", base, path, {
                        "metrics": metrics,
                        "startTimeInUsecs": start_usecs,
                        "endTimeInUsecs": end_usecs,
                        "intervalInSecs": 60,
                    })

        # Legacy PE v1: this is commonly where hypervisor_cpu_usage_ppm lives.
        add("PEv1 cluster", "/PrismGateway/services/rest/v1", "/cluster")
        add_cluster_stats("PEv1 cluster", "/PrismGateway/services/rest/v1")

        # PE v2 PrismGateway and /api/nutanix variants. Different AOS builds
        # expose different URL forms, so try both.
        for base in ["/PrismGateway/services/rest/v2.0", "/api/nutanix/v2.0"]:
            add("PEv2 cluster", base, "/cluster")
            add("PEv2 clusters", base, "/clusters")
            add("PEv2 cluster uuid", base, f"/clusters/{self.cluster_ext_id}")
            add_cluster_stats("PEv2 cluster", base)

        errors = []
        for label, url, params in attempts:
            try:
                resp = _req.get(url, auth=auth, verify=False, timeout=20, params=params)
                if resp.status_code == 200:
                    payload = resp.json()
                    if self._find_numeric_values(payload, {
                        "hypervisor_cpu_usage_ppm", "hypervisorCpuUsagePpm",
                        "cpu_usage_ppm", "cpuUsagePpm",
                        "controller_vm_cpu_usage_ppm", "controllerVmCpuUsagePpm",
                    }):
                        return {"data": {"_pe_stats": payload, "_source": label, "_url": url}}
                    # Keep the response only as a last-resort debug object.
                    errors.append({"source": label, "status": resp.status_code, "note": "200 but no CPU metric keys"})
                else:
                    errors.append({"source": label, "status": resp.status_code})
            except Exception as exc:
                errors.append({"source": label, "error": str(exc)})

        # Return errors so the raw JSON shows exactly what was tried.
        return {"data": {"_pe_stats_errors": errors}}

    def _storage_stats(self):
        """Fetch storage statistics from Prism Central v4 stats links.

        PC v4 exposes per-object stats from links, for example the
        storage-container-stats links already present on each container:
        /api/clustermgmt/v4.2/stats/storage-containers/<container_ext_id>
        """
        results = []

        def path_from_href(href: str) -> Optional[str]:
            if not href or "/api/" not in href:
                return None
            return href.split("/api", 1)[1]

        # 1) Follow storage-container-stats links returned in storage_containers.
        for c in self.data.get("storage_containers", []) or []:
            for link in c.get("links", []) or []:
                if link.get("rel") == "storage-container-stats":
                    path = path_from_href(link.get("href", ""))
                    if path:
                        try:
                            stat = self.client.get(path)
                            if stat:
                                results.append({"containerExtId": c.get("containerExtId") or c.get("extId"),
                                                "name": c.get("name"),
                                                "stats": stat})
                        except Exception:
                            pass

        # 2) Try direct cluster-scoped storage stats endpoints as a fallback.
        for ver in [self._api_version(), "v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            for path in [
                f"/clustermgmt/{ver}/stats/clusters/{self.cluster_ext_id}",
                f"/clustermgmt/{ver}/stats/storage-pools",
            ]:
                try:
                    stat = self.client.get(path)
                    if stat:
                        results.append({"source": path, "stats": stat})
                except Exception:
                    continue

        if results:
            return {"data": {"_storage_object_stats": results}}

        # 3) Final fallback to PE v2 /cluster, which often contains usageStats.
        return self._pe_v2_stats()

    def _cluster_stats(self):
        """Fetch CPU and memory stats.

        IMPORTANT: Some PC v4 stats endpoints return a valid JSON payload, but
        do not include CPU usage keys. The previous version returned the first
        non-empty stats response, which prevented the PE v2 fallback from being
        used. This version only returns a stats payload if it actually contains
        CPU or memory usage values.
        """
        detected = self._api_version()
        versions = []
        for ver in [detected, "v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            if ver not in versions:
                versions.append(ver)

        def has_cpu_or_memory_values(payload: Any) -> bool:
            cpu_keys = {
                "cpuUsagePpm", "hypervisorCpuUsagePpm", "hypervisor_cpu_usage_ppm",
                "controllerVmCpuUsagePpm", "controller_vm_cpu_usage_ppm",
                "cpu_usage_ppm", "aggregateCpuUsagePpm", "aggregate_cpu_usage_ppm",
            }
            mem_keys = {
                "memoryUsagePpm", "hypervisorMemoryUsagePpm", "hypervisor_memory_usage_ppm",
                "controllerVmMemoryUsagePpm", "controller_vm_memory_usage_ppm",
                "memory_usage_ppm", "aggregateMemoryUsagePpm", "aggregate_memory_usage_ppm",
            }
            return bool(self._find_numeric_values(payload, cpu_keys | mem_keys))

        # 1) Try PC v4 cluster-level stats with the required time range.
        # PC v4 stats returns time-series data such as:
        #   data.hypervisorCpuUsagePpm[].value
        # where the value is PPM.  PPM / 10000 = percent.
        end_time = datetime.now(timezone.utc).replace(microsecond=0)
        start_time = end_time - timedelta(minutes=10)
        stats_params = {
            "$startTime": start_time.isoformat().replace("+00:00", "Z"),
            "$endTime": end_time.isoformat().replace("+00:00", "Z"),
        }

        for ver in versions:
            for path in [
                f"/clustermgmt/{ver}/stats/clusters/{self.cluster_ext_id}",
            ]:
                try:
                    result = self.client.get(path, params=stats_params)
                    if result and result.get("data") is not None and has_cpu_or_memory_values(result):
                        return result
                except Exception as exc:
                    print(f"        PC cluster stats failed for {ver}: {exc}")
                    continue

        # 2) Try PC v4 host-level stats with the same required time range.
        host_stats = []
        hosts = self.data.get("nodes") or self._nodes() or []
        for h in hosts:
            host_id = (h.get("extId") or h.get("hostExtId") or h.get("uuid") or
                       h.get("nodeUuid") or h.get("host", {}).get("extId"))
            if not host_id:
                continue
            for ver in versions:
                for path in [
                    f"/clustermgmt/{ver}/stats/hosts/{host_id}",
                    f"/clustermgmt/{ver}/stats/host-nodes/{host_id}",
                ]:
                    try:
                        result = self.client.get(path, params=stats_params)
                        if result and result.get("data") is not None and has_cpu_or_memory_values(result):
                            host_stats.append({"hostExtId": host_id, "stats": result})
                            raise StopIteration
                    except StopIteration:
                        break
                    except Exception:
                        continue
                else:
                    continue
                break

        if host_stats:
            return {"data": {"_host_stats": host_stats}}

        # 3) No Prism Element fallback here. This report is intended to use
        # Prism Central for cluster CPU stats.
        return None

    def _cluster_stats_7d(self):
        """Fetch 7 days of Prism Central cluster stats for CPU charting.

        PC v4 requires a time range for stats endpoints. This returns the
        cluster stats payload containing data.hypervisorCpuUsagePpm[].value
        where value is PPM. The analyser converts PPM to percent.
        """
        detected = self._api_version()
        versions = []
        for ver in [detected, "v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            if ver not in versions:
                versions.append(ver)

        end_time = datetime.now(timezone.utc).replace(microsecond=0)
        start_time = end_time - timedelta(days=7)
        stats_params = {
            "$startTime": start_time.isoformat().replace("+00:00", "Z"),
            "$endTime": end_time.isoformat().replace("+00:00", "Z"),
        }

        for ver in versions:
            path = f"/clustermgmt/{ver}/stats/clusters/{self.cluster_ext_id}"
            try:
                result = self.client.get(path, params=stats_params)
                if result and result.get("data") and self._find_numeric_values(result, {"hypervisorCpuUsagePpm", "hypervisor_cpu_usage_ppm"}):
                    return result
            except Exception as exc:
                print(f"        PC 7-day cluster stats failed for {ver}: {exc}")
                continue
        return None

    def _networks(self):
        """Fetch subnets/VLANs using detected API version."""
        detected = self.client.detect_api_version("networking")
        for ver in [detected, "v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            for filt in [
                f"clusterExtId eq '{self.cluster_ext_id}'",
                f"clusterReference eq '{self.cluster_ext_id}'",
            ]:
                try:
                    result = self.client.paginate(
                        f"/networking/{ver}/config/subnets",
                        extra_params={"$filter": filt}
                    )
                    if result:
                        return result
                except Exception:
                    continue
        return []

    def _network_details(self):
        """Best-effort network inventory collection for NICs, bonds, bridges, and vSwitches.

        Prism Central API coverage varies across AOS/PC releases. The report uses
        whatever this method can retrieve and gracefully marks unsupported/missing
        fields as unavailable instead of failing the health check.
        """
        detected = self.client.detect_api_version("networking")
        versions = []
        for ver in [detected, "v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            if ver not in versions:
                versions.append(ver)

        details = {
            "virtual_switches": [],
            "host_nics": [],
            "bonds": [],
            "bridges": [],
            "raw_networking": [],
            "errors": [],
        }

        def add_error(source, exc):
            details["errors"].append({"source": source, "error": str(exc)[:300]})

        # Priority collection: physical NIC details. Keep this early and focused
        # so the report does not exhaust PC API rate limits on lower-value
        # discovery endpoints before calling the known-good host-nics API.
        host_items = self.data.get("nodes") or []
        if isinstance(host_items, dict):
            host_items = host_items.get("data") or []
        cm_versions = []
        try:
            cm_versions.append(self._api_version())
        except Exception:
            pass
        for ver in ["v4.2", "v4.1", "v4.0.b1", "v4.0"]:
            if ver not in cm_versions:
                cm_versions.append(ver)

        # Use cached node IDs first; only fetch the host list if those are absent.
        if not any(isinstance(h, dict) and (h.get("extId") or h.get("uuid") or h.get("nodeUuid")) for h in host_items):
            for ver in cm_versions[:2]:
                host_path = f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/hosts"
                try:
                    fetched_hosts = self.client.paginate(host_path)
                    if fetched_hosts:
                        host_items = fetched_hosts
                        break
                except Exception as exc:
                    add_error(host_path, exc)

        for host in host_items or []:
            if not isinstance(host, dict):
                continue
            host_id = host.get("extId") or host.get("uuid") or host.get("nodeUuid")
            host_name = host.get("hostName") or host.get("name") or host_id
            if not host_id:
                continue
            for ver in cm_versions[:2]:
                path = f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/hosts/{host_id}/host-nics"
                try:
                    raw = self.client.get(path)
                    result = raw.get("data", []) if isinstance(raw, dict) else []
                    if result:
                        for nic in result:
                            if isinstance(nic, dict):
                                nic.setdefault("hostName", host_name)
                                nic.setdefault("hostExtId", host_id)
                                details["host_nics"].append(nic)
                        break
                except Exception as exc:
                    add_error(path, exc)
                    continue

        # Best-effort cluster action introduced in newer PC/AOS releases. This
        # is the most likely Prism Central path for host NIC, bond/uplink, and
        # bridge details. Different releases accept different request shapes, so
        # try a small set and let the later parser extract whatever is returned.
        node_ids = []
        try:
            for n in (self.data.get("nodes") or []):
                nid = n.get("extId") or n.get("uuid") or n.get("nodeUuid")
                if nid:
                    node_ids.append(nid)
        except Exception:
            node_ids = []

        for ver in []:  # disabled by default; noisy and often rate-limited/unsupported
            action_path = f"/clustermgmt/{ver}/config/clusters/{self.cluster_ext_id}/$actions/fetch-node-networking-details"
            action_bodies = [
                {},
                {"clusterExtId": self.cluster_ext_id},
                {"nodeExtIds": node_ids},
                {"hostExtIds": node_ids},
                {"nodeList": node_ids},
            ]
            for body in action_bodies:
                try:
                    result = self.client.post(action_path, body=body)
                    if result:
                        details["raw_networking"].append(result)
                        break
                except Exception as exc:
                    add_error(action_path, exc)
                    continue
            if details["raw_networking"]:
                break

        # Try common PC networking inventory endpoints. These are intentionally
        # best-effort because names/filters differ by PC API release.
        endpoint_specs = [
            ("virtual_switches", [
                "/networking/{ver}/config/virtual-switches",
                "/networking/{ver}/config/virtual-switches?",
            ]),
            ("host_nics", [
                "/networking/{ver}/config/host-nics",
                "/networking/{ver}/config/nics",
                "/networking/{ver}/config/physical-nics",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/host-nics",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/physical-nics",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/nics",
            ]),
            ("bonds", [
                "/networking/{ver}/config/bonds",
                "/networking/{ver}/config/host-nic-teams",
                "/networking/{ver}/config/host-nic-bonds",
                "/networking/{ver}/config/uplinks",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/bonds",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/uplinks",
                "/clustermgmt/{ver}/config/clusters/" + self.cluster_ext_id + "/host-nic-bonds",
            ]),
            ("bridges", [
                "/networking/{ver}/config/bridges",
            ]),
        ]

        filters = [
            {"$filter": f"clusterExtId eq '{self.cluster_ext_id}'"},
            {"$filter": f"clusterReference eq '{self.cluster_ext_id}'"},
            {"$filter": f"clusterReferenceList/any(c:c eq '{self.cluster_ext_id}')"},
            {},
        ]

        for key, paths in endpoint_specs:
            for ver in versions:
                if details[key]:
                    break
                for path_tpl in paths:
                    if details[key]:
                        break
                    path = path_tpl.format(ver=ver).replace("?", "")
                    for extra in filters:
                        try:
                            result = self.client.paginate(path, extra_params=extra)
                            if result:
                                # Client-side filter if the endpoint does not support the filter.
                                filtered = []
                                for item in result:
                                    text = json.dumps(item, default=str)
                                    if (self.cluster_ext_id in text) or not extra:
                                        filtered.append(item)
                                details[key] = filtered or result
                                break
                        except Exception as exc:
                            if not extra:
                                add_error(path, exc)
                            continue

        # Physical NIC collection is handled by the priority pass above.
        # Do not repeat the same host-nics calls here; repeated probing can
        # trigger Prism Central API 429 rate-limit responses on multi-cluster runs.

        # Use PE legacy APIs as another source when reachable. These often expose
        # host network inventory on some AOS releases.
        for pe_path, key in [
            ("/hosts", "host_nics"),
            ("/networks", "bridges"),
        ]:
            if details.get(key):
                continue
            try:
                pe = self._pe_v2_get(pe_path)
                entities = pe.get("entities", []) if isinstance(pe, dict) else []
                if entities:
                    details[key] = entities
            except Exception as exc:
                add_error("PEv2 " + pe_path, exc)

        return details

    def _licensing(self):
        # PC-level licensing API paths (try newest first).
        for path in [
            "/licensing/v4.0/config/licenses",
            "/licensing/v4.0/config/license-keys",
            "/licensing/v4.0/config/clusters",
            "/licensing/v4.0.b1/config/licenses",
        ]:
            try:
                result = self.client.get(path)
                if result and result.get("data"):
                    return result
            except Exception:
                continue
        return None

    def _ncc_checks(self):
        try:
            return self.client.paginate(
                "/opsmgmt/v4.0/monitoring/alerts",
                extra_params={
                    "$filter": (
                        f"severity eq 'HEALTH_CHECK' and "
                        f"clusterUuids eq '{self.cluster_ext_id}'"
                    )
                }
            )
        except Exception:
            return []


# ---------------------------------------------------------------------------
# Health Analyser
# ---------------------------------------------------------------------------

class HealthAnalyser:

    STATUS_HEALTHY     = "Healthy"
    STATUS_RECOMMENDED = "Recommended"
    STATUS_CRITICAL    = "Critical"

    def __init__(self, raw: dict, customer: str, cluster_name: str, os_compat_csv: str = "", aos_eol_csv: str = ""):
        self.raw          = raw
        self.customer     = customer
        self.cluster_name = cluster_name
        self.os_compat_csv = _find_required_file("OS Compatibility Matrix CSV", OS_COMPAT_CSV_FILENAMES, os_compat_csv)
        self.aos_eol_csv   = _find_required_file("AOS/NOS EOL information CSV", AOS_EOL_CSV_FILENAMES, aos_eol_csv)

    def _safe_list(self, key) -> list:
        v = self.raw.get(key)
        if isinstance(v, list):
            return v
        if isinstance(v, dict) and "data" in v:
            d = v["data"]
            return d if isinstance(d, list) else ([d] if d else [])
        return []

    def _cluster_first(self) -> dict:
        ci = self.raw.get("cluster_info", {})
        if isinstance(ci, dict):
            data = ci.get("data", ci)
            if isinstance(data, list) and data:
                return data[0]
            if isinstance(data, dict):
                return data
        return {}

    # ── section analysers ────────────────────────────────────────────────

    def analyse_cluster_info(self) -> dict:
        c   = self._cluster_first()
        cfg = c.get("config", {})
        bi  = cfg.get("buildInfo", {})

        # Node count: cluster_info.nodes.numberOfNodes (confirmed from raw JSON)
        nodes_ref  = c.get("nodes", {})
        node_count = nodes_ref.get("numberOfNodes", "N/A")
        # Also count nodeList as cross-check
        if node_count == "N/A":
            node_list = nodes_ref.get("nodeList", [])
            node_count = len(node_list) if node_list else "N/A"

        # NCC version: config.clusterSoftwareMap is a LIST of {softwareType, version}
        # e.g. [{"softwareType": "NCC", "version": "ncc-5.3.1.1"}, ...]
        csm     = cfg.get("clusterSoftwareMap", [])
        ncc_ver = "N/A"
        aos_full = "N/A"
        if isinstance(csm, list):
            for entry in csm:
                stype = entry.get("softwareType", "")
                ver   = entry.get("version", "")
                if stype == "NCC":
                    # Strip "ncc-" prefix for display
                    ncc_ver = ver.replace("ncc-", "") if ver else "N/A"
                elif stype == "NOS":
                    aos_full = ver

        # AOS version: buildInfo.version is only major.minor ("7.5").
        # Extract the full x.y.z.w from fullVersion or the NOS clusterSoftwareMap entry.
        full_ver_str = bi.get("fullVersion", "")
        m = re.search(r"(\d+\.\d+\.\d+\.\d+)", full_ver_str)
        if m:
            aos_ver = m.group(1)          # e.g. "7.5.1.6"
        else:
            # Fall back to NOS entry in clusterSoftwareMap
            m2 = re.search(r"(\d+\.\d+\.\d+\.\d+)", aos_full)
            aos_ver = m2.group(1) if m2 else bi.get("version", "N/A")

        # Hypervisor
        hyp_types = cfg.get("hypervisorTypes", [])
        hyp = str(hyp_types[0]) if hyp_types else "N/A"

        # AHV version from per-node API call
        ahv_ver = self.raw.get("ahv_version") or "N/A"

        # Network details
        net = c.get("network", {})
        cluster_vip  = net.get("externalAddress", {}).get("ipv4", {}).get("value", "N/A")
        data_svc_ip  = net.get("externalDataServiceIp", {}).get("ipv4", {}).get("value", "N/A")
        dns_servers  = [s.get("ipv4", {}).get("value", s.get("fqdn", {}).get("value", ""))
                        for s in net.get("nameServerIpList", [])]
        ntp_servers  = [s.get("fqdn", {}).get("value", s.get("ipv4", {}).get("value", ""))
                        for s in net.get("ntpServerIpList", [])]

        # Cluster type / fault tolerance
        fault_state  = cfg.get("faultToleranceState", {})
        cluster_type = cfg.get("clusterType", "N/A")
        cluster_arch = cfg.get("clusterArch", "N/A")

        return {
            "cluster_name":      c.get("name", self.cluster_name),
            "cluster_uuid":      c.get("extId", "N/A"),
            "node_count":        node_count,
            "aos_version":       aos_ver,
            "ahv_version":       ahv_ver,
            "ncc_version":       ncc_ver,
            "timezone":          cfg.get("timezone", "N/A"),
            "redundancy_factor": cfg.get("redundancyFactor", "N/A"),
            "hypervisor":        hyp,
            "cluster_vip":       cluster_vip,
            "data_svc_ip":       data_svc_ip,
            "dns_servers":       ", ".join(filter(None, dns_servers)) or "N/A",
            "ntp_servers":       ", ".join(filter(None, ntp_servers)) or "N/A",
            "cluster_type":      cluster_type,
            "cluster_arch":      cluster_arch,
            "fault_tolerance":   fault_state.get("currentClusterFaultTolerance", "N/A"),
            "license_type":      None,  # populated by analyse_licensing
            "vm_count":          c.get("vmCount", "N/A"),
            "upgrade_status":    c.get("upgradeStatus", "N/A"),
            "remote_support_enabled": cfg.get("isRemoteSupportEnabled", "N/A"),
            "password_remote_login_enabled": cfg.get("isPasswordRemoteLoginEnabled", "N/A"),
            "pulse_enabled":     cfg.get("pulseStatus", {}).get("isEnabled", "N/A"),
        }

    def analyse_virtual_machines(self) -> dict:
        vms  = self._safe_list("virtual_machines")
        on_  = [v for v in vms if str(v.get("powerState", "")).upper() == "ON"]
        off_ = [v for v in vms if str(v.get("powerState", "")).upper() == "OFF"]
        recs = []

        GiB = 1024 ** 3
        cluster_info_for_os = self.analyse_cluster_info()
        current_aos_for_os = cluster_info_for_os.get("aos_version", "")
        os_matrix = _load_os_compatibility_matrix(self.os_compat_csv, current_aos_for_os)

        # Windows build number → friendly name (source: Microsoft docs)
        WIN_BUILD = {
            "26100": "Windows 11 24H2",       "22631": "Windows 11 23H2",
            "22621": "Windows 11 22H2",       "22000": "Windows 11 21H2",
            "19045": "Windows 10 22H2",       "19044": "Windows 10 21H2",
            "19043": "Windows 10 21H1",       "19042": "Windows 10 20H2",
            "19041": "Windows 10 2004",       "18363": "Windows 10 1909",
            "26040": "Windows Server 2025",   "20348": "Windows Server 2022",
            "17763": "Windows Server 2019",   "14393": "Windows Server 2016",
            "9600":  "Windows Server 2012 R2","9200":  "Windows Server 2012",
            "7601":  "Windows Server 2008 R2",
        }

        def _os_name(gt: dict) -> str:
            gi  = gt.get("guestInfo", {})
            build = str(gi.get("guestOsBuildNumber", ""))

            # Prefer build-number lookup for Windows (most accurate)
            if build and build in WIN_BUILD:
                return WIN_BUILD[build]

            # Fall back to parsing the raw OS string
            os_raw = (gi.get("guestOsFullName") or gt.get("guestOsVersion", ""))
            for prefix in ["windows:64:", "linux:64:", "linux:32:", "windows:32:"]:
                if os_raw.startswith(prefix):
                    os_raw = os_raw[len(prefix):]

            # Simplify common Linux distro names
            replacements = [
                ("Rocky Linux",   "Rocky Linux"),
                ("CentOS Linux",  "CentOS"),
                ("Red Hat Enterprise Linux", "RHEL"),
                ("Ubuntu",        "Ubuntu"),
                ("Debian",        "Debian"),
                ("SUSE Linux",    "SUSE"),
            ]
            for full, short in replacements:
                if full.lower() in os_raw.lower():
                    # Extract version number
                    import re as _re
                    m = _re.search(r"[0-9]+(?:\.[0-9]+)*", os_raw)
                    return f"{short} {m.group(0)}" if m else short

            # Windows fallback from string (no build number available)
            win_map = [
                ("WindowsServer2025",    "Windows Server 2025"),
                ("WindowsServer2022",    "Windows Server 2022"),
                ("WindowsServer2019",    "Windows Server 2019"),
                ("WindowsServer2016",    "Windows Server 2016"),
                ("WindowsServer2012R2",  "Windows Server 2012 R2"),
                ("WindowsServer2012",    "Windows Server 2012"),
                ("Windows11",            "Windows 11"),
                ("Windows10",            "Windows 10"),
                ("Windows7",             "Windows 7"),
            ]
            for key, name in win_map:
                if key.lower() in os_raw.lower().replace(" ", ""):
                    return name

            return os_raw.strip() or "Unknown"

        def _primary_ip(vm: dict) -> str:
            for nic in vm.get("nics", []):
                learned = (nic.get("networkInfo", nic.get("nicNetworkInfo", {}))
                           .get("ipv4Info", {}).get("learnedIpAddresses", []))
                if learned:
                    return learned[0].get("value", "N/A")
            return "N/A"

        vm_list = []
        # System VMs that don't need NGT and should be excluded from recommendations
        def _is_system_vm(name: str) -> bool:
            n = name.upper()
            return (
                ("PCVM" in n and n.startswith("NTNX-")) or  # Prism Central VM
                n == "FOUNDATION" or                         # Foundation VM
                n.startswith("FOUNDATION-") or
                "NUTANIX-MOVE" in n or                       # Nutanix Move appliance
                n.startswith("NTNX-") and ("CVM" in n)       # Any other Nutanix system VM
            )

        ngt_missing  = []
        ngt_outdated = []
        os_legacy = []
        os_unsupported = []
        os_supported = []
        system_vm_ignored = []

        for vm in vms:
            vm_name = vm.get("name", "Unknown")
            is_system_vm = _is_system_vm(vm_name)
            gt        = vm.get("guestTools", {})
            ngt_inst  = gt.get("isInstalled", False)
            ngt_ver   = gt.get("version", "") if ngt_inst else ""
            ngt_avail = gt.get("availableVersion", "")
            ngt_reach = gt.get("isReachable", False)

            if not ngt_inst:
                ngt_status = "Not Installed"
                if not is_system_vm:
                    ngt_missing.append(vm_name)
            elif not ngt_reach:
                ngt_status = "Unreachable"
            else:
                ngt_status = "Enabled"

            if ngt_inst and ngt_avail and ngt_ver and ngt_ver != ngt_avail:
                if not is_system_vm:
                    ngt_outdated.append(vm_name)

            vcpus    = vm.get("numSockets", 0) * vm.get("numCoresPerSocket", 1)
            mem_gib  = round(vm.get("memorySizeBytes", 0) / GiB, 1)
            disk_gib = round(sum(
                d.get("backingInfo", {}).get("diskSizeBytes", 0)
                for d in vm.get("disks", [])
            ) / GiB, 1)

            # CD-ROMs: list mounted ISOs; ignore empty drives
            cdrom_list = []
            for cd in vm.get("cdRoms", []):
                iso_type = cd.get("isoType", "")
                backing  = cd.get("backingInfo", cd.get("backing", {}))
                iso_file = (backing.get("isoFile", "") or
                            backing.get("dataSourceReference", {}).get("name", ""))
                if iso_file:
                    # Show just the filename, not the full datastore path
                    cdrom_list.append(iso_file.split("/")[-1])
                elif iso_type not in ("", "OTHER", "EMPTY"):
                    cdrom_list.append(iso_type)
            cdrom_str = ", ".join(cdrom_list) if cdrom_list else "—"

            os_name = _os_name(gt)
            if is_system_vm:
                os_support_info = {"support": "Ignored", "matched_os": "Nutanix system VM"}
                os_support = "Ignored"
                system_vm_ignored.append(vm_name)
            else:
                os_support_info = _lookup_os_support(os_name, os_matrix)
                os_support = os_support_info.get("support", "Unsupported")
                if os_support == "Supported":
                    os_supported.append(vm_name)
                elif os_support == "Legacy Support":
                    os_legacy.append(vm_name)
                else:
                    os_unsupported.append(vm_name)

            vm_list.append({
                "name":        vm_name,
                "power":       vm.get("powerState", "UNKNOWN"),
                "ip":          _primary_ip(vm),
                "os":          os_name,
                "os_support":  os_support,
                "os_match":    os_support_info.get("matched_os", "N/A"),
                "vcpus":       vcpus,
                "mem_gib":     mem_gib,
                "disk_gib":    disk_gib,
                "cdrom":       cdrom_str,
                "ngt_status":  ngt_status,
                "ngt_version": ngt_ver or "N/A",
            })

        # CD-ROM recommendations
        cdrom_mounted = [
            vm["name"] for vm in vm_list
            if vm["cdrom"] != "—" and not _is_system_vm(vm["name"])
        ]

        if ngt_missing:
            recs.append(
                f"Install Nutanix Guest Tools (NGT) on the following VM(s): "
                f"{', '.join(ngt_missing)}. "
                f"NGT enables crash-consistent snapshots, VSS integration, "
                f"and self-service restore."
            )
        if ngt_outdated:
            recs.append(
                f"Upgrade Nutanix Guest Tools to the latest available version on: "
                f"{', '.join(ngt_outdated)}. "
                f"Upgrade via Prism Central → VM → Actions → Install NGT."
            )
        if cdrom_mounted:
            recs.append(
                f"Unmount CD-ROM ISO on the following VM(s): "
                f"{', '.join(cdrom_mounted)}. "
                f"Leaving ISOs mounted consumes storage snapshot space and is a "
                f"Nutanix best practice to resolve."
            )

        if system_vm_ignored:
            recs.append(
                "Foundation and Prism Central system VMs are excluded from NGT and guest OS "
                "compatibility recommendations because they are Nutanix appliances and do not support NGT."
            )

        if os_unsupported:
            recs.append(
                f"Upgrade or replace unsupported guest operating systems as soon as practical: "
                f"{', '.join(os_unsupported)}."
            )
        if os_legacy:
            recs.append(
                f"Plan migration of legacy guest operating systems to supported versions: "
                f"{', '.join(os_legacy)}."
            )

        # Determine overall status
        if os_unsupported:
            status = self.STATUS_CRITICAL
        elif os_legacy or ngt_missing or cdrom_mounted:
            status = self.STATUS_RECOMMENDED
        else:
            status = self.STATUS_HEALTHY

        return {
            "status":          status,
            "total":           len(vms),
            "powered_on":      len(on_),
            "powered_off":     len(off_),
            "vm_list":         vm_list,
            "ngt_missing":     ngt_missing,
            "ngt_outdated":    ngt_outdated,
            "cdrom_mounted":   cdrom_mounted,
            "os_supported":     os_supported,
            "os_legacy":        os_legacy,
            "os_unsupported":   os_unsupported,
            "system_vm_ignored": system_vm_ignored,
            "os_matrix_loaded": bool(os_matrix),
            "os_aos_version":   _aos_minor_version(current_aos_for_os),
            "recommendations": recs or ["No immediate action required."],
        }

    def _vm_summary_counts(self) -> dict:
        """Called after both VMs and CVMs are analysed to build the combined totals."""
        vms  = self.raw.get("virtual_machines", [])
        nodes = self._safe_list("nodes")
        vm_on  = len([v for v in vms if str(v.get("powerState", "")).upper() == "ON"])
        vm_off = len([v for v in vms if str(v.get("powerState", "")).upper() == "OFF"])
        cvm_count = len(nodes)
        return {
            "user_vms_on":  vm_on,
            "user_vms_off": vm_off,
            "cvm_count":    cvm_count,
            "total_on":     vm_on + cvm_count,   # CVMs are always on if node is healthy
            "total_vms":    len(vms) + cvm_count,
        }

    def analyse_health(self) -> dict:
        alerts   = self._safe_list("alerts")
        critical = [a for a in alerts if str(a.get("severity", "")).upper() in ("CRITICAL", "ERROR")]
        warnings = [a for a in alerts if str(a.get("severity", "")).upper() == "WARNING"]
        recs = []
        if critical:
            recs.append(f"Investigate {len(critical)} Critical alert(s) immediately.")
        if warnings:
            recs.append(f"Review {len(warnings)} Warning alert(s) with Nutanix Support.")
        status = (self.STATUS_CRITICAL    if critical else
                  self.STATUS_RECOMMENDED if warnings else
                  self.STATUS_HEALTHY)
        return {
            "status":          status,
            "total_alerts":    len(alerts),
            "critical_alerts": len(critical),
            "warning_alerts":  len(warnings),
            "alert_details": [
                {
                    "severity": a.get("severity", "UNKNOWN"),
                    "title":    a.get("title", a.get("message", "Unknown alert")),
                    "source_host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "last_occurred":  a.get("lastOccurred") or a.get("creationTime", ""),
                    "classification": a.get("classification", "N/A"),
                    "impact_type": a.get("impactType", "N/A"),
                }
                for a in alerts[:20]
            ],
            "recommendations": recs or ["No immediate action required."],
        }

    def analyse_cvms(self) -> dict:
        """Build CVM inventory from hosts, enriched with optional PE VM details."""
        nodes = self._safe_list("nodes")
        pe_cvms = self._safe_list("cvm_virtual_machines")
        cluster_info = (self.raw.get("cluster_info", {}) or {}).get("data", {}) or {}
        pubkeys = ((cluster_info.get("config", {}) or {}).get("authorizedPublicKeyList", []) or [])

        def ip_value(obj):
            if not isinstance(obj, dict):
                return None
            if obj.get("value"):
                return obj.get("value")
            if obj.get("ipv4"):
                return ip_value(obj.get("ipv4"))
            if obj.get("ip"):
                return ip_value(obj.get("ip"))
            return None

        def bytes_to_gib(v):
            try:
                if v is None or v == "":
                    return "N/A"
                return round(float(v) / 1024**3, 0)
            except Exception:
                return "N/A"

        def walk_dicts(obj):
            if isinstance(obj, dict):
                yield obj
                for val in obj.values():
                    yield from walk_dicts(val)
            elif isinstance(obj, list):
                for val in obj:
                    yield from walk_dicts(val)

        def first_nested(obj, keys):
            for d in walk_dicts(obj):
                for key in keys:
                    if d.get(key) not in (None, ""):
                        return d.get(key)
            return None

        def vm_name_from_key(node_uuid, cvm_ip):
            for k in pubkeys:
                if not isinstance(k, dict):
                    continue
                name = str(k.get("name", ""))
                key = str(k.get("key", ""))
                if name not in (str(node_uuid), str(cvm_ip)):
                    continue
                # Public key comments commonly end with nutanix@ntnx-...-cvm.
                comment = key.split()[-1] if key.split() else ""
                if "@" in comment:
                    candidate = comment.split("@", 1)[1]
                    if candidate:
                        return candidate
            return "N/A"

        def extract_vm_ips(vm):
            ips = set()
            def walk(x):
                if isinstance(x, dict):
                    # Common IP structures
                    if "ipAddress" in x and isinstance(x.get("ipAddress"), str):
                        ips.add(x.get("ipAddress"))
                    if "ip_address" in x and isinstance(x.get("ip_address"), str):
                        ips.add(x.get("ip_address"))
                    val = ip_value(x)
                    if val and isinstance(val, str) and val.count(".") == 3:
                        ips.add(val)
                    for v in x.values():
                        walk(v)
                elif isinstance(x, list):
                    for v in x:
                        walk(v)
            walk(vm)
            return ips

        pe_by_ip = {}
        pe_by_name = {}
        for vm in pe_cvms:
            if not isinstance(vm, dict):
                continue
            nm = str(vm.get("name") or vm.get("vm_name") or vm.get("vmName") or "")
            if nm:
                pe_by_name[nm.lower()] = vm
            for ip in extract_vm_ips(vm):
                pe_by_ip[str(ip)] = vm

        def get_vm_vcpus(vm):
            if not isinstance(vm, dict):
                return "N/A"
            # PC/PE structures vary: explicit total vCPU, or sockets*cores*threads.
            direct = first_nested(vm, ("num_vcpus", "numVcpus", "numVCPUs", "vcpus"))
            if direct not in (None, ""):
                try:
                    return int(direct)
                except Exception:
                    return direct
            try:
                sockets = int(first_nested(vm, ("numSockets", "num_sockets", "num_sockets_per_vm", "numSocketsPerVm")) or 1)
                cores = int(first_nested(vm, ("numCoresPerSocket", "num_cores_per_socket", "numCores")) or 1)
                threads = int(first_nested(vm, ("numThreadsPerCore", "num_threads_per_core")) or 1)
                val = sockets * cores * threads
                return val if val > 0 else "N/A"
            except Exception:
                return "N/A"

        def get_vm_memory_gib(vm):
            if not isinstance(vm, dict):
                return "N/A"
            b = first_nested(vm, ("memorySizeBytes", "memory_size_bytes"))
            if b not in (None, ""):
                return bytes_to_gib(b)
            mib = first_nested(vm, ("memory_size_mib", "memorySizeMib", "memorySizeMiB", "memorySizeInMib"))
            if mib not in (None, ""):
                try:
                    return round(float(mib) / 1024, 0)
                except Exception:
                    return mib
            mb = first_nested(vm, ("memory_mb", "memoryMb", "memorySizeInMb"))
            if mb not in (None, ""):
                try:
                    return round(float(mb) / 1024, 0)
                except Exception:
                    return mb
            return "N/A"

        def get_vm_power(vm, default_power):
            if isinstance(vm, dict):
                p = first_nested(vm, ("powerState", "power_state", "power_state_mechanism", "state"))
                if p:
                    return "ON" if str(p).upper() in ("ON", "POWERED_ON", "ACPI_ON", "NORMAL") else "OFF"
            return default_power

        cvms = []
        for n in nodes:
            node_uuid = n.get("extId") or n.get("uuid") or n.get("nodeUuid")
            cvm_ip = (n.get("controllerVm", {})
                        .get("externalAddress", {})
                        .get("ipv4", {}).get("value", "N/A"))
            host_ip = (n.get("hypervisor", {})
                        .get("externalAddress", {})
                        .get("ipv4", {}).get("value", "N/A"))

            inferred_name = vm_name_from_key(node_uuid, cvm_ip)
            matched_vm = pe_by_ip.get(str(cvm_ip)) or (pe_by_name.get(str(inferred_name).lower()) if inferred_name != "N/A" else None)

            # Prefer the authoritative name returned by the Cluster Management
            # CVM API. The fallback name inferred from SSH keys or node metadata
            # may use the internal Linux hostname instead of the configured CVM
            # display name shown in Prism.
            if matched_vm:
                api_cvm_name = str(
                    matched_vm.get("name")
                    or matched_vm.get("vm_name")
                    or matched_vm.get("vmName")
                    or ""
                ).strip()
                if api_cvm_name:
                    inferred_name = api_cvm_name

            acro_state = n.get("hypervisor", {}).get("acropolisConnectionState", "CONNECTED")
            default_power = "ON" if acro_state == "CONNECTED" else "OFF"
            cvm_power = get_vm_power(matched_vm, default_power)
            cvm_in_maint = n.get("controllerVm", {}).get("isInMaintenanceMode", False)
            cvm_status = "Critical" if (cvm_power == "OFF" or cvm_in_maint) else "Normal"

            cvms.append({
                "cvm_name": inferred_name,
                "host_name": n.get("hostName", "Unknown"),
                "cvm_ip": cvm_ip,
                "host_ip": host_ip,
                "cvm_memory_gib": get_vm_memory_gib(matched_vm),
                "cvm_vcpus": get_vm_vcpus(matched_vm),
                "cvm_power": cvm_power,
                "cvm_status": cvm_status,
            })
        critical_cvms = [c for c in cvms if c["cvm_status"] == "Critical"]
        return {
            "cvms": cvms,
            "status": self.STATUS_CRITICAL if critical_cvms else self.STATUS_HEALTHY,
            "critical_count": len(critical_cvms),
        }

    def analyse_protection(self) -> dict:
        policies = self._safe_list("protection_policies")
        recovery_plans = self._safe_list("recovery_plans")
        cluster = self._cluster_first()
        cluster_ext_id = str(cluster.get("extId") or cluster.get("uuid") or "")

        catalog = self._safe_list("cluster_catalog")
        cluster_names = {
            str(item.get("extId") or item.get("uuid")): str(item.get("name") or "")
            for item in catalog if isinstance(item, dict)
        }
        if cluster_ext_id:
            cluster_names[cluster_ext_id] = self.cluster_name

        def _cluster_name(ext_id):
            ext_id = str(ext_id or "")
            if not ext_id:
                return "All clusters"
            return cluster_names.get(ext_id) or f"{ext_id[:8]}…"

        def _replication_cluster_ids(location):
            if not isinstance(location, dict):
                return []
            sublocation = location.get("replicationSubLocation") or {}
            values = sublocation.get("clusterExtIds") or []
            return [str(value) for value in values if value]

        def _location_display(location):
            ids = _replication_cluster_ids(location)
            return ", ".join(_cluster_name(value) for value in ids) if ids else "All clusters"

        def _format_rpo(seconds):
            if not isinstance(seconds, (int, float)):
                return "N/A"
            seconds = int(seconds)
            if seconds == 0:
                return "Synchronous"
            if seconds % 86400 == 0:
                value, unit = seconds // 86400, "day"
            elif seconds % 3600 == 0:
                value, unit = seconds // 3600, "hour"
            elif seconds % 60 == 0:
                value, unit = seconds // 60, "minute"
            else:
                return f"{seconds} seconds"
            return f"{value} {unit}{'' if value == 1 else 's'}"

        def _format_retention(retention):
            if not isinstance(retention, dict) or not retention:
                return "N/A"
            object_type = str(retention.get("$objectType") or "")
            parts = []
            for label, key in [("Local", "local"), ("Remote", "remote")]:
                value = retention.get(key)
                if isinstance(value, dict):
                    interval = str(value.get("snapshotIntervalType") or "").replace("_", " ").title()
                    frequency = value.get("frequency")
                    detail = " ".join(str(v) for v in [interval, f"x{frequency}" if frequency is not None else ""] if v)
                elif isinstance(value, (int, float)):
                    detail = f"{int(value)} snapshots"
                else:
                    detail = ""
                if detail:
                    parts.append(f"{label}: {detail}")
            if parts:
                return "; ".join(parts)
            return object_type.rsplit(".", 1)[-1].replace("Retention", "") or "Configured"

        policy_rows = []
        schedule_rows = []
        paused_schedule_count = 0

        for policy in policies:
            if not isinstance(policy, dict):
                continue
            locations = policy.get("replicationLocations") or []
            label_map = {
                str(location.get("label")): location
                for location in locations
                if isinstance(location, dict) and location.get("label")
            }
            current_labels = {
                label for label, location in label_map.items()
                if not _replication_cluster_ids(location) or cluster_ext_id in _replication_cluster_ids(location)
            }
            if cluster_ext_id and label_map and not current_labels:
                continue

            roles = {
                "Primary" if location.get("isPrimary") is True else "Recovery"
                for label, location in label_map.items() if label in current_labels
            }
            configurations = policy.get("replicationConfigurations") or []
            relevant_configurations = []
            policy_paused = 0
            for configuration in configurations:
                if not isinstance(configuration, dict):
                    continue
                source_label = str(configuration.get("sourceLocationLabel") or "")
                remote_label = str(configuration.get("remoteLocationLabel") or "")
                if current_labels and source_label not in current_labels and remote_label not in current_labels:
                    continue
                schedule = configuration.get("schedule") or {}
                paused = schedule.get("isReplicationPaused") is True
                if paused:
                    paused_schedule_count += 1
                    policy_paused += 1
                source = _location_display(label_map.get(source_label, {}))
                target = _location_display(label_map.get(remote_label, {}))
                schedule_rows.append({
                    "policy": policy.get("name") or "Unnamed",
                    "direction": f"{source} → {target}",
                    "rpo": _format_rpo(schedule.get("recoveryPointObjectiveTimeSeconds")),
                    "retention": _format_retention(schedule.get("retention")),
                    "recovery_point_type": str(schedule.get("recoveryPointType") or "N/A").replace("_", " ").title(),
                    "status": self.STATUS_CRITICAL if paused else self.STATUS_HEALTHY,
                })
                relevant_configurations.append(configuration)

            policy_rows.append({
                "name": policy.get("name") or "Unnamed",
                "role": " / ".join(sorted(roles)) if roles else "Applicable",
                "category_count": len(policy.get("categoryIds") or []),
                "schedule_count": len(relevant_configurations),
                "paused_count": policy_paused,
                "status": self.STATUS_CRITICAL if policy_paused else self.STATUS_HEALTHY,
            })

        def _dr_location(location):
            if not isinstance(location, dict):
                return "N/A", []
            refs = location.get("clusters") or []
            ids = []
            names = []
            for ref in refs:
                if not isinstance(ref, dict):
                    continue
                ext_id = str(ref.get("extId") or ref.get("uuid") or "")
                if ext_id:
                    ids.append(ext_id)
                names.append(str(ref.get("name") or _cluster_name(ext_id)))
            return (", ".join(value for value in names if value) or "Remote domain manager"), ids

        recovery_plan_rows = []
        for plan in recovery_plans:
            if not isinstance(plan, dict):
                continue

            if plan.get("_api_generation") == "v3" or (
                isinstance(plan.get("metadata"), dict)
                and plan.get("metadata", {}).get("kind") == "recovery_plan"
            ):
                status_block = plan.get("status") if isinstance(plan.get("status"), dict) else {}
                spec_block = plan.get("spec") if isinstance(plan.get("spec"), dict) else {}
                resources = status_block.get("resources") or spec_block.get("resources") or {}
                parameters = resources.get("parameters") or {}
                locations = parameters.get("availability_zone_list") or []

                location_names = []
                location_ids = []
                for location in locations:
                    if not isinstance(location, dict):
                        location_names.append("N/A")
                        location_ids.append([])
                        continue
                    refs = location.get("cluster_reference_list") or []
                    names = []
                    ids = []
                    for ref in refs:
                        if not isinstance(ref, dict):
                            continue
                        ext_id = str(ref.get("uuid") or ref.get("extId") or "")
                        if ext_id:
                            ids.append(ext_id)
                        names.append(str(ref.get("name") or _cluster_name(ext_id)))
                    location_names.append(", ".join(value for value in names if value) or "N/A")
                    location_ids.append(ids)

                referenced_ids = {
                    ext_id for ids in location_ids for ext_id in ids if ext_id
                }
                if cluster_ext_id and referenced_ids and cluster_ext_id not in referenced_ids:
                    continue

                primary_index = parameters.get("primary_location_index")
                if not isinstance(primary_index, int) or not 0 <= primary_index < len(location_names):
                    primary_index = 0 if location_names else -1
                primary_name = location_names[primary_index] if primary_index >= 0 else "N/A"
                recovery_names = [
                    name for index, name in enumerate(location_names)
                    if index != primary_index and name != "N/A"
                ]

                mapping_count = 0
                for mapping in parameters.get("network_mapping_list") or []:
                    if not isinstance(mapping, dict):
                        continue
                    zone_mappings = mapping.get("availability_zone_network_mapping_list") or []
                    mapping_count += len(zone_mappings) if isinstance(zone_mappings, list) else 0

                plan_state = str(status_block.get("state") or "N/A").upper()
                if plan_state == "COMPLETE":
                    plan_status = self.STATUS_HEALTHY
                elif plan_state in {"ERROR", "FAILED"}:
                    plan_status = self.STATUS_CRITICAL
                else:
                    plan_status = self.STATUS_RECOMMENDED

                recovery_plan_rows.append({
                    "name": status_block.get("name") or spec_block.get("name") or "Unnamed",
                    "primary_location": primary_name,
                    "recovery_location": ", ".join(recovery_names) or "N/A",
                    "stage_count": len(resources.get("stage_list") or []),
                    "network_mapping_count": mapping_count,
                    "status": plan_status,
                })
                continue

            primary_name, primary_ids = _dr_location(plan.get("primaryLocation"))
            recovery_name, recovery_ids = _dr_location(plan.get("recoveryLocation"))
            referenced_ids = set(primary_ids + recovery_ids)
            if cluster_ext_id and referenced_ids and cluster_ext_id not in referenced_ids:
                continue
            recovery_plan_rows.append({
                "name": plan.get("name") or "Unnamed",
                "primary_location": primary_name,
                "recovery_location": recovery_name,
                "stage_count": len(plan.get("_stages") or []),
                "network_mapping_count": len(plan.get("_network_mappings") or []),
                "status": self.STATUS_HEALTHY,
            })

        if paused_schedule_count:
            status = self.STATUS_CRITICAL
            recs = ["Resume paused protection-policy replication after resolving the underlying issue."]
        elif not policy_rows:
            status = self.STATUS_RECOMMENDED
            recs = ["Configure or assign a Prism Central Protection Policy for workloads that require recovery protection."]
        else:
            status = self.STATUS_HEALTHY
            recs = [
                "Confirm protected categories include all critical workloads.",
                "Verify RPO and retention settings align with business recovery requirements.",
            ]

        return {
            "status": status,
            "policy_count": len(policy_rows),
            "policies": policy_rows,
            "schedules": schedule_rows,
            "schedule_count": len(schedule_rows),
            "paused_schedule_count": paused_schedule_count,
            "recovery_plan_count": len(recovery_plan_rows),
            "recovery_plans": recovery_plan_rows,
            "recommendations": recs,
        }

    def _find_numeric_values(self, obj: Any, wanted_keys: set[str]) -> list[float]:
        """Recursively find Nutanix numeric stat values.

        Handles both direct key/value formats, for example:
          {"hypervisor_cpu_usage_ppm": 12345}

        and time-series formats commonly returned by Prism Element stats, for example:
          {"stats_specific_responses": [
              {"metric": "hypervisor_cpu_usage_ppm", "values": [[ts, 12345], ...]}
          ]}
        """
        def norm(x: Any) -> str:
            return str(x).replace("-", "_").lower()

        wanted_norm = {norm(k) for k in wanted_keys}

        def add_numeric(v: Any, found: list[float]) -> None:
            if v is None:
                return
            if isinstance(v, list):
                for point in v:
                    # Prism Element time-series values can be [timestamp, value]
                    # or objects containing value/avg/average.
                    if isinstance(point, (list, tuple)) and point:
                        add_numeric(point[-1], found)
                    elif isinstance(point, dict):
                        for vk in ("value", "avg", "average", "last", "max", "min"):
                            if vk in point:
                                add_numeric(point.get(vk), found)
                    else:
                        add_numeric(point, found)
                return
            try:
                found.append(float(v))
            except (TypeError, ValueError):
                pass

        found: list[float] = []

        if isinstance(obj, dict):
            # Handle PE stats time-series objects where the metric name is stored
            # as data and the actual numbers are in values/series fields.
            metric_name = obj.get("metric") or obj.get("name") or obj.get("metricName")
            if metric_name is not None and norm(metric_name) in wanted_norm:
                for value_key in (
                    "values", "data", "series", "timeSeries", "time_series", "stats",
                    "value", "avg", "average", "last", "max", "min"
                ):
                    if value_key in obj:
                        add_numeric(obj.get(value_key), found)

            for k, v in obj.items():
                if norm(k) in wanted_norm:
                    add_numeric(v, found)
                found.extend(self._find_numeric_values(v, wanted_keys))

        elif isinstance(obj, list):
            for item in obj:
                found.extend(self._find_numeric_values(item, wanted_keys))

        return found

    def _ppm_to_pct(self, raw_stats: Any, key: str) -> Optional[float]:
        """Extract ppm stat and convert to percentage from v4, host stats, or PE v2."""
        if not isinstance(raw_stats, dict):
            return None

        cpu_keys = {
            "cpuUsagePpm", "hypervisorCpuUsagePpm", "hypervisor_cpu_usage_ppm",
            "controllerVmCpuUsagePpm", "controller_vm_cpu_usage_ppm",
            "cpu_usage_ppm", "aggregateCpuUsagePpm", "aggregate_cpu_usage_ppm",
        }
        mem_keys = {
            "memoryUsagePpm", "hypervisorMemoryUsagePpm", "hypervisor_memory_usage_ppm",
            "controllerVmMemoryUsagePpm", "controller_vm_memory_usage_ppm",
            "memory_usage_ppm", "aggregateMemoryUsagePpm", "aggregate_memory_usage_ppm",
            "aggregateHypervisorMemoryUsagePpm", "aggregate_hypervisor_memory_usage_ppm",
            "overallMemoryUsagePpm", "overall_memory_usage_ppm",
        }
        wanted = cpu_keys if key == "cpuUsagePpm" else mem_keys

        # Fast path: v4 stats return time-series lists [{timestamp, value}, ...]
        # newest-first. Grab index [0] for the most recent sample directly.
        data = raw_stats if isinstance(raw_stats, dict) else {}
        data = data.get("data", data)
        if isinstance(data, list) and data:
            data = data[0]
        if isinstance(data, dict):
            for k in wanted:
                v = data.get(k)
                if v is None:
                    continue
                if isinstance(v, list) and v:
                    item = v[0]
                    raw_val = item.get("value") if isinstance(item, dict) else item
                elif isinstance(v, (int, float, str)):
                    raw_val = v
                else:
                    continue
                try:
                    fval = float(raw_val)
                    if fval > 1000:           # ppm range → convert
                        return round(fval / 10000, 2)
                    elif 0 <= fval <= 100:    # already a percentage
                        return round(fval, 2)
                except (TypeError, ValueError):
                    continue

        # Fallback: recursive search (handles PE v2 and legacy shapes)
        values = self._find_numeric_values(raw_stats, wanted)
        if not values:
            return None
        percentages = [v / 10000 if v > 1000 else v for v in values if 0 <= v <= 1_000_000]
        return round(percentages[0], 2) if percentages else None

    def _extract_time_series_pct(self, raw_stats: Any, wanted_keys: set[str]) -> list[dict]:
        """Extract timestamp/value time-series from PC v4 stats and convert PPM to percent."""
        def norm(x: Any) -> str:
            return str(x).replace("-", "_").lower()

        wanted_norm = {norm(k) for k in wanted_keys}
        points: list[dict] = []

        def walk(obj: Any) -> None:
            if isinstance(obj, dict):
                # PC v4 format: data.hypervisorCpuUsagePpm = [{timestamp, value}, ...]
                for k, v in obj.items():
                    if norm(k) in wanted_norm and isinstance(v, list):
                        for item in v:
                            if isinstance(item, dict) and "timestamp" in item and "value" in item:
                                try:
                                    val = float(item["value"])
                                    pct = val / 10000 if val > 1000 else val
                                    points.append({"timestamp": str(item["timestamp"]), "value": round(pct, 2)})
                                except (TypeError, ValueError):
                                    pass
                    walk(v)
            elif isinstance(obj, list):
                for item in obj:
                    walk(item)

        walk(raw_stats)

        # De-dupe and sort oldest to newest
        by_ts = {}
        for p in points:
            by_ts[p["timestamp"]] = p
        return [by_ts[ts] for ts in sorted(by_ts)]

    def _downsample_time_series(self, points: list[dict], max_points: int = 120) -> list[dict]:
        """Reduce dense time-series to a readable number of points for the report graph."""
        if len(points) <= max_points:
            return points
        bucket_size = max(1, len(points) // max_points)
        sampled = []
        for i in range(0, len(points), bucket_size):
            bucket = points[i:i + bucket_size]
            vals = [p["value"] for p in bucket if isinstance(p.get("value"), (int, float))]
            if not vals:
                continue
            sampled.append({"timestamp": bucket[-1]["timestamp"], "value": round(sum(vals) / len(vals), 2)})
        return sampled[-max_points:]

    def analyse_cpu(self) -> dict:
        usage = self._ppm_to_pct(self.raw.get("cluster_stats"), "cpuUsagePpm")
        cpu_history = self._extract_time_series_pct(
            self.raw.get("cluster_stats_7d") or self.raw.get("cluster_stats"),
            {"hypervisorCpuUsagePpm", "hypervisor_cpu_usage_ppm"}
        )
        cpu_history = self._downsample_time_series(cpu_history, max_points=120)

        history_vals = [p.get("value") for p in cpu_history if isinstance(p.get("value"), (int, float))]
        avg_from_history = round(sum(history_vals) / len(history_vals), 2) if history_vals else None
        peak_from_history = round(max(history_vals), 2) if history_vals else None
        avg_usage = usage if usage is not None else avg_from_history
        peak_usage = peak_from_history if peak_from_history is not None else usage
        headroom_pct = round(100 - avg_usage, 2) if isinstance(avg_usage, (int, float)) else "N/A"

        nodes = self._safe_list("nodes")
        vms = self._safe_list("virtual_machines")
        cvms = self.analyse_cvms().get("cvms", [])
        host_by_id = {n.get("extId") or n.get("uuid"): n for n in nodes}

        total_sockets = sum(int(n.get("numberOfCpuSockets") or 0) for n in nodes)
        physical_cores = sum(int(n.get("numberOfCpuCores") or 0) for n in nodes)
        logical_cpus = sum(int(n.get("numberOfCpuThreads") or 0) for n in nodes)
        total_user_vm_vcpus = 0
        user_vcpu_by_host = {}
        for vm in vms:
            vcpus = int(vm.get("numSockets") or 0) * int(vm.get("numCoresPerSocket") or 1)
            total_user_vm_vcpus += vcpus
            h = vm.get("host", {}) if isinstance(vm.get("host"), dict) else {}
            host_id = h.get("extId") or h.get("uuid") or "Unknown"
            user_vcpu_by_host[host_id] = user_vcpu_by_host.get(host_id, 0) + vcpus

        cvm_vcpu_by_host_name = {}
        total_cvm_vcpus = 0
        for cvm in cvms:
            try:
                cvm_vcpus = int(float(cvm.get("cvm_vcpus") or 0))
            except (TypeError, ValueError):
                cvm_vcpus = 0
            host_name = str(cvm.get("host_name") or "Unknown")
            cvm_vcpu_by_host_name[host_name] = cvm_vcpu_by_host_name.get(host_name, 0) + cvm_vcpus
            total_cvm_vcpus += cvm_vcpus

        total_vcpus = total_user_vm_vcpus + total_cvm_vcpus

        vcpu_pcore_ratio = round(total_vcpus / physical_cores, 2) if physical_cores else None
        ratio_display = f"{vcpu_pcore_ratio:.2f} : 1" if vcpu_pcore_ratio is not None else "N/A"
        if vcpu_pcore_ratio is None:
            ratio_status = self.STATUS_HEALTHY
        elif vcpu_pcore_ratio > 4:
            ratio_status = self.STATUS_CRITICAL
        elif vcpu_pcore_ratio > 2:
            ratio_status = self.STATUS_RECOMMENDED
        else:
            ratio_status = self.STATUS_HEALTHY

        host_cpu_distribution = []
        for n in sorted(nodes, key=lambda x: str(x.get("hostName", x.get("name", ""))).lower()):
            host_id = n.get("extId") or n.get("uuid")
            host_name = n.get("hostName") or n.get("name") or host_id or "Unknown"
            host_cores = int(n.get("numberOfCpuCores") or 0)
            user_vm_vcpus = int(user_vcpu_by_host.get(host_id, 0))
            cvm_vcpus = int(cvm_vcpu_by_host_name.get(str(host_name), 0))
            host_vcpus = user_vm_vcpus + cvm_vcpus
            host_ratio = round(host_vcpus / host_cores, 2) if host_cores else None
            host_cpu_distribution.append({
                "host": host_name,
                "user_vm_vcpus": user_vm_vcpus,
                "cvm_vcpus": cvm_vcpus,
                "vcpus": host_vcpus,
                "physical_cores": host_cores if host_cores else "N/A",
                "vcpu_pcore_ratio": f"{host_ratio:.2f} : 1" if host_ratio is not None else "N/A",
            })

        status = self.STATUS_HEALTHY
        if (isinstance(avg_usage, (int, float)) and avg_usage > 85) or ratio_status == self.STATUS_CRITICAL:
            status = self.STATUS_CRITICAL
        elif ((isinstance(avg_usage, (int, float)) and avg_usage >= 70) or
              (isinstance(peak_usage, (int, float)) and peak_usage >= 90) or
              ratio_status == self.STATUS_RECOMMENDED):
            status = self.STATUS_RECOMMENDED

        if status == self.STATUS_CRITICAL:
            recs = ["CPU utilization or oversubscription exceeds recommended thresholds. Consider adding compute capacity or redistributing workloads."]
        elif status == self.STATUS_RECOMMENDED:
            recs = ["Monitor CPU utilization growth and review workloads contributing to elevated utilization or oversubscription."]
        else:
            recs = ["CPU utilization remained below recommended thresholds during the assessment period. No action is recommended."]

        return {
            "status":                status,
            "average_cpu_usage_pct": avg_usage if avg_usage is not None else "N/A",
            "peak_cpu_usage_pct":    peak_usage if peak_usage is not None else "N/A",
            "cpu_headroom_pct":      headroom_pct,
            "cpu_history":           cpu_history,
            "physical_hosts":        len(nodes),
            "total_cpu_sockets":     total_sockets if total_sockets else "N/A",
            "physical_cores":        physical_cores if physical_cores else "N/A",
            "logical_cpus":          logical_cpus if logical_cpus else "N/A",
            "total_user_vm_vcpus":   total_user_vm_vcpus,
            "total_cvm_vcpus":       total_cvm_vcpus,
            "total_vcpus":           total_vcpus,
            "vcpu_pcore_ratio":      ratio_display,
            "oversubscription_status": ratio_status,
            "host_cpu_distribution": host_cpu_distribution,
            "recommendations":       recs,
        }

    def _alert_text(self, alert: dict) -> str:
        """Return searchable text for a normalized active alert."""
        parts = [
            alert.get("title", ""),
            alert.get("message", ""),
            alert.get("severity", ""),
            alert.get("classification", ""),
            alert.get("impactType", ""),
            alert.get("sourceHost", ""),
            alert.get("sourceEntity", ""),
        ]
        return " ".join(str(x) for x in parts if x).lower()

    def _is_cluster_or_host_alert(self, alert: dict) -> bool:
        """Return True for cluster/host-level alerts and exclude Prism Central-only alerts."""
        source_type = str(alert.get("sourceType") or "").lower()
        source_host = str(alert.get("sourceHost") or alert.get("sourceEntity") or "").lower()

        # Section-level correlated alerts should focus on the assessed cluster and hosts.
        # Exclude Prism Central / PCVM sources so PC-level alerts do not appear in CPU,
        # memory, storage, or other subsystem summaries.
        pc_markers = ("prism central", "pcvm", "pc01", "domain manager")
        if any(marker in source_host for marker in pc_markers):
            return False

        # PC v3 typically reports cluster/host alerts as cluster, node, or host.
        # If the type is unavailable, keep the alert because it has already been
        # filtered by originating_cluster_uuid during collection.
        if source_type in ("", "cluster", "node", "host"):
            return True
        return False

    def _alerts_matching_keywords(self, keywords: list[str]) -> list[dict]:
        """Return active cluster/host alerts whose normalized text matches one or more keywords."""
        matches = []
        lowered = [k.lower() for k in keywords]
        for alert in self._safe_list("alerts"):
            if not self._is_cluster_or_host_alert(alert):
                continue
            text = self._alert_text(alert)
            if any(k in text for k in lowered):
                matches.append(alert)
        return matches

    def _correlated_status(self, base_status: str, alerts: list[dict]) -> str:
        """Escalate a section status based on correlated active alerts."""
        severities = {str(a.get("severity", "")).upper() for a in alerts}
        if severities.intersection({"CRITICAL", "ERROR", "FATAL"}):
            return self.STATUS_CRITICAL
        if "WARNING" in severities:
            return self.STATUS_RECOMMENDED if base_status == self.STATUS_HEALTHY else base_status
        return base_status

    def _alert_titles(self, alerts: list[dict], limit: int = 3) -> str:
        titles = []
        for alert in alerts:
            title = str(alert.get("title") or alert.get("message") or "Alert")
            if title not in titles:
                titles.append(title)
            if len(titles) >= limit:
                break
        unique_count = len({str(a.get("title") or a.get("message") or "Alert") for a in alerts})
        extra = unique_count - len(titles)
        suffix = f" and {extra} additional alert(s)" if extra > 0 else ""
        return "; ".join(titles) + suffix if titles else "N/A"

    def analyse_memory(self) -> dict:
        usage = self._ppm_to_pct(self.raw.get("cluster_stats"), "memoryUsagePpm")
        mem_history = self._extract_time_series_pct(
            self.raw.get("cluster_stats_7d") or self.raw.get("cluster_stats"),
            {"aggregateHypervisorMemoryUsagePpm", "aggregate_hypervisor_memory_usage_ppm",
             "hypervisorMemoryUsagePpm", "memoryUsagePpm"}
        )
        mem_history = self._downsample_time_series(mem_history, max_points=120)

        history_vals = [p.get("value") for p in mem_history if isinstance(p.get("value"), (int, float))]
        avg_from_history = round(sum(history_vals) / len(history_vals), 2) if history_vals else None
        peak_from_history = round(max(history_vals), 2) if history_vals else None
        avg_usage = usage if usage is not None else avg_from_history
        peak_usage = peak_from_history if peak_from_history is not None else usage
        headroom_pct = round(100 - avg_usage, 2) if isinstance(avg_usage, (int, float)) else "N/A"

        nodes = self._safe_list("nodes")
        vms = self._safe_list("virtual_machines")
        cvms = self.analyse_cvms().get("cvms", [])
        GiB = 1024 ** 3

        total_memory_gib = round(sum(float(n.get("memorySizeBytes") or 0) for n in nodes) / GiB, 1) if nodes else 0
        total_user_vm_memory_gib = round(sum(float(vm.get("memorySizeBytes") or 0) for vm in vms) / GiB, 1)
        total_cvm_memory_gib = 0.0
        for cvm in cvms:
            try:
                total_cvm_memory_gib += float(cvm.get("cvm_memory_gib") or 0)
            except (TypeError, ValueError):
                pass
        total_cvm_memory_gib = round(total_cvm_memory_gib, 1)
        total_vm_memory_gib = round(total_user_vm_memory_gib + total_cvm_memory_gib, 1)
        memory_allocation_pct = round((total_vm_memory_gib / total_memory_gib) * 100, 2) if total_memory_gib else None

        if memory_allocation_pct is None:
            allocation_status = self.STATUS_HEALTHY
        elif memory_allocation_pct > 100:
            allocation_status = self.STATUS_CRITICAL
        elif memory_allocation_pct >= 80:
            allocation_status = self.STATUS_RECOMMENDED
        else:
            allocation_status = self.STATUS_HEALTHY

        vm_mem_by_host: dict[str, float] = {}
        vm_count_by_host: dict[str, int] = {}
        for vm in vms:
            mem_gib = float(vm.get("memorySizeBytes") or 0) / GiB
            h = vm.get("host", {}) if isinstance(vm.get("host"), dict) else {}
            host_id = h.get("extId") or h.get("uuid") or "Unknown"
            vm_mem_by_host[host_id] = vm_mem_by_host.get(host_id, 0.0) + mem_gib
            vm_count_by_host[host_id] = vm_count_by_host.get(host_id, 0) + 1

        cvm_mem_by_host_name: dict[str, float] = {}
        for cvm in cvms:
            try:
                cvm_mem = float(cvm.get("cvm_memory_gib") or 0)
            except (TypeError, ValueError):
                cvm_mem = 0.0
            host_name = str(cvm.get("host_name") or "Unknown")
            cvm_mem_by_host_name[host_name] = cvm_mem_by_host_name.get(host_name, 0.0) + cvm_mem

        host_memory_distribution = []
        for n in sorted(nodes, key=lambda x: str(x.get("hostName", x.get("name", ""))).lower()):
            host_id = n.get("extId") or n.get("uuid")
            host_name = n.get("hostName") or n.get("name") or host_id or "Unknown"
            physical_gib = float(n.get("memorySizeBytes") or 0) / GiB
            user_vm_allocated_gib = vm_mem_by_host.get(host_id, 0.0)
            cvm_allocated_gib = cvm_mem_by_host_name.get(str(host_name), 0.0)
            allocated_gib = user_vm_allocated_gib + cvm_allocated_gib
            alloc_pct = round((allocated_gib / physical_gib) * 100, 2) if physical_gib else None
            host_memory_distribution.append({
                "host": host_name,
                "vm_count": int(vm_count_by_host.get(host_id, 0)),
                "user_vm_memory_gib": round(user_vm_allocated_gib, 1),
                "cvm_memory_gib": round(cvm_allocated_gib, 1),
                "allocated_memory_gib": round(allocated_gib, 1),
                "physical_memory_gib": round(physical_gib, 1) if physical_gib else "N/A",
                "allocation_pct": f"{alloc_pct:.2f}%" if alloc_pct is not None else "N/A",
            })

        utilization_status = self.STATUS_HEALTHY
        if (isinstance(avg_usage, (int, float)) and avg_usage > 85) or allocation_status == self.STATUS_CRITICAL:
            utilization_status = self.STATUS_CRITICAL
        elif ((isinstance(avg_usage, (int, float)) and avg_usage >= 70) or
              (isinstance(peak_usage, (int, float)) and peak_usage >= 90) or
              allocation_status == self.STATUS_RECOMMENDED):
            utilization_status = self.STATUS_RECOMMENDED

        memory_alerts = self._alerts_matching_keywords(["dimm", "cecc", "ecc", "memory", "ram"])
        status = self._correlated_status(utilization_status, memory_alerts)

        if memory_alerts:
            recs = [
                "Investigate and remediate active memory hardware alerts before treating memory health as normal.",
                f"Correlated memory alert(s): {self._alert_titles(memory_alerts)}.",
            ]
            if utilization_status != self.STATUS_HEALTHY:
                recs.append("Memory utilization or allocation also exceeds recommended thresholds. Review capacity and VM memory allocations.")
        elif status == self.STATUS_CRITICAL:
            recs = ["Memory utilization or allocation exceeds recommended thresholds. Consider adding memory capacity or rightsizing workloads."]
        elif status == self.STATUS_RECOMMENDED:
            recs = ["Monitor memory utilization growth and review VM memory allocations for right-sizing opportunities."]
        else:
            recs = ["Memory utilization remained below recommended thresholds during the assessment period. No action is recommended."]

        return {
            "status": status,
            "average_memory_usage_pct": avg_usage if avg_usage is not None else "N/A",
            "peak_memory_usage_pct": peak_usage if peak_usage is not None else "N/A",
            "memory_headroom_pct": headroom_pct,
            "mem_history": mem_history,
            "physical_hosts": len(nodes),
            "total_physical_memory_gib": total_memory_gib if total_memory_gib else "N/A",
            "total_user_vm_memory_gib": total_user_vm_memory_gib,
            "total_cvm_memory_gib": total_cvm_memory_gib,
            "total_vm_memory_gib": total_vm_memory_gib,
            "memory_allocation_pct": f"{memory_allocation_pct:.2f}%" if memory_allocation_pct is not None else "N/A",
            "memory_allocation_status": allocation_status,
            "memory_utilization_status": utilization_status,
            "memory_alert_count": len(memory_alerts),
            "memory_alerts": [
                {
                    "severity": a.get("severity", "UNKNOWN"),
                    "title": a.get("title", a.get("message", "Alert")),
                    "source": a.get("sourceType") or a.get("classification") or "N/A",
                    "host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "source_host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "last_occurred": a.get("lastOccurred") or a.get("creationTime", ""),
                }
                for a in memory_alerts[:10]
            ],
            "host_memory_distribution": host_memory_distribution,
            "recommendations": recs,
        }

    def analyse_network(self) -> dict:
        nets  = self._safe_list("networks")
        nodes = self._safe_list("nodes")
        details = self.raw.get("network_details") or {}
        cluster_data = (self.raw.get("cluster_info") or {}).get("data", {})
        cluster_ext_id = cluster_data.get("extId") or cluster_data.get("uuid") or cluster_data.get("clusterUuid")
        cluster_network = cluster_data.get("network", {}) if isinstance(cluster_data, dict) else {}

        def ip_from_obj(obj):
            if not obj:
                return "N/A"
            if isinstance(obj, str):
                return obj
            if isinstance(obj, dict):
                if obj.get("value"):
                    return obj.get("value")
                if obj.get("ipv4", {}).get("value"):
                    return obj.get("ipv4", {}).get("value")
                if obj.get("ip", {}).get("ipv4", {}).get("value"):
                    return obj.get("ip", {}).get("ipv4", {}).get("value")
                if obj.get("fqdn", {}).get("value"):
                    return obj.get("fqdn", {}).get("value")
            return "N/A"

        def first(*vals):
            for val in vals:
                if val not in (None, "", [], {}):
                    return val
            return "N/A"

        def norm_status(*vals):
            text = " ".join(str(v) for v in vals if v is not None).lower()
            if any(x in text for x in ["down", "failed", "failure", "error", "critical", "disconnected"]):
                return self.STATUS_CRITICAL
            if any(x in text for x in ["warning", "degraded", "unknown", "inactive"]):
                return self.STATUS_RECOMMENDED
            if any(x in text for x in ["up", "active", "connected", "normal", "healthy", "true"]):
                return self.STATUS_HEALTHY
            return "N/A"

        ipmi_missing = [
            n.get("hostName", n.get("uuid", "Unknown"))
            for n in nodes
            if not n.get("ipmiAddress") and not n.get("ipmi", {}).get("ip")
        ]

        network_alerts = self._alerts_matching_keywords([
            "network", "bond", "ovs", "bridge", "nic", "interface", "link", "mtu",
            "vlan", "latency", "cvm unreachable", "host unreachable", "gateway"
        ])

        # Build node IP summary.
        host_ip_summary = []
        for n in nodes:
            host_name = n.get("hostName") or n.get("name") or n.get("uuid") or "Unknown"
            host_ip = ip_from_obj(n.get("hypervisor", {}).get("externalAddress") or n.get("hostIp") or n.get("hypervisorAddress"))
            cvm_ip = ip_from_obj(n.get("controllerVm", {}).get("externalAddress") or n.get("controllerVmIp"))
            ipmi_ip = ip_from_obj(n.get("ipmi", {}).get("ip") or n.get("ipmiAddress"))
            host_ip_summary.append({
                "host": host_name,
                "ahv_ip": host_ip,
                "cvm_ip": cvm_ip,
                "ipmi_ip": ipmi_ip,
                "status": n.get("nodeStatus") or n.get("maintenanceState") or "N/A",
                "uuid": n.get("extId") or n.get("uuid") or n.get("nodeUuid"),
            })

        host_name_by_uuid = {str(n.get("uuid")): n.get("host") for n in host_ip_summary if n.get("uuid")}

        # VLAN / bridge / vSwitch details from subnet inventory.
        networks = []
        bridge_map = {}
        vswitch_map = {}

        def ip_assignment_service(subnet):
            """Return Nutanix IPAM or External IPAM based on subnet/IPAM fields.

            Prism Central releases expose IPAM settings with slightly different
            names. If Nutanix-managed IPAM/DHCP fields are present, report
            Nutanix IPAM; otherwise report External IPAM.
            """
            keys = {str(k).lower() for k in subnet.keys()} if isinstance(subnet, dict) else set()
            text = json.dumps(subnet, default=str).lower() if isinstance(subnet, dict) else ""
            explicit = first(
                subnet.get("ipAssignmentService") if isinstance(subnet, dict) else None,
                subnet.get("ipamType") if isinstance(subnet, dict) else None,
                subnet.get("ipam") if isinstance(subnet, dict) else None,
                subnet.get("ipAssignmentType") if isinstance(subnet, dict) else None,
                subnet.get("ipAssignmentMode") if isinstance(subnet, dict) else None,
            )
            explicit_text = str(explicit).lower()
            if any(x in explicit_text for x in ["nutanix", "internal", "managed"]):
                return "Nutanix IPAM"
            if any(x in explicit_text for x in ["external", "none", "unmanaged"]):
                return "External IPAM"
            if any(k in keys for k in ["ipconfig", "ipaddresspoollist", "ippools", "dhcpoptions", "poollist", "reservedipaddresses"]):
                return "Nutanix IPAM"
            if any(x in text for x in ["dhcpserveraddress", "ipaddresspools", "defaultgatewayip", "poollist"]):
                return "Nutanix IPAM"
            return "External IPAM"

        for n in nets[:100]:
            vlan = ("Untagged" if n.get("networkId") == 0
                    else str(n.get("networkId")) if n.get("networkId") is not None
                    else str(n.get("vlanId") or n.get("vlan") or "N/A"))
            bridge = n.get("bridgeName") or n.get("bridge") or "N/A"
            vswitch = n.get("virtualSwitchReference") or n.get("virtualSwitchExtId") or n.get("virtualSwitch") or "N/A"
            name = n.get("name", n.get("subnetName", "Unknown"))
            networks.append({"name": name, "vlan_id": vlan, "bridge": bridge, "virtual_switch": vswitch, "ip_assignment": ip_assignment_service(n)})
            if bridge != "N/A":
                bridge_map.setdefault(bridge, {"bridge": bridge, "virtual_switch": vswitch, "vlans": [], "status": self.STATUS_HEALTHY})["vlans"].append(vlan)
            if vswitch != "N/A":
                vswitch_map.setdefault(str(vswitch), {"virtual_switch": str(vswitch), "bridges": set(), "vlans": []})
                vswitch_map[str(vswitch)]["bridges"].add(bridge)
                vswitch_map[str(vswitch)]["vlans"].append(vlan)

        bridge_summary = list(bridge_map.values())
        for b in bridge_summary:
            b["vlans"] = ", ".join(str(v) for v in sorted(set(b["vlans"]), key=str))

        virtual_switch_summary = []
        raw_vswitches_all = details.get("virtual_switches") or []

        def vs_clusters_for_current(vswitch):
            clusters = vswitch.get("clusters") or []
            if not cluster_ext_id:
                return clusters
            return [c for c in clusters if str(c.get("extId") or c.get("clusterExtId") or "") == str(cluster_ext_id)]

        raw_vswitches = [v for v in raw_vswitches_all if vs_clusters_for_current(v)] or raw_vswitches_all
        if raw_vswitches:
            for v in raw_vswitches[:20]:
                name = first(v.get("name"), v.get("virtualSwitchName"), v.get("extId"), v.get("uuid"))
                mode = first(v.get("bondMode"), v.get("mode"), v.get("uplinkMode"), v.get("loadBalancingMode"))
                status = norm_status(v.get("status"), v.get("state"), v.get("health"), v.get("isActive"))
                virtual_switch_summary.append({"name": name, "mode": mode, "status": status, "extId": v.get("extId")})
        else:
            for vs, data in vswitch_map.items():
                virtual_switch_summary.append({
                    "name": vs,
                    "mode": "N/A",
                    "status": self.STATUS_HEALTHY,
                    "bridges": ", ".join(sorted(set(str(x) for x in data["bridges"] if x != "N/A"))) or "N/A",
                })

        def host_name_from_id(value):
            value = str(value or "")
            if value in host_name_by_uuid:
                return host_name_by_uuid[value]
            for n in nodes:
                if value and value in json.dumps(n, default=str):
                    return n.get("hostName") or n.get("name") or value
            return value or "N/A"

        # Extract bond/uplink membership from Prism Central virtual switch inventory.
        # PC exposes the virtual switch bond mode and the host NICs assigned to the
        # switch even when it does not expose physical link-state or speed.
        for v in raw_vswitches:
            vs_name = first(v.get("name"), v.get("virtualSwitchName"), v.get("extId"), v.get("uuid"))
            vs_ext_id = first(v.get("extId"), v.get("uuid"))
            mode = first(v.get("bondMode"), v.get("mode"), v.get("uplinkMode"), v.get("loadBalancingMode"))
            for c in vs_clusters_for_current(v):
                for h in c.get("hosts") or []:
                    host_uuid = h.get("extId") or h.get("hostExtId") or h.get("hostUuid") or h.get("nodeUuid")
                    host = host_name_from_id(host_uuid)
                    bridge = h.get("internalBridgeName") or h.get("bridgeName") or "N/A"
                    members = h.get("hostNics") or h.get("interfaces") or h.get("uplinks") or []
                    if isinstance(members, list):
                        member_text = ", ".join(str(m.get("name") if isinstance(m, dict) else m) for m in members)
                    else:
                        member_text = str(members)
                    if member_text and member_text != "[]":
                        bond_row = {
                            "hostName": host,
                            "name": vs_name,
                            "mode": mode,
                            "memberInterfaces": member_text,
                            "status": self.STATUS_HEALTHY,
                            "bridge": bridge,
                            "virtualSwitchExtId": vs_ext_id,
                        }
                        if bond_row not in details["bonds"]:
                            details["bonds"].append(bond_row)
                    for member in (members if isinstance(members, list) else []):
                        nic_name = str(member.get("name") if isinstance(member, dict) else member)
                        existing_nic = any(
                            str(x.get("hostName") or "") == str(host) and
                            str(x.get("name") or x.get("interfaceName") or "") == str(nic_name)
                            for x in details.get("host_nics") or []
                            if isinstance(x, dict)
                        )
                        if not existing_nic:
                            nic_row = {
                                "hostName": host,
                                "name": nic_name,
                                "interfaceName": nic_name,
                                "linkStatus": "Not available from PC",
                                "speed": "N/A",
                                "status": "unknown",
                                "virtualSwitchExtId": vs_ext_id,
                            }
                            if nic_row not in details["host_nics"]:
                                details["host_nics"].append(nic_row)

        # Extract interface/bond-like objects from action payloads if PC returned
        # nested networking details rather than clean inventory lists.
        def walk(obj):
            if isinstance(obj, dict):
                yield obj
                for val in obj.values():
                    yield from walk(val)
            elif isinstance(obj, list):
                for item in obj:
                    yield from walk(item)

        def lower_keys(obj):
            return {str(k).lower() for k in obj.keys()} if isinstance(obj, dict) else set()

        for payload in details.get("raw_networking") or []:
            for item in walk(payload):
                keys = lower_keys(item)
                joined = " ".join(keys)
                if any(k in keys for k in ["macaddress", "mac", "linkstatus", "linkstate", "speedmbps", "speedinmbps", "interfacename", "devicename"]):
                    if item not in details["host_nics"]:
                        details["host_nics"].append(item)
                if any(x in joined for x in ["bond", "uplink", "lacp", "slaves", "slaveinterfaces", "memberinterfaces"]):
                    if item not in details["bonds"]:
                        details["bonds"].append(item)

        def format_speed_from_kbps(value):
            try:
                if value in (None, "", "N/A"):
                    return "—"
                kbps = float(value)
                if kbps <= 0:
                    return "—"
                gbps = kbps / 1000000.0
                if gbps >= 1:
                    return f"{gbps:g} Gbps"
                mbps = kbps / 1000.0
                return f"{mbps:g} Mbps"
            except Exception:
                return str(value or "—")

        def format_capacity_mbps(value):
            try:
                if value in (None, "", "N/A"):
                    return "N/A"
                mbps = float(value)
                if mbps >= 1000:
                    return f"{mbps / 1000.0:g} Gbps"
                return f"{mbps:g} Mbps"
            except Exception:
                return str(value or "N/A")

        def link_from_interface_status(value):
            text = str(value).strip().lower()
            if text in {"1", "true", "up", "active", "connected"}:
                return "Up"
            if text in {"0", "false", "down", "inactive", "disconnected"}:
                return "Down"
            return str(value) if value not in (None, "") else "N/A"

        # NIC Summary from Prism Central host-nics API and other best-effort sources.
        nic_summary = []
        seen_nics = set()
        for nic in (details.get("host_nics") or [])[:120]:
            if not isinstance(nic, dict):
                continue
            host = host_name_from_id(first(nic.get("hostName"), nic.get("hostExtId"), nic.get("hostUuid"), nic.get("nodeUuid"), nic.get("hostId")))
            name = first(nic.get("name"), nic.get("interfaceName"), nic.get("deviceName"), nic.get("macAddress"), nic.get("uuid"), nic.get("extId"))
            mac = first(nic.get("macAddress"), nic.get("mac"))
            raw_status = first(nic.get("interfaceStatus"), nic.get("linkStatus"), nic.get("linkState"), nic.get("state"), nic.get("status"), nic.get("isLinkUp"))
            link = link_from_interface_status(raw_status)
            speed = format_speed_from_kbps(first(nic.get("linkSpeedInKbps"), nic.get("speedKbps"), nic.get("speedInKbps")))
            if speed == "—":
                speed = format_capacity_mbps(first(nic.get("speedMbps"), nic.get("speed"), nic.get("linkSpeed"), nic.get("speedInMbps"))) if any(k in nic for k in ["speedMbps", "speed", "linkSpeed", "speedInMbps"]) else "—"
            capacity = format_capacity_mbps(first(nic.get("linkCapacityInMbps"), nic.get("capacityMbps"), nic.get("maxSpeedMbps")))
            mtu = first(nic.get("mtuInBytes"), nic.get("mtu"), nic.get("mtuBytes"))
            adapter = first(nic.get("hostDescription"), nic.get("description"), nic.get("adapter"))
            driver = first(nic.get("driverVersion"), nic.get("driver"))
            firmware = first(nic.get("firmwareVersion"), nic.get("firmware"))
            virtual_switch = first(nic.get("virtualSwitchExtId"), nic.get("virtualSwitch"), nic.get("virtualSwitchName"))
            status = self.STATUS_HEALTHY if link == "Up" else (self.STATUS_RECOMMENDED if link == "Down" else norm_status(link, nic.get("health"), nic.get("status")))
            # Bond membership is applied after bond_summary is built. Until then,
            # this is only the raw link-derived status.
            # Skip PE host objects that do not actually look like NIC/interface rows.
            if name == "N/A" and mac == "N/A" and speed in ("N/A", "—"):
                continue
            key = (str(host), str(name), str(mac))
            if key in seen_nics:
                continue
            seen_nics.add(key)
            nic_summary.append({
                "host": host,
                "interface": name,
                "link": link,
                "speed": speed,
                "capacity": capacity,
                "mac": mac,
                "mtu": str(mtu),
                "adapter": adapter,
                "driver": driver,
                "firmware": firmware,
                "virtual_switch": virtual_switch,
                "status": status,
            })

        # Bond Summary from explicit bond inventory when available, otherwise infer
        # a bridge/vSwitch row so the report still documents resiliency objects.
        bond_summary = []
        for bond in (details.get("bonds") or [])[:40]:
            host = host_name_from_id(first(bond.get("hostName"), bond.get("hostExtId"), bond.get("hostUuid"), bond.get("nodeUuid"), bond.get("hostId")))
            name = first(bond.get("name"), bond.get("bondName"), bond.get("interfaceName"), bond.get("extId"), bond.get("uuid"))
            mode = first(bond.get("mode"), bond.get("bondMode"), bond.get("lacpMode"), bond.get("loadBalancingMode"))
            members = first(bond.get("memberInterfaces"), bond.get("interfaces"), bond.get("slaveInterfaces"), bond.get("uplinks"))
            if isinstance(members, list):
                members = ", ".join(str(m.get("name") if isinstance(m, dict) else m) for m in members)
            status = norm_status(bond.get("status"), bond.get("state"), bond.get("health"), bond.get("isActive"))
            bond_summary.append({"host": host, "bond": name, "mode": mode, "members": members, "status": status})

        def eth_sort_key(name):
            text = str(name or "")
            m = re.search(r"(\d+)$", text)
            return (re.sub(r"\d+$", "", text), int(m.group(1)) if m else 9999, text)

        # Remove virtual-switch placeholder NIC rows when the host-nics API
        # returned authoritative data for the same host/interface. This prevents
        # rows with Link = "Not available from PC", Speed/MAC/MTU = N/A from
        # replacing or duplicating real physical NIC inventory.
        real_nic_keys = set()
        for n in nic_summary:
            if str(n.get("link") or "").lower() != "not available from pc" and (n.get("mac") not in (None, "", "N/A")):
                real_nic_keys.add((str(n.get("host") or ""), str(n.get("interface") or "")))
        if real_nic_keys:
            nic_summary = [
                n for n in nic_summary
                if not (
                    (str(n.get("host") or ""), str(n.get("interface") or "")) in real_nic_keys
                    and str(n.get("link") or "").lower() == "not available from pc"
                )
            ]

        # Build a host/interface membership set so unused NICs can be ignored,
        # while down NICs that are actual bond members are treated as critical.
        bonded_members = set()
        for bond in bond_summary:
            host = str(bond.get("host") or "")
            for member in re.split(r"[,\s]+", str(bond.get("members") or "")):
                member = member.strip()
                if member:
                    bonded_members.add((host, member))

        for nic in nic_summary:
            member_key = (str(nic.get("host") or ""), str(nic.get("interface") or ""))
            in_bond = member_key in bonded_members
            nic["in_bond"] = in_bond
            if not in_bond:
                nic["status"] = "Ignored"
            elif nic.get("link") == "Up":
                nic["status"] = self.STATUS_HEALTHY
            elif nic.get("link") == "Down":
                nic["status"] = self.STATUS_CRITICAL
            else:
                nic["status"] = self.STATUS_RECOMMENDED

        nic_summary.sort(key=lambda n: (str(n.get("host") or ""), eth_sort_key(n.get("interface"))))
        bond_summary.sort(key=lambda b: (str(b.get("host") or ""), str(b.get("bond") or "")))
        networks.sort(key=lambda n: (str(n.get("name") or ""), str(n.get("vlan_id") or "")))

        configured_dns = [ip_from_obj(x) for x in cluster_network.get("nameServerIpList", []) or []]
        configured_ntp = [ip_from_obj(x) for x in cluster_network.get("ntpServerIpList", []) or []]
        external_subnet = cluster_network.get("externalSubnet") or "N/A"
        internal_subnet = cluster_network.get("internalSubnet") or "N/A"
        cluster_vip = ip_from_obj(cluster_network.get("externalAddress"))
        data_services_ip = ip_from_obj(cluster_network.get("externalDataServiceIp"))

        dns_status = self.STATUS_HEALTHY if configured_dns else self.STATUS_RECOMMENDED
        ntp_status = self.STATUS_HEALTHY if configured_ntp else self.STATUS_RECOMMENDED
        ipmi_status = self.STATUS_RECOMMENDED if ipmi_missing else self.STATUS_HEALTHY
        alert_status = self._correlated_status(self.STATUS_HEALTHY, network_alerts)

        component_statuses = [dns_status, ntp_status, ipmi_status, alert_status]
        if any((n.get("status") == self.STATUS_CRITICAL) for n in nic_summary + bond_summary + bridge_summary):
            component_statuses.append(self.STATUS_CRITICAL)
        elif any((n.get("status") == self.STATUS_RECOMMENDED) for n in nic_summary + bond_summary + bridge_summary):
            component_statuses.append(self.STATUS_RECOMMENDED)

        if self.STATUS_CRITICAL in component_statuses:
            status = self.STATUS_CRITICAL
        elif self.STATUS_RECOMMENDED in component_statuses or len(nets) == 0:
            status = self.STATUS_RECOMMENDED
        else:
            status = self.STATUS_HEALTHY

        recs = []
        if ipmi_missing:
            recs.append(f"Configure IPMI on host(s): {', '.join(ipmi_missing[:5])}.")
        if network_alerts:
            recs.append(f"Review correlated network alert(s): {self._alert_titles(network_alerts)}.")
        if not configured_dns:
            recs.append("Configure DNS servers for the cluster.")
        if not configured_ntp:
            recs.append("Configure NTP servers for the cluster.")
        if not recs:
            recs.append("No immediate action required.")

        return {
            "status": status,
            "vlan_count": len(nets),
            "ipmi_missing_hosts": ipmi_missing,
            "cluster_vip": cluster_vip,
            "data_services_ip": data_services_ip,
            "external_subnet": external_subnet,
            "internal_subnet": internal_subnet,
            "dns_servers": configured_dns,
            "ntp_servers": configured_ntp,
            "dns_status": dns_status,
            "ntp_status": ntp_status,
            "ipmi_status": ipmi_status,
            "virtual_switch_count": len(virtual_switch_summary),
            "bridge_count": len(bridge_summary),
            "bond_count": len(bond_summary),
            "nic_count": len(nic_summary),
            "network_alert_count": len(network_alerts),
            "networks": networks[:50],
            "host_ip_summary": host_ip_summary,
            "virtual_switch_summary": virtual_switch_summary,
            "bridge_summary": bridge_summary,
            "bond_summary": bond_summary,
            "nic_summary": nic_summary,
            "network_alerts": [
                {
                    "severity": a.get("severity", "UNKNOWN"),
                    "title": a.get("title", a.get("message", "Alert")),
                    "source_host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "last_occurred": a.get("lastOccurred") or a.get("creationTime", ""),
                }
                for a in network_alerts[:10]
            ],
            "recommendations": recs,
        }

    def _analyse_storage_encryption(self, containers: list) -> dict:
        """Build a best-effort storage encryption summary from collected Prism data.

        Prism versions expose encryption details differently. The most reliable field
        currently collected by the script is per-container software encryption
        (isSoftwareEncryptionEnabled). If cluster/KMS details are present in raw data,
        this helper also attempts to identify internal vs external KMS.
        """
        visible = [
            c for c in (containers or [])
            if c.get("name") not in {
                "NutanixManagementShare", "SelfServiceContainer",
                "NutanixMetadataContainer",
            }
            and not c.get("name", "").startswith("objects")
            and not c.get("isInternal", False)
        ]

        def _truthy(v):
            if isinstance(v, bool):
                return v
            if isinstance(v, (int, float)):
                return v != 0
            if isinstance(v, str):
                return v.strip().lower() in {"true", "enabled", "enable", "on", "yes", "y"}
            return False

        # Per-container software encryption is the cleanest source in the collected data.
        encrypted_containers = []
        for c in visible:
            enabled = (
                c.get("isSoftwareEncryptionEnabled")
                if "isSoftwareEncryptionEnabled" in c else
                c.get("softwareEncryptionEnabled", c.get("encryptionEnabled", False))
            )
            if _truthy(enabled):
                encrypted_containers.append(c.get("name", "Unknown"))

        encryption_enabled = bool(encrypted_containers)
        if not visible and containers:
            # Fallback when all containers are internal/system: still check any container.
            encryption_enabled = any(_truthy(c.get("isSoftwareEncryptionEnabled", c.get("encryptionEnabled", False))) for c in containers)

        # Recursively search collected data for KMS / key-management hints.
        kms_candidates = []
        def _walk(obj, path=""):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    key_path = f"{path}.{k}" if path else str(k)
                    lk = str(k).lower()
                    if any(token in lk for token in ["kms", "keymanagement", "key_management", "keymanager", "key_manager"]):
                        kms_candidates.append((key_path, v))
                    _walk(v, key_path)
            elif isinstance(obj, list):
                for i, v in enumerate(obj[:100]):
                    _walk(v, f"{path}[{i}]")
        try:
            _walk(self.raw)
        except Exception:
            pass

        def _text(v):
            if isinstance(v, dict):
                for key in ["name", "serverName", "kmsName", "keyManagerName", "type", "provider", "vendor", "url", "ipAddress"]:
                    if v.get(key):
                        return str(v.get(key))
                return ""
            if isinstance(v, list):
                vals = [_text(x) for x in v]
                vals = [x for x in vals if x]
                return ", ".join(vals[:5])
            return str(v) if v not in (None, "") else ""

        kms_text = "; ".join([_text(v) for _, v in kms_candidates if _text(v)])
        kms_text_l = kms_text.lower()
        key_mgmt = "N/A"
        external_kms = "N/A"
        if encryption_enabled:
            if any(x in kms_text_l for x in ["external", "thales", "hytrust", "fortanix", "vormetric", "ciphertrust", "kms server"]):
                key_mgmt = "External KMS"
                external_kms = kms_text if kms_text else "Detected, name not available"
            elif any(x in kms_text_l for x in ["internal", "local", "native", "nutanix"]):
                key_mgmt = "Internal KMS"
            else:
                key_mgmt = "Not available from collected Prism Central APIs"

        container_rows = []
        for c in visible:
            enabled = c.get("isSoftwareEncryptionEnabled", c.get("softwareEncryptionEnabled", c.get("encryptionEnabled", False)))
            container_rows.append({
                "name": c.get("name", "Unknown"),
                "software_encryption": "Enabled" if _truthy(enabled) else "Disabled",
            })

        return {
            "status": "Enabled" if encryption_enabled else "Disabled",
            "encryption_type": "Software Encryption" if encryption_enabled else "N/A",
            "key_management": key_mgmt,
            "external_kms": external_kms,
            "encrypted_containers": encrypted_containers,
            "container_rows": container_rows,
        }

    def analyse_storage(self) -> dict:
        containers = self._safe_list("storage_containers")
        usage_pct  = None

        def first_number(obj: Any, names: set[str]) -> Optional[float]:
            vals = self._find_numeric_values(obj, names)
            return vals[0] if vals else None

        def has_useful_data(s: Any) -> bool:
            """Return True only if the stats dict contains real metric data, not just errors."""
            if not isinstance(s, dict):
                return False
            data = s.get("data", s)
            if isinstance(data, dict):
                return any(k for k in data if not k.startswith("_pe_stats") and not k.startswith("$"))
            return bool(data)

        stats = None
        for candidate in [self.raw.get("storage_stats"), self.raw.get("cluster_stats")]:
            if has_useful_data(candidate):
                stats = candidate
                break
        used_names = {
            "storageUsageBytes", "usedStorageCapacityBytes", "usedCapacityBytes",
            "usageBytes", "logicalUsageBytes", "physicalUsageBytes",
            "storage.usage_bytes", "storage_usage_bytes", "used_bytes",
        }
        total_names = {
            "storageCapacityBytes", "totalStorageCapacityBytes", "capacityBytes",
            "maxCapacityBytes", "storage.capacity_bytes", "storage_capacity_bytes",
            "total_bytes",
        }

        if isinstance(stats, dict):
            used = first_number(stats, used_names)
            total = first_number(stats, total_names)
            if used is not None and total:
                usage_pct = round(float(used) / float(total) * 100, 2)

        # If stats are unavailable, do not leave the section blank. We can still
        # report the pool capacity from storage container inventory.
        pool_capacity_tib = None
        if containers:
            max_cap = containers[0].get("maxCapacityBytes")
            if max_cap:
                pool_capacity_tib = round(max_cap / (1024 ** 4), 2)

        utilization_status = self.STATUS_HEALTHY
        recs   = ["No Recommendations – Storage best practices are being followed."]
        if usage_pct is not None and usage_pct > 80:
            utilization_status = self.STATUS_CRITICAL
            recs   = ["Storage utilization exceeds 80%. Immediate capacity planning required."]
        elif usage_pct is not None and usage_pct > 70:
            utilization_status = self.STATUS_RECOMMENDED
            recs   = ["Storage utilization is elevated. Review growth trends and capacity planning."]

        storage_alerts = self._alerts_matching_keywords([
            "disk", "ssd", "nvme", "stargate", "mount", "mounted",
            "unqualified", "drive", "metadata", "extent", "curator",
        ])
        status = self._correlated_status(utilization_status, storage_alerts)
        if storage_alerts:
            recs = [
                "Investigate and remediate active storage hardware or storage service alerts before treating storage health as normal.",
                f"Correlated storage alert(s): {self._alert_titles(storage_alerts)}.",
            ]
            if utilization_status != self.STATUS_HEALTHY:
                recs.append("Storage utilization also exceeds recommended thresholds. Review growth trends and capacity planning.")

        # Storage detail fields — prefer per-container stats, fall back to cluster_stats_7d
        GiB = 1024**3
        TiB = 1024**4
        PPM = 1_000_000

        def _ts_val(d: dict, field: str) -> Optional[float]:
            """Extract the most-recent value from a time-series field or scalar."""
            v = d.get(field)
            if isinstance(v, list) and v:
                item = v[0]
                val = item.get("value") if isinstance(item, dict) else item
            else:
                val = v
            try:
                return float(val) if val is not None else None
            except Exception:
                return None

        def _latest_bytes(field, divisor=1, prefer_container=False):
            # Try per-container stats first when requested
            if prefer_container:
                for cs in self.raw.get("container_stats", []):
                    d = cs.get("stats", {}).get("data", cs.get("stats", {}))
                    if isinstance(d, list) and d: d = d[0]
                    if isinstance(d, dict):
                        val = _ts_val(d, field)
                        if val is not None:
                            return round(val / divisor, 2)
            # Fall back to cluster_stats_7d
            src7 = self.raw.get("cluster_stats_7d") or self.raw.get("cluster_stats")
            if not isinstance(src7, dict): return None
            d = src7.get("data", src7)
            if isinstance(d, list) and d: d = d[0]
            if not isinstance(d, dict): return None
            val = _ts_val(d, field)
            return round(val / divisor, 2) if val is not None else None

        # Per-container data reduction ratio from container stats
        def _container_dr_ratio() -> Optional[float]:
            for cs in self.raw.get("container_stats", []):
                d = cs.get("stats", {}).get("data", cs.get("stats", {}))
                if isinstance(d, list) and d: d = d[0]
                if not isinstance(d, dict): continue
                for field in ["dataReductionRatio", "data_reduction_ratio"]:
                    val = _ts_val(d, field)
                    if val is not None:
                        return round(val / PPM, 2) if val > 100 else round(val, 2)
            # Fallback: logical / physical
            lu = _latest_bytes("logicalStorageUsageBytes", prefer_container=True)
            pu = _latest_bytes("storageUsageBytes", prefer_container=True)
            try:
                return round(lu / pu, 2) if lu and pu and pu > 0 else None
            except Exception:
                return None

        storage_detail = {
            "free_capacity_tib":    _latest_bytes("freePhysicalStorageBytes", TiB, True),
            "used_physical_gib":    _latest_bytes("storageUsageBytes",         GiB, True),
            "snapshot_gib":         _latest_bytes("snapshotCapacityBytes",     GiB, True),
            "max_capacity_tib":     _latest_bytes("storageCapacityBytes",      TiB),
            "logical_usage_gib":    _latest_bytes("logicalStorageUsageBytes",  GiB, True),
            "savings_gib":          _latest_bytes("overallSavingsBytes",       GiB, True),
            "savings_ratio":        _latest_bytes("overallSavingsRatio",       PPM),
            "recycle_bin_gib":      _latest_bytes("recycleBinUsageBytes",      GiB),
            "data_reduction_ratio": _container_dr_ratio(),
        }

        # 7-day storage usage % trend — derived from storageUsageBytes / storageCapacityBytes
        storage_history = []
        stats_src = self.raw.get("cluster_stats_7d") or self.raw.get("cluster_stats")
        if isinstance(stats_src, dict):
            d = stats_src.get("data", stats_src)
            if isinstance(d, list) and d: d = d[0]
            if isinstance(d, dict):
                used_ts  = d.get("storageUsageBytes", [])
                cap_ts   = d.get("storageCapacityBytes", [])
                cap_by_ts = {pt["timestamp"]: pt.get("value", 0)
                             for pt in cap_ts if isinstance(pt, dict) and "timestamp" in pt}
                default_cap = cap_ts[0].get("value", 1) if cap_ts and isinstance(cap_ts[0], dict) else 1
                for pt in (used_ts if isinstance(used_ts, list) else []):
                    if not isinstance(pt, dict): continue
                    try:
                        cap = cap_by_ts.get(pt["timestamp"], default_cap) or default_cap
                        pct = round(float(pt["value"]) / float(cap) * 100, 2)
                        storage_history.append({"timestamp": pt["timestamp"], "value": pct})
                    except Exception:
                        continue
        storage_history = self._downsample_time_series(storage_history, max_points=120)

        visible_containers = [
            c for c in containers
            if c.get("name") not in {
                "NutanixManagementShare", "SelfServiceContainer",
                "NutanixMetadataContainer",
            }
            and not c.get("name", "").startswith("objects")
            and not c.get("isInternal", False)
        ]

        def _fmt_capacity_bytes(value) -> str:
            """Return a readable capacity string for container configuration values."""
            try:
                if value is None or value == "":
                    return "N/A"
                n = float(value)
                if n <= 0:
                    return "None"
                if n >= TiB:
                    return f"{round(n / TiB, 2):g} TiB"
                if n >= GiB:
                    return f"{round(n / GiB, 2):g} GiB"
                return f"{round(n / (1024 ** 2), 2):g} MiB"
            except Exception:
                return "N/A"

        def _compression_delay_text(container: dict) -> str:
            enabled = container.get("isCompressionEnabled", container.get("compressionEnabled", False))
            if not enabled:
                return "N/A"
            delay = container.get("compressionDelaySecs")
            try:
                seconds = int(delay or 0)
            except Exception:
                return "N/A"
            if seconds <= 0:
                return "Immediate"
            if seconds % 3600 == 0:
                hours = seconds // 3600
                return f"{hours} hour" + ("s" if hours != 1 else "")
            if seconds % 60 == 0:
                minutes = seconds // 60
                return f"{minutes} minutes"
            return f"{seconds} seconds"

        encryption_summary = self._analyse_storage_encryption(containers)

        return {
            "status":               status,
            "disk_utilization_pct": usage_pct if usage_pct is not None else "N/A",
            "storage_utilization_status": utilization_status,
            "storage_alert_count": len(storage_alerts),
            "storage_alerts": [
                {
                    "severity": a.get("severity", "UNKNOWN"),
                    "title": a.get("title", a.get("message", "Alert")),
                    "source": a.get("sourceType") or a.get("classification") or "N/A",
                    "host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "source_host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "last_occurred": a.get("lastOccurred") or a.get("creationTime", ""),
                }
                for a in storage_alerts[:10]
            ],
            "pool_capacity_tib":    pool_capacity_tib,
            "storage_history":      storage_history,
            "storage_detail":       storage_detail,
            "encryption":           encryption_summary,
            "container_count":      len(visible_containers),
            "containers": [
                {
                    "name":         c.get("name", "Unknown"),
                    "compression":  c.get("isCompressionEnabled", c.get("compressionEnabled", False)),
                    "dedup":        c.get("onDiskDedup", c.get("fingerPrintOnWrite", "OFF")),
                    "erasure_code": c.get("erasureCode", "OFF"),
                    "rf":           c.get("replicationFactor", "N/A"),
                    "max_capacity_tib": round(c.get("maxCapacityBytes", 0) / 1024**4, 2)
                                        if c.get("maxCapacityBytes") else "N/A",
                    "reserved_capacity_logical": _fmt_capacity_bytes(c.get("logicalExplicitReservedCapacityBytes")),
                    "advertised_capacity_logical": _fmt_capacity_bytes(c.get("logicalAdvertisedCapacityBytes")) if "logicalAdvertisedCapacityBytes" in c else "N/A",
                    "compression_delay": _compression_delay_text(c),
                    "software_encryption": ("Enabled" if str(c.get("isSoftwareEncryptionEnabled", c.get("softwareEncryptionEnabled", c.get("encryptionEnabled", False)))).lower() in {"true", "enabled", "enable", "on", "yes", "1"} else "Disabled"),
                }
                for c in visible_containers
            ],
            "recommendations": recs,
        }

    def analyse_licensing(self) -> dict:
        lic        = self.raw.get("licensing")
        expiry     = "N/A"
        violations = []
        if isinstance(lic, dict):
            data = lic.get("data", [])
            if isinstance(data, list) and data:
                first = data[0]
            elif isinstance(data, dict):
                first = data
            else:
                first = {}
            # v4 API expiry field names vary by version
            expiry = (first.get("expiryDate") or
                      first.get("expirationDate") or
                      first.get("supportExpirationDate") or
                      first.get("clusterExpiryDate") or "N/A")
            violations  = (first.get("violations") or
                           first.get("licenseViolations") or [])
            lic_name    = first.get("name", "N/A")
            lic_type    = first.get("type", first.get("category", "N/A"))
        status = self.STATUS_CRITICAL if violations else self.STATUS_HEALTHY
        recs   = (["License violations detected – contact Nutanix Support immediately."]
                  if violations else
                  ["No license violations detected.",
                   "Plan renewal before expiry date to avoid service disruption."])
        return {
            "status":          status,
            "expiry_date":     expiry,
            "license_name":    lic_name if lic else "N/A",
            "license_type":    lic_type if lic else "N/A",
            "violations":      violations,
            "recommendations": recs,
        }

    def analyse_security(self) -> dict:
        """Assess security configuration separately from software lifecycle."""
        c = self._cluster_first()
        config = c.get("config", {}) if isinstance(c, dict) else {}

        hardening_raw = self.raw.get("security_hardening") or {}
        if not isinstance(hardening_raw, dict):
            hardening_raw = {}
        pc_security = hardening_raw.get("pc_security_summary") or {}
        pc_config_summary = pc_security.get("securityConfigSummary") or {}
        pe_cluster = hardening_raw.get("pe_cluster") or {}
        if isinstance(pe_cluster, dict) and isinstance(pe_cluster.get("data"), dict):
            pe_cluster = pe_cluster["data"]
        if not isinstance(pe_cluster, dict):
            pe_cluster = {}
        cvm_security = pe_cluster.get("security_compliance_config") or {}
        ahv_security = pe_cluster.get("hypervisor_security_compliance_config") or {}

        def _first_known(*values):
            return next((value for value in values if value is not None), None)

        def _combined_bool(*values):
            known = [value for value in values if isinstance(value, bool)]
            return (all(known) if known else None)

        def _bool_text(value):
            return "Enabled" if value is True else ("Disabled" if value is False else "N/A")

        def _hardening_status(value):
            if value is None:
                return "N/A"
            return self.STATUS_HEALTHY if value is True else self.STATUS_RECOMMENDED

        def _hardware_generation(node):
            """Return the Nutanix platform generation encoded in the model."""
            if not isinstance(node, dict):
                return None
            model_values = [
                node.get("blockModel"),
                node.get("block_model"),
                node.get("model"),
                node.get("hardwareModel"),
                node.get("hardware_model"),
            ]
            for model_value in model_values:
                if not model_value:
                    continue
                match = re.search(r"(?:^|[^A-Z0-9])G(?:EN)?[-_ ]?(\d+)(?:$|[^0-9])", str(model_value), re.IGNORECASE)
                if match:
                    return int(match.group(1))
            return None

        def _node_model(node):
            if not isinstance(node, dict):
                return "N/A"
            return str(
                node.get("blockModel")
                or node.get("block_model")
                or node.get("model")
                or node.get("hardwareModel")
                or node.get("hardware_model")
                or "N/A"
            )

        password_remote_login = config.get("isPasswordRemoteLoginEnabled")
        remote_support = config.get("isRemoteSupportEnabled")
        pulse_enabled = (config.get("pulseStatus") or {}).get("isEnabled")
        nodes = self._safe_list("nodes")
        secure_boot_values = [n.get("isSecureBooted") for n in nodes if isinstance(n, dict) and n.get("isSecureBooted") is not None]
        secure_boot_eligible_nodes = [
            n for n in nodes
            if isinstance(n, dict) and (_hardware_generation(n) is None or _hardware_generation(n) >= 8)
        ]
        secure_boot_eligible_values = [
            n.get("isSecureBooted") for n in secure_boot_eligible_nodes
            if n.get("isSecureBooted") is not None
        ]
        pre_g8_nodes = [
            n for n in nodes
            if isinstance(n, dict) and _hardware_generation(n) is not None and _hardware_generation(n) < 8
        ]
        secure_boot_enabled = bool(secure_boot_eligible_values) and all(bool(v) for v in secure_boot_eligible_values)
        if secure_boot_eligible_values:
            secure_boot_text = "Enabled on all supported hosts" if secure_boot_enabled else "Disabled on one or more supported hosts"
            secure_boot_status = self.STATUS_HEALTHY if secure_boot_enabled else self.STATUS_RECOMMENDED
        elif pre_g8_nodes and len(pre_g8_nodes) == len(nodes):
            secure_boot_text = "Not supported on pre-G8 platforms"
            secure_boot_status = self.STATUS_HEALTHY
        else:
            secure_boot_text = "N/A"
            secure_boot_status = "N/A"
        secure_boot_enabled_count = sum(1 for v in secure_boot_eligible_values if bool(v))

        pc_full_config = pc_security.get("securityConfig") or pc_security.get("securityConfiguration") or {}
        if not isinstance(pc_full_config, dict):
            pc_full_config = {}
        cluster_lockdown = _first_known(
            pc_config_summary.get("isClusterLockdownEnabled"),
            pe_cluster.get("enable_lock_down"),
            pe_cluster.get("enableLockDown"),
        )
        log_forwarding = pc_config_summary.get("isLogForwardingEnabled")
        defense_banner = _first_known(
            pc_config_summary.get("isConsentBannerEnabled"),
            pc_full_config.get("isClusterDefenseConsentBannerEnabled"),
            pc_full_config.get("isAhvDefenseConsentBannerEnabled"),
            _combined_bool(
                cvm_security.get("enable_banner"),
                ahv_security.get("enable_banner"),
            ),
        )
        network_segmentation = (
            ((c.get("network") or {}).get("backplane") or {}).get("isSegmentationEnabled")
            if isinstance(c, dict) else None
        )
        host_secure_boot = []
        for node in sorted(nodes, key=lambda n: str(n.get("hostName") or n.get("name") or n.get("extId") or "").lower()):
            value = node.get("isSecureBooted") if isinstance(node, dict) else None
            generation = _hardware_generation(node)
            if generation is not None and generation < 8:
                boot_value = "Not supported (pre-G8)"
                boot_status = self.STATUS_HEALTHY
            else:
                boot_value = "Enabled" if value is True else ("Disabled" if value is False else "N/A")
                boot_status = self.STATUS_HEALTHY if value is True else (self.STATUS_RECOMMENDED if value is False else "N/A")
            host_secure_boot.append({
                "host": node.get("hostName") or node.get("name") or node.get("extId") or "Unknown",
                "model": _node_model(node),
                "secure_boot": boot_value,
                "status": boot_status,
            })

        containers = self._safe_list("storage_containers")
        encryption = self._analyse_storage_encryption(containers)

        security_alerts = self._alerts_matching_keywords([
            "default password",
            "password based ssh",
            "password-based ssh",
            "reset current passwords",
            "password expir",
            "encryption key",
            "key backup",
            "certificate",
            "authentication",
            "unauthorized",
            "security",
            "tls",
            "ssl",
        ])
        critical_security_alerts = [
            a for a in security_alerts
            if str(a.get("severity", "")).upper() in {"CRITICAL", "ERROR", "FATAL"}
        ]

        findings = []
        recs = []
        status = self.STATUS_HEALTHY

        if remote_support is True:
            status = self.STATUS_CRITICAL
            findings.append("Remote Support is enabled")
            recs.append("Disable Remote Support when it is not actively required by Nutanix Support.")

        if cluster_lockdown is False:
            if status == self.STATUS_HEALTHY:
                status = self.STATUS_RECOMMENDED
            findings.append("Cluster Lockdown is disabled")
            recs.append("Enable Cluster Lockdown after validating key-based SSH access and operational access requirements.")

        if password_remote_login is True:
            if status == self.STATUS_HEALTHY:
                status = self.STATUS_RECOMMENDED
            findings.append("Password-based remote login is enabled")
            recs.append("Disable password-based remote login where operationally appropriate and use key-based SSH access.")

        if secure_boot_eligible_values and not secure_boot_enabled:
            if status == self.STATUS_HEALTHY:
                status = self.STATUS_RECOMMENDED
            findings.append("Secure Boot is disabled on one or more G8-or-newer/eligible hosts")
            recs.append("Enable Host Secure Boot on supported G8-or-newer platforms during an approved maintenance procedure.")

        if encryption.get("status") == "Disabled":
            if status == self.STATUS_HEALTHY:
                status = self.STATUS_RECOMMENDED
            findings.append("Storage software encryption is disabled")
            recs.append("Review data-at-rest encryption requirements and enable Nutanix software encryption where required by policy.")

        if pulse_enabled is False:
            if status == self.STATUS_HEALTHY:
                status = self.STATUS_RECOMMENDED
            findings.append("Pulse is disabled")
            recs.append("Enable Pulse to provide proactive support telemetry, subject to the organization's data-sharing policy.")

        alert_text = " ".join(self._alert_text(a) for a in security_alerts)
        if "default password" in alert_text:
            recs.insert(0, "Change default host and CVM credentials immediately, then verify the related alerts clear.")
        if "reset current passwords" in alert_text or "password expir" in alert_text:
            recs.append("Reset expired or flagged administrative passwords in accordance with the organization's password policy.")
        if "encryption key" in alert_text or "key backup" in alert_text:
            recs.append("Review the data-at-rest encryption key warning and confirm a current key backup is stored securely.")
        if security_alerts and not any(x in alert_text for x in ("default password", "reset current passwords", "password expir", "encryption key", "key backup")):
            recs.append("Review and remediate the active security alerts listed in this section.")

        status = self._correlated_status(status, security_alerts)

        if not recs:
            recs.append("No immediate security configuration changes are recommended based on the collected Prism Central data.")

        # Preserve recommendation order while removing duplicate guidance.
        recs = list(dict.fromkeys(recs))

        def _setting_status(value, healthy_when):
            if value is None:
                return "N/A"
            return self.STATUS_HEALTHY if value is healthy_when else self.STATUS_RECOMMENDED

        configuration_items = [
            {
                "item": "Password-Based Remote Login",
                "value": "Enabled" if password_remote_login is True else ("Disabled" if password_remote_login is False else "N/A"),
                "status": _setting_status(password_remote_login, False),
            },
            {
                "item": "Remote Support",
                "value": "Enabled" if remote_support is True else ("Disabled" if remote_support is False else "N/A"),
                "status": (self.STATUS_CRITICAL if remote_support is True else (self.STATUS_HEALTHY if remote_support is False else "N/A")),
            },
            {
                "item": "Pulse",
                "value": "Enabled" if pulse_enabled is True else ("Disabled" if pulse_enabled is False else "N/A"),
                "status": _setting_status(pulse_enabled, True),
            },
            {
                "item": "Data-at-Rest Encryption",
                "value": encryption.get("status", "N/A"),
                "status": (self.STATUS_HEALTHY if encryption.get("status") == "Enabled" else self.STATUS_RECOMMENDED),
            },
        ]

        security_hardening_items = [
            {
                "item": "Host Secure Boot",
                "value": secure_boot_text,
                "status": secure_boot_status,
            },
            {
                "item": "AOS Network Segmentation",
                "value": _bool_text(network_segmentation),
                "status": _hardening_status(network_segmentation),
            },
            {
                "item": "Cluster Lockdown",
                "value": _bool_text(cluster_lockdown),
                "status": _hardening_status(cluster_lockdown),
            },
            {
                "item": "Log Forwarding",
                "value": _bool_text(log_forwarding),
                "status": _hardening_status(log_forwarding),
            },
            {
                "item": "Defense Consent Banner",
                "value": _bool_text(defense_banner),
                "status": _hardening_status(defense_banner),
            },
        ]

        return {
            "status": status,
            "password_remote_login": "Enabled" if password_remote_login is True else ("Disabled" if password_remote_login is False else "N/A"),
            "remote_support": "Enabled" if remote_support is True else ("Disabled" if remote_support is False else "N/A"),
            "pulse": "Enabled" if pulse_enabled is True else ("Disabled" if pulse_enabled is False else "N/A"),
            "secure_boot": secure_boot_text,
            "storage_encryption": encryption.get("status", "N/A"),
            "key_management": encryption.get("key_management", "N/A"),
            "external_kms": encryption.get("external_kms", "N/A"),
            "encryption_type": encryption.get("encryption_type", "N/A"),
            "encrypted_containers": len(encryption.get("encrypted_containers", [])),
            "configuration_items": configuration_items,
            "security_hardening_items": security_hardening_items,
            "cluster_lockdown": _bool_text(cluster_lockdown),
            "log_forwarding": _bool_text(log_forwarding),
            "network_segmentation": _bool_text(network_segmentation),
            "host_secure_boot": host_secure_boot,
            "secure_boot_enabled_count": secure_boot_enabled_count,
            "secure_boot_host_count": len(secure_boot_eligible_values),
            "secure_boot_note": "Secure Boot support begins with Nutanix G8 platforms.",
            "security_alert_count": len(security_alerts),
            "critical_security_alert_count": len(critical_security_alerts),
            "security_alerts": [
                {
                    "severity": a.get("severity", "UNKNOWN"),
                    "title": a.get("title", a.get("message", "Alert")),
                    "source_host": a.get("sourceHost") or a.get("sourceEntity") or "N/A",
                    "last_occurred": a.get("lastOccurred") or a.get("creationTime", ""),
                }
                for a in security_alerts[:15]
            ],
            "findings": findings,
            "recommendations": recs,
        }

    def analyse_software_lifecycle(self) -> dict:
        """Assess AOS lifecycle and software/firmware upgrade planning."""
        c = self._cluster_first()
        bi = c.get("config", {}).get("buildInfo", {})
        full_ver = bi.get("fullVersion", "")
        m = re.search(r"(\d+\.\d+\.\d+\.\d+)", full_ver)
        aos = m.group(1) if m else bi.get("version", "N/A")
        aos_lifecycle = _load_aos_eol_info(self.aos_eol_csv, aos)
        status = aos_lifecycle.get("report_status", self.STATUS_RECOMMENDED)
        recs = []
        latest = aos_lifecycle.get("latest_version", "N/A")
        if status == self.STATUS_CRITICAL:
            recs.append("Upgrade AOS to a supported release as soon as possible.")
        elif latest not in ("", "N/A") and _version_tuple(aos) < _version_tuple(latest):
            recs.append(f"Upgrade AOS from {aos} to {latest} via LCM.")
        elif status == self.STATUS_HEALTHY:
            recs.append("Current AOS release is within the supported lifecycle window.")
        else:
            recs.append("Review AOS lifecycle status and plan an upgrade during the next maintenance window.")
        recs += [
            "Upgrade AHV to the version recommended for the selected AOS release.",
            "Upgrade NCC to the latest supported version.",
            "Run an LCM inventory and apply supported software and firmware updates.",
        ]
        return {
            "status": status,
            "aos_version": aos,
            "aos_lifecycle": aos_lifecycle,
            "recommendations": recs,
        }

    def analyse_ncc(self) -> dict:
        checks = self._safe_list("ncc_checks")
        recs   = []
        if checks:
            recs.append(f"Review {len(checks)} NCC health check result(s).")
        recs += [
            "Clean up any orphan VM snapshots (Nutanix KB-375).",
            "Schedule regular NCC runs and export results for review.",
        ]
        return {
            "status":      self.STATUS_RECOMMENDED if checks else self.STATUS_HEALTHY,
            "check_count": len(checks),
            "checks": [
                {"title": ch.get("title", ch.get("message", "NCC Check")), "severity": ch.get("severity", "INFO")}
                for ch in checks[:15]
            ],
            "recommendations": recs,
        }

    def analyse_all(self) -> dict:
        return {
            "customer":   self.customer,
            "date":       datetime.now().strftime("%B %d, %Y"),
            "cluster":    self.analyse_cluster_info(),
            "vms":        self.analyse_virtual_machines(),
            "cvms":       self.analyse_cvms(),
            "vm_counts":  self._vm_summary_counts(),
            "health":     self.analyse_health(),
            "protection": self.analyse_protection(),
            "cpu":        self.analyse_cpu(),
            "memory":     self.analyse_memory(),
            "network":    self.analyse_network(),
            "storage":    self.analyse_storage(),
            "licensing":  self.analyse_licensing(),
            "security":   self.analyse_security(),
            "software_lifecycle": self.analyse_software_lifecycle(),
            "ncc":        self.analyse_ncc(),
        }


# ---------------------------------------------------------------------------
# Report Generator  (docx via Node.js)
# ---------------------------------------------------------------------------

REPORT_JS = r"""
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, VerticalAlign, PageNumber, TableLayoutType,
  Bookmark, InternalHyperlink
} = require('docx');
const fs = require('fs');

const DATA = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

// Standard report layout settings
const REPORT_FONT = "Calibri";       // Change to "Calibri" if broader compatibility is needed
const REPORT_FONT_SIZE = 20;       // 10 pt in docx half-points
const TABLE_FONT_SIZE = 18;        // 9 pt for tables, consistent across dense report tables
const TABLE_HEADER_SIZE = 18;      // 9 pt table headers
const CONTENT_WIDTH = 10800;       // Letter portrait content width: 8.5in - 0.5in margins left/right
const CHART_WIDTH = 720;          // 7.5in content width at 96px/in
const CHART_HEIGHT = 340;         // Consistent full-width chart height
const H1_SIZE = 32;                // 16 pt
const H2_SIZE = 26;                // 13 pt
const PARA_AFTER = 120;            // 6 pt in twips
const COMPACT_AFTER = 40;           // 2 pt in twips
const TIGHT_AFTER = 0;              // no extra spacing
const PAGE_MARGIN = 720;           // 0.5 inch in twips
const HEADER_FOOTER_DISTANCE = 432; // 0.3 inch in twips
// Table rows use cantSplit to prevent row content from being divided across pages.
// Header rows use tableHeader so Word repeats column headers on continuation pages.

const NUTANIX_BLUE   = "005F9E";
const NUTANIX_LIGHT  = "D6E8F7";
const HEADER_GREY    = "404040";
const ROW_ALT        = "F2F7FB";
const STATUS_COLORS  = { "Healthy": "00843D", "Recommended": "E5A000", "Critical": "CC0000" };
const LINK_BLUE = "0563C1";
const SECTION_BOOKMARKS = {
  "Virtual Machines Summary": "sec_virtual_machines",
  "Alerts Summary": "sec_alerts",
  "Data Protection Summary": "sec_data_protection",
  "Cluster CPU Summary": "sec_cluster_cpu_summary",
  "Cluster Memory Summary": "sec_cluster_memory_summary",
  "Cluster Storage Summary": "sec_cluster_storage_summary",
  "Network Summary": "sec_network",
  "Licensing Summary": "sec_licensing",
  "Security Summary": "sec_security",
  "Software Lifecycle Summary": "sec_software_lifecycle",
  "NCC Health Checks Summary": "sec_ncc_health_checks",
};

function statusRun(status) {
  return new TextRun({ text: `● ${status}`, bold: true, color: STATUS_COLORS[status] || "000000", size: REPORT_FONT_SIZE, font: REPORT_FONT });
}

function heading1(text) {
  const titleRun = new TextRun({ text, font: REPORT_FONT, bold: true, size: H1_SIZE, color: NUTANIX_BLUE });
  const children = SECTION_BOOKMARKS[text]
    ? [new Bookmark({ id: SECTION_BOOKMARKS[text], children: [titleRun] })]
    : [titleRun];
  return new Paragraph({
    keepNext: true,
    heading: HeadingLevel.HEADING_1,
    children,
    spacing: { before: 360, after: PARA_AFTER },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NUTANIX_BLUE, space: 4 } },
  });
}

function heading2(text) {
  return new Paragraph({
    keepNext: true,
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: REPORT_FONT, bold: true, size: H2_SIZE, color: HEADER_GREY })],
    spacing: { before: 160, after: 60 },
  });
}

function compactHeading2(text) {
  // Use the same formatting as standard Heading 2 so executive-summary
  // headings behave consistently in Word.
  return heading2(text);
}

function body(text, opts = {}) {
  return new Paragraph({ children: [new TextRun({ text, font: REPORT_FONT, size: REPORT_FONT_SIZE, ...opts })], spacing: { before: 0, after: opts.after === 0 ? 0 : (opts.after || PARA_AFTER) } });
}

function compactBody(text, opts = {}) {
  return new Paragraph({ children: [new TextRun({ text, font: REPORT_FONT, size: REPORT_FONT_SIZE, ...opts })], spacing: { before: 0, after: COMPACT_AFTER } });
}

function bulletItem(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text, font: REPORT_FONT, size: REPORT_FONT_SIZE })], spacing: { before: 0, after: TIGHT_AFTER } });
}

function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

const B = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const BORDERS = { top: B, bottom: B, left: B, right: B };

const cell = (text, opts = {}) => new TableCell({
  borders: BORDERS,
  shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
  margins: { top: 60, bottom: 60, left: 120, right: 120 },
  width: { size: opts.width || 4680, type: WidthType.DXA },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({ children: [new TextRun({ text: String(text), font: REPORT_FONT, size: opts.size || TABLE_FONT_SIZE, bold: !!opts.bold, color: opts.color || "000000" })], spacing: { before: 0, after: 0 } })],
});

const hdrCell = (text, width = 4680) => cell(text, { shading: NUTANIX_BLUE, color: "FFFFFF", bold: true, width });

function sectionLinkCell(text, width = 3500, shading = "FFFFFF") {
  const bookmark = SECTION_BOOKMARKS[text];
  const run = new TextRun({
    text: String(text || ""),
    font: REPORT_FONT,
    size: TABLE_FONT_SIZE,
    color: bookmark ? LINK_BLUE : "000000",
    underline: bookmark ? {} : undefined,
  });
  return new TableCell({
    borders: BORDERS,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: bookmark ? [new InternalHyperlink({ anchor: bookmark, children: [run] })] : [run],
      spacing: { before: 0, after: 0 },
    })],
  });
}

function twoColTable(rows, widths = [3600, 7200]) {
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((row, i) => new TableRow({ cantSplit: true, children: row.map((val, j) => cell(val, { width: widths[j], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT })) })),
  });
}

function sectionTable(label, status, observation, recommendations) {
  const obsLines = String(observation || "").split("\n").filter(l => l.trim());
  const recArray = Array.isArray(recommendations) ? recommendations : [recommendations];

  const mkPara = (text, bullet = false, bold = false) => {
    const cleanText = String(text || "").replace(/^[•\-]\s*/, "");
    let runs = [new TextRun({ text: cleanText, font: REPORT_FONT, size: REPORT_FONT_SIZE, bold })];
    if (bullet && cleanText.includes(":")) {
      // Bold only the label portion of bullets. Use the first colon followed
      // by a space so labels containing internal colons, such as
      // "Cluster vCPU:pCore Ratio:", are handled correctly while values like
      // "1.12 : 1" remain unbolded.
      let idx = cleanText.indexOf(": ");
      if (idx < 0) idx = cleanText.lastIndexOf(":");
      runs = [
        new TextRun({ text: cleanText.slice(0, idx + 1), font: REPORT_FONT, size: REPORT_FONT_SIZE, bold: true }),
        new TextRun({ text: cleanText.slice(idx + 1), font: REPORT_FONT, size: REPORT_FONT_SIZE }),
      ];
    }
    return new Paragraph({
      numbering: bullet ? { reference: "bullets", level: 0 } : undefined,
      children: runs,
      spacing: { before: 0, after: bullet ? TIGHT_AFTER : COMPACT_AFTER },
    });
  };

  const obsParas = obsLines.length
    ? obsLines.map(l => mkPara(l.replace(/^•\s*/, ""), l.trim().startsWith("•")))
    : [mkPara("No observations recorded.")];

  const recParas = recArray.filter(r => String(r || "").trim()).length
    ? recArray.filter(r => String(r || "").trim()).map(r => mkPara(r, true))
    : [mkPara("No immediate action required.", true)];

  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [2300, 8500],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", 2300), hdrCell("Comments", 8500)] }),
      new TableRow({ cantSplit: true, children: [
        cell("Status", { width: 2300, shading: NUTANIX_LIGHT, bold: true }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 8500, type: WidthType.DXA },
          children: [new Paragraph({ children: [statusRun(status)], spacing: { before: 0, after: 0 } })],
        }),
      ]}),
      new TableRow({ cantSplit: true, children: [
        cell("Observation", { width: 2300, shading: NUTANIX_LIGHT, bold: true }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 8500, type: WidthType.DXA },
          children: obsParas,
        }),
      ]}),
      new TableRow({ cantSplit: true, children: [
        cell("Recommendations", { width: 2300, shading: NUTANIX_LIGHT, bold: true }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 8500, type: WidthType.DXA },
          children: recParas,
        }),
      ]}),
    ],
  });
}

const D = DATA;
const C = D.cluster;

const summaryStatuses = [
  ["Alerts Summary",               D.health.status],
  ["Virtual Machines Summary",     D.vms.status],
  ["Data Protection Summary",      D.protection.status],
  ["Cluster CPU Summary",  D.cpu.status],
  ["Cluster Memory Summary", D.memory.status],
  ["Network Summary",              D.network.status],
  ["Cluster Storage Summary", D.storage.status],
  ["Licensing Summary",            D.licensing.status],
  ["Security Summary", D.security.status],
  ["Software Lifecycle Summary", D.software_lifecycle.status],
  ["NCC Health Checks Summary",    D.ncc.status],
];


function makeChart(chartPath, history, heading, noDataMsg) {
  if (chartPath && fs.existsSync(chartPath)) {
    return [
      heading2(heading),
      new Paragraph({
        children: [new ImageRun({ data: fs.readFileSync(chartPath), transformation: { width: CHART_WIDTH, height: CHART_HEIGHT }, type: "png" })],
        spacing: { before: 120, after: 120 },
      }),
    ];
  }
  if (Array.isArray(history) && history.length >= 2) {
    return [body(`${heading}: graph could not be rendered — install matplotlib on the host running the script.`)];
  }
  return [body(noDataMsg)];
}

function cpuChart(history) {
  return makeChart(
    D.cpu.cpu_chart_path,
    history,
    "CPU Usage Trend (7 Days)",
    "CPU usage graph: not enough historical data returned from Prism Central."
  );
}

function memChart(history) {
  return makeChart(
    D.memory.mem_chart_path,
    history,
    "Memory Usage Trend (7 Days)",
    "Memory usage graph: not enough historical data returned from Prism Central."
  );
}

function storageChart(history) {
  return makeChart(
    D.storage.storage_chart_path,
    history,
    "Storage Usage Trend (7 Days)",
    "Storage usage graph: not enough historical data returned from Prism Central."
  );
}

function vmDetailTable() {
  const vms = [...(D.vms.vm_list || [])].sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), undefined, { numeric: true, sensitivity: "base" }));

  const CW  = [1700, 550, 1350, 1500, 1150, 480, 600, 600, 820, 820, 1230];
  const hdrs = ["VM Name","Status","IP Address","Operating System","OS Support","vCPU","Mem\n(GiB)","Disk\n(GiB)","CD-ROM","NGT Status","NGT Ver"];

  const ngtColor = (s) => s === "Enabled" ? "00843D" : s === "Not Installed" ? "CC0000" : "E5A000";
  const osSupportColor = (s) => s === "Supported" ? "00843D" : s === "Legacy Support" ? "E5A000" : s === "Ignored" ? "666666" : "CC0000";
  const osSupportFill = (s, bg) => s === "Unsupported" ? "F4CCCC" : s === "Legacy Support" ? "FFF2CC" : bg;

  const mkCell = (text, width, opts = {}) => new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    width: { size: width, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), font: REPORT_FONT, size: opts.size || TABLE_FONT_SIZE, bold: !!opts.bold, color: opts.color || "000000" })],
      spacing: { before: 0, after: 0 },
    })],
  });

  const hdrRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: hdrs.map((h, i) => new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
      shading: { fill: NUTANIX_BLUE, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      width: { size: CW[i], type: WidthType.DXA },
      children: [new Paragraph({ alignment: i > 3 ? AlignmentType.CENTER : AlignmentType.LEFT, children: [new TextRun({ text: h, font: REPORT_FONT, bold: true, size: TABLE_HEADER_SIZE, color: "FFFFFF" })] })],
    })),
  });

  const dataRows = (!vms.length) ? [
    new TableRow({ cantSplit: true, children: [
      mkCell("N/A", CW[0], { fill: "FFFFFF", bold: true }),
      mkCell("N/A", CW[1], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[2], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[3], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[4], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[5], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[6], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[7], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[8], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[9], { fill: "FFFFFF", center: true }),
      mkCell("N/A", CW[10], { fill: "FFFFFF", center: true }),
    ]})
  ] : vms.map((vm, i) => {
    const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    return new TableRow({ cantSplit: true, children: [
      mkCell(vm.name,        CW[0], { fill: bg, bold: true }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, width: { size: CW[1], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: vm.power === "ON" ? "On" : "Off", font: REPORT_FONT, size: TABLE_FONT_SIZE, bold: true, color: vm.power === "ON" ? "00843D" : "CC0000" })] })],
      }),
      mkCell(vm.ip,          CW[2], { fill: bg, size: 18 }),
      mkCell(vm.os,          CW[3], { fill: bg, size: 18 }),
      mkCell(vm.os_support || "Unsupported", CW[4], { fill: osSupportFill(vm.os_support, bg), size: 18, bold: true, color: osSupportColor(vm.os_support) }),
      mkCell(vm.vcpus,       CW[5], { fill: bg, center: true }),
      mkCell(vm.mem_gib,     CW[6], { fill: bg, center: true }),
      mkCell(vm.disk_gib,    CW[7], { fill: bg, center: true }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
        shading: { fill: (vm.cdrom && vm.cdrom !== "—") ? "FFF3CD" : bg, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        width: { size: CW[9], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({
          text: vm.cdrom || "—",
          font: REPORT_FONT, size: TABLE_FONT_SIZE,
          color: (vm.cdrom && vm.cdrom !== "—") ? "7A5200" : "000000",
          bold: !!(vm.cdrom && vm.cdrom !== "—"),
        })] })],
      }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, width: { size: CW[9], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: vm.ngt_status, font: REPORT_FONT, size: TABLE_FONT_SIZE, bold: true, color: ngtColor(vm.ngt_status) })] })],
      }),
      mkCell(vm.ngt_version, CW[10], { fill: bg, center: true }),
    ]});
  });

  return [
    heading2("Virtual Machine Inventory"),
    new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [hdrRow, ...dataRows] }),
  ];
}

function cvmTable() {
  const cvms = [...((D.cvms || {}).cvms || [])].sort((a, b) => String(a.host_name || "").localeCompare(String(b.host_name || ""), undefined, { numeric: true, sensitivity: "base" }));
  if (!cvms.length) return [body("No CVM data available — host-nodes API may require additional permissions.")];

  const CW  = [2600, 2300, 1700, 1500, 1500, 1200];
  const hdrs = ["CVM Name", "Host", "CVM IP Address", "Memory (GiB)", "vCPUs", "Status"];

  const mkHdr = (text, width) => new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
    shading: { fill: NUTANIX_BLUE, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    width: { size: width, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text, font: REPORT_FONT, bold: true, size: TABLE_HEADER_SIZE, color: "FFFFFF" })] })],
  });

  const mkCell = (text, width, opts = {}) => new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    width: { size: width, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text ?? "N/A"), font: REPORT_FONT, size: opts.size || TABLE_FONT_SIZE, bold: !!opts.bold, color: opts.color || "000000" })],
      spacing: { before: 0, after: 0 },
    })],
  });

  const hdrRow = new TableRow({ tableHeader: true, cantSplit: true, children: hdrs.map((h, i) => mkHdr(h, CW[i])) });
  const dataRows = cvms.map((cvm, i) => {
    const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    const powerText  = String(cvm.cvm_power || "N/A").toUpperCase() === "ON" ? "On" : "Off";
    const powerColor = powerText === "On" ? "00843D" : "CC0000";
    return new TableRow({ cantSplit: true, children: [
      mkCell(cvm.cvm_name || "N/A", CW[0], { fill: bg, bold: true }),
      mkCell(cvm.host_name || "N/A", CW[1], { fill: bg, bold: true }),
      mkCell(cvm.cvm_ip || "N/A", CW[2], { fill: bg }),
      mkCell(cvm.cvm_memory_gib ?? "N/A", CW[3], { fill: bg, center: true }),
      mkCell(cvm.cvm_vcpus ?? "N/A", CW[4], { fill: bg, center: true }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
        shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, width: { size: CW[5], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: powerText, font: REPORT_FONT, size: TABLE_FONT_SIZE, bold: true, color: powerColor })], spacing: { before: 0, after: 0 } })],
      }),
    ]});
  });

  return [
    heading2("Controller VM Inventory"),
    new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [hdrRow, ...dataRows] }),
  ];
}

function summaryTable() {
  const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Section", 3500), hdrCell("Status", 1900), hdrCell("Section", 3500), hdrCell("Status", 1900)] })];
  for (let i = 0; i < summaryStatuses.length; i += 2) {
    const left = summaryStatuses[i];
    const right = summaryStatuses[i + 1] || ["", ""];
    const shade = (i / 2) % 2 === 0 ? "FFFFFF" : ROW_ALT;
    rows.push(new TableRow({ cantSplit: true, children: [
      sectionLinkCell(left[0], 3500, shade),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 1900, type: WidthType.DXA }, shading: { fill: shade, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [statusRun(left[1])], spacing: { before: 0, after: 0 } })] }),
      sectionLinkCell(right[0], 3500, shade),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 1900, type: WidthType.DXA }, shading: { fill: shade, type: ShadingType.CLEAR }, children: [new Paragraph({ children: right[0] ? [statusRun(right[1])] : [], spacing: { before: 0, after: 0 } })] }),
    ]}));
  }
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [3500, 1900, 3500, 1900], rows });
}

function correlatedAlertTable(alerts) {
  const list = Array.isArray(alerts) ? alerts : [];
  if (list.length === 0) return body("No correlated active alerts detected.", { italic: true });
  const CW = [1100, 5300, 2400, 2000];
  const hdrs = ["Severity", "Title", "Source Host", "Last Occurred"];
  const sevColor = (sev) => String(sev || "").toUpperCase() === "CRITICAL" ? "CC0000" : String(sev || "").toUpperCase() === "WARNING" ? "E5A000" : "666666";
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: hdrs.map((h, i) => hdrCell(h, CW[i])) }),
      ...list.map((a, i) => {
        const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
        return new TableRow({ cantSplit: true, children: [
          cell(String(a.severity || "UNKNOWN").toUpperCase(), { width: CW[0], shading: bg, bold: true, color: sevColor(a.severity) }),
          cell(a.title || "Alert", { width: CW[1], shading: bg }),
          cell(a.source_host || a.host || "N/A", { width: CW[2], shading: bg }),
          cell(a.last_occurred || "N/A", { width: CW[3], shading: bg }),
        ]});
      }),
    ],
  });
}

function securityStatusCell(status, width, shading) {
  const value = String(status || "N/A");
  const runs = STATUS_COLORS[value]
    ? [statusRun(value)]
    : [new TextRun({ text: value, font: REPORT_FONT, size: TABLE_FONT_SIZE, color: "666666" })];
  return new TableCell({
    borders: BORDERS,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: runs, spacing: { before: 0, after: 0 } })],
  });
}

function securityConfigurationSummaryTable() {
  const items = Array.isArray(D.security.configuration_items) ? D.security.configuration_items : [];
  const CW = [4200, 4200, 2400];
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Setting", CW[0]), hdrCell("Value", CW[1]), hdrCell("Status", CW[2])] }),
      ...items.map((item, i) => {
        const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
        return new TableRow({ cantSplit: true, children: [
          cell(item.item || "N/A", { width: CW[0], shading: bg, bold: true }),
          cell(item.value || "N/A", { width: CW[1], shading: bg }),
          securityStatusCell(item.status, CW[2], bg),
        ]});
      }),
    ],
  });
}

function securityHardeningSummaryTable() {
  const items = Array.isArray(D.security.security_hardening_items) ? D.security.security_hardening_items : [];
  const CW = [4800, 3600, 2400];
  if (items.length === 0) return body("Security hardening information was not available from the collected Prism APIs.", { italic: true });
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Hardening Control", CW[0]), hdrCell("Value", CW[1]), hdrCell("Status", CW[2])] }),
      ...items.map((item, i) => {
        const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
        return new TableRow({ cantSplit: true, children: [
          cell(item.item || "N/A", { width: CW[0], shading: bg, bold: true }),
          cell(item.value || "N/A", { width: CW[1], shading: bg }),
          securityStatusCell(item.status, CW[2], bg),
        ]});
      }),
    ],
  });
}

function hostSecureBootSummaryTable() {
  const hosts = Array.isArray(D.security.host_secure_boot) ? D.security.host_secure_boot : [];
  const CW = [3400, 2800, 2800, 1800];
  if (hosts.length === 0) return body("Host Secure Boot information was not available from the collected Prism Central APIs.", { italic: true });
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Host", CW[0]), hdrCell("Model", CW[1]), hdrCell("Secure Boot", CW[2]), hdrCell("Status", CW[3])] }),
      ...hosts.map((host, i) => {
        const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
        return new TableRow({ cantSplit: true, children: [
          cell(host.host || "N/A", { width: CW[0], shading: bg, bold: true }),
          cell(host.model || "N/A", { width: CW[1], shading: bg }),
          cell(host.secure_boot || "N/A", { width: CW[2], shading: bg }),
          securityStatusCell(host.status, CW[3], bg),
        ]});
      }),
    ],
  });
}

function encryptionSecuritySummaryTable() {
  const enabled = String(D.security.storage_encryption || "").toLowerCase() === "enabled";
  const rows = [["Encryption Status", D.security.storage_encryption || "N/A"]];
  if (enabled) {
    rows.push(["Encryption Type", D.security.encryption_type || "N/A"]);
    rows.push(["Key Management", D.security.key_management || "N/A"]);
    if (String(D.security.key_management || "") === "External KMS") {
      rows.push(["External KMS", D.security.external_kms || "N/A"]);
    }
    rows.push(["Encrypted User Containers", String(D.security.encrypted_containers ?? 0)]);
  }
  const CW = [4200, 6600];
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", CW[0]), hdrCell("Value", CW[1])] }),
      ...rows.map(([item, value], i) => {
        const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
        return new TableRow({ cantSplit: true, children: [
          cell(item, { width: CW[0], shading: bg, bold: true }),
          cell(value, { width: CW[1], shading: bg }),
        ]});
      }),
    ],
  });
}

function networkHealthSummaryTable() {
  const rows = [
    ["Cluster VIP", D.network.cluster_vip || "N/A"],
    ["Data Services IP", D.network.data_services_ip || "N/A"],
    ["External Subnet", D.network.external_subnet || "N/A"],
    ["Internal Subnet", D.network.internal_subnet || "N/A"],
    ["VLANs Discovered", String(D.network.vlan_count || 0)],
    ["Virtual Switches", String(D.network.virtual_switch_count || 0)],
    ["OVS Bridges", String(D.network.bridge_count || 0)],
    ["Network Bonds", String(D.network.bond_count || 0)],
    ["Physical NICs", String(D.network.nic_count || 0)],
    ["Active Network Alerts", String(D.network.network_alert_count || 0)],
  ];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [4200, 6600], rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", 4200), hdrCell("Value", 6600)] }),
    ...rows.map(([k, v], i) => new TableRow({ cantSplit: true, children: [
      cell(k, { width: 4200, shading: i % 2 === 0 ? NUTANIX_LIGHT : "FFFFFF", bold: true }),
      cell(v, { width: 6600, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
    ]})),
  ]});
}

function networkStatusRun(value) {
  const text = String(value || "").toLowerCase();
  if (["ignore", "ignored"].some(x => text === x)) return new TextRun({ text: "Ignored", font: REPORT_FONT, size: TABLE_FONT_SIZE, bold: true, color: "666666" });
  if (["normal", "healthy", "up", "active", "connected"].some(x => text.includes(x))) return statusRun("Healthy");
  if (["warning", "degraded", "unknown", "inactive"].some(x => text.includes(x))) return statusRun("Recommended");
  if (["critical", "down", "failed", "error", "disconnected"].some(x => text.includes(x))) return statusRun("Critical");
  return new TextRun({ text: String(value || "N/A"), font: REPORT_FONT, size: TABLE_FONT_SIZE });
}

function hostNetworkSummaryTable() {
  const list = D.network.host_ip_summary || [];
  if (!list.length) return body("No host network summary data available.");
  const CW = [2600, 2100, 2100, 2100, 1900];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: ["Host", "AHV IP", "CVM IP", "IPMI IP", "Status"].map((h, i) => hdrCell(h, CW[i])) }),
    ...list.map((n, i) => new TableRow({ cantSplit: true, children: [
      cell(n.host || "N/A", { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(n.ahv_ip || "N/A", { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(n.cvm_ip || "N/A", { width: CW[2], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(n.ipmi_ip || "N/A", { width: CW[3], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: CW[4], type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [networkStatusRun(n.status)], spacing: { before: 0, after: 0 } })] }),
    ]})),
  ]});
}

function vlanSummaryTable() {
  if (!D.network.networks || D.network.networks.length === 0) return body("No VLAN/subnet data available.");
  const CW = [3000, 1200, 1300, 3600, 1700];
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: ["Network Name", "VLAN ID", "Bridge", "Virtual Switch", "IP Assignment"].map((h, i) => hdrCell(h, CW[i])) }),
      ...D.network.networks.map((n, i) => new TableRow({ cantSplit: true, children: [
        cell(n.name || "N/A",                     { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(String(n.vlan_id || "N/A"),          { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(n.bridge || "N/A",                   { width: CW[2], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(n.virtual_switch || "N/A",           { width: CW[3], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(n.ip_assignment || "External IPAM",  { width: CW[4], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      ]})),
    ],
  });
}

function bridgeSummaryTable() {
  const list = D.network.bridge_summary || [];
  if (!list.length) return body("No OVS bridge data available from the collected APIs.", { italic: true });
  const CW = [2500, 3300, 3800, 1200];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: ["Bridge", "VLANs", "Virtual Switch", "Status"].map((h, i) => hdrCell(h, CW[i])) }),
    ...list.map((b, i) => new TableRow({ cantSplit: true, children: [
      cell(b.bridge || "N/A", { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(b.vlans || "N/A", { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(b.virtual_switch || "N/A", { width: CW[2], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: CW[3], type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [networkStatusRun(b.status)], spacing: { before: 0, after: 0 } })] }),
    ]})),
  ]});
}

function bondSummaryTable() {
  const list = D.network.bond_summary || [];
  if (!list.length) return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [3600, 7200], rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", 3600), hdrCell("Result", 7200)] }),
    new TableRow({ cantSplit: true, children: [
      cell("Bond Information", { width: 3600, shading: "FFFFFF", bold: true }),
      cell("Not available from collected Prism Central APIs", { width: 7200, shading: ROW_ALT }),
    ]}),
  ]});
  const CW = [2300, 2000, 2000, 3300, 1200];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: ["Host", "Bond", "Mode", "Members", "Status"].map((h, i) => hdrCell(h, CW[i])) }),
    ...list.map((b, i) => new TableRow({ cantSplit: true, children: [
      cell(b.host || "N/A", { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(b.bond || "N/A", { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(b.mode || "N/A"), { width: CW[2], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(b.members || "N/A"), { width: CW[3], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: CW[4], type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [networkStatusRun(b.status)], spacing: { before: 0, after: 0 } })] }),
    ]})),
  ]});
}

function nicSummaryTable() {
  const list = D.network.nic_summary || [];
  if (!list.length) return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [3600, 7200], rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", 3600), hdrCell("Result", 7200)] }),
    new TableRow({ cantSplit: true, children: [
      cell("Physical NIC Information", { width: 3600, shading: "FFFFFF", bold: true }),
      cell("Not available from collected Prism Central APIs", { width: 7200, shading: ROW_ALT }),
    ]}),
  ]});
  const CW = [1800, 900, 900, 1200, 1200, 2200, 800, 1800];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: ["Host", "NIC", "Link", "Speed", "Capacity", "MAC Address", "MTU", "Status"].map((h, i) => hdrCell(h, CW[i])) }),
    ...list.map((n, i) => new TableRow({ cantSplit: true, children: [
      cell(n.host || "N/A", { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(n.interface || "N/A", { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(n.link || "N/A"), { width: CW[2], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(n.speed || "—"), { width: CW[3], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(n.capacity || "N/A"), { width: CW[4], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(n.mac || "N/A"), { width: CW[5], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      cell(String(n.mtu || "N/A"), { width: CW[6], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: CW[7], type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [networkStatusRun(n.status)], spacing: { before: 0, after: 0 } })] }),
    ]})),
  ]});
}

function dnsNtpSummaryTable() {
  const rows = [
    ["DNS Servers", (D.network.dns_servers || []).join(", ") || "N/A"],
    ["NTP Servers", (D.network.ntp_servers || []).join(", ") || "N/A"],
  ];
  const CW = [2600, 8200];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: ["Setting", "Value"].map((h, i) => hdrCell(h, CW[i])) }),
    ...rows.map(([setting, value], i) => new TableRow({ cantSplit: true, children: [
      cell(setting, { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(value, { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
    ]})),
  ]});
}

function fmt(val, suffix) {
  if (val === null || val === undefined || val === "N/A") return "N/A";
  return String(val) + (suffix ? " " + suffix : "");
}

function pct2(val) {
  const n = pctNumber(val);
  if (n === null) return "N/A";
  return n.toFixed(2) + "%";
}

function containerDetailTables() {
  const sd = D.storage.storage_detail || {};
  const containers = D.storage.containers || [];
  const out = [];

  // ── Per-container configuration detail ────────────────────────────────
  if (containers.length > 0) {
    out.push(heading2("Storage Container Configuration"));
    for (const c of containers) {
      out.push(new Table({
    layout: TableLayoutType.AUTOFIT,
        width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [5400, 5400],
        rows: [
          new TableRow({ cantSplit: true, children: [
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } },
              shading: { fill: NUTANIX_BLUE, type: ShadingType.CLEAR },
              margins: { top: 60, bottom: 60, left: 120, right: 120 },
              columnSpan: 2,
              width: { size: CONTENT_WIDTH, type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: "STORAGE CONTAINER DETAILS — " + c.name, font: REPORT_FONT, bold: true, size: TABLE_HEADER_SIZE, color: "FFFFFF" })] })],
            }),
          ]}),
          ...[
            ["Name",                  c.name],
            ["Replication Factor",    String(c.rf)],
            ["Max Capacity (Physical)", fmt(c.max_capacity_tib !== "N/A" ? c.max_capacity_tib : sd.max_capacity_tib, "TiB")],
            ["Reserved Capacity (Logical)", c.reserved_capacity_logical || "N/A"],
            ["Advertised Capacity (Logical)", c.advertised_capacity_logical || "N/A"],
            ["Free Capacity (Physical)", fmt(sd.free_capacity_tib, "TiB")],
            ["Used (Physical)",       fmt(sd.used_physical_gib, "GiB")],
            ["Snapshot Usage",        fmt(sd.snapshot_gib, "GiB")],
            ["Logical Usage",         fmt(sd.logical_usage_gib, "GiB")],
            ["Data Reduction Savings", fmt(sd.savings_gib, "GiB")],
            ["Data Reduction Ratio",  sd.data_reduction_ratio ? sd.data_reduction_ratio + " : 1" : "N/A"],
            ["Overall Efficiency",    sd.savings_ratio ? sd.savings_ratio + " : 1" : "N/A"],
            ["Recycle Bin",           fmt(sd.recycle_bin_gib, "GiB")],
            ["Compression",           c.compression ? "On" : "Off"],
            ["Compression Delay",     c.compression_delay || "N/A"],
            ["Capacity Deduplication",c.dedup && c.dedup !== "OFF" ? "On (" + c.dedup + ")" : "Off"],
            ["Erasure Coding",        c.erasure_code && c.erasure_code !== "OFF" ? "On" : "Off"],
            ["Software Encryption",   c.software_encryption || "N/A"],
          ].map(([label, value], i) => new TableRow({ cantSplit: true, children: [
            cell(label, { width: 5400, shading: i % 2 === 0 ? NUTANIX_LIGHT : "FFFFFF", bold: true }),
            cell(value, { width: 5400, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
          ]})),
        ],
      }));
    }
  }
  return out;
}



// ── Executive dashboard and insight helpers ─────────────────────────────
function pctNumber(v) {
  if (v === null || v === undefined || v === "N/A" || v === "Not available") return null;
  const n = Number(String(v).replace("%", ""));
  return Number.isFinite(n) ? n : null;
}

function boolText(v) {
  if (v === true) return "Yes";
  if (v === false) return "No";
  if (v === "true") return "Yes";
  if (v === "false") return "No";
  return (v === null || v === undefined || v === "") ? "N/A" : String(v);
}

function severityRank(sev) {
  const s = String(sev || "").toUpperCase();
  if (s.includes("CRITICAL")) return 3;
  if (s.includes("WARNING")) return 2;
  if (s.includes("INFO")) return 1;
  return 0;
}

function statusPenalty(status) {
  if (status === "Critical") return 8;
  if (status === "Recommended") return 3;
  return 0;
}

function executiveHealthScore() {
  let score = 100;
  score -= Math.min((D.health.critical_alerts || 0) * 6, 36);
  score -= Math.min((D.health.warning_alerts || 0) * 3, 18);
  for (const [_, st] of summaryStatuses) score -= statusPenalty(st);
  const storagePct = pctNumber(D.storage.disk_utilization_pct);
  const cpuPct = pctNumber(D.cpu.average_cpu_usage_pct);
  const memPct = pctNumber(D.memory.average_memory_usage_pct);
  if (storagePct !== null && storagePct >= 80) score -= 8;
  if (cpuPct !== null && cpuPct >= 80) score -= 6;
  if (memPct !== null && memPct >= 80) score -= 6;
  return Math.max(0, Math.min(100, Math.round(score)));
}

function scoreLabel(score) {
  if (score >= 90) return "Healthy";
  if (score >= 75) return "Healthy with Recommendations";
  if (score >= 60) return "Needs Attention";
  return "Action Required";
}

function scoreStatus(score) {
  if (score >= 90) return "Healthy";
  if (score >= 70) return "Recommended";
  return "Critical";
}

function barText(pct, blocks = 10) {
  const n = pctNumber(pct);
  if (n === null) return "N/A";
  const filled = Math.max(0, Math.min(blocks, Math.round((n / 100) * blocks)));
  return "█".repeat(filled) + "░".repeat(blocks - filled) + ` ${n.toFixed(1)}%`;
}

function countBar(n, max) {
  const blocks = max > 0 ? Math.max(1, Math.round((n / max) * 10)) : 0;
  return "█".repeat(blocks) + "░".repeat(10 - blocks) + ` ${n}`;
}

function severityColor(sev) {
  const s = String(sev || "").toUpperCase();
  if (s.includes("CRITICAL")) return "F8D7DA";
  if (s.includes("WARNING")) return "FFF3CD";
  if (s.includes("INFO")) return "D6E8F7";
  return "FFFFFF";
}

function priorityColor(priority) {
  if (priority === "High") return "F8D7DA";
  if (priority === "Medium") return "FFF3CD";
  if (priority === "Low") return "D6E8F7";
  return "FFFFFF";
}

function executiveHealthScoreTable() {
  const score = executiveHealthScore();
  const label = scoreLabel(score);
  const status = scoreStatus(score);
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [3600, 7200],
    rows: [
      new TableRow({ cantSplit: true, children: [
        new TableCell({ borders: BORDERS, shading: { fill: NUTANIX_BLUE, type: ShadingType.CLEAR }, margins: { top: 180, bottom: 180, left: 160, right: 160 }, width: { size: 3600, type: WidthType.DXA }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(score), font: REPORT_FONT, bold: true, size: 72, color: "FFFFFF" })] }), new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "/ 100", font: REPORT_FONT, bold: true, size: 22, color: "FFFFFF" })] })] }),
        new TableCell({ borders: BORDERS, margins: { top: 180, bottom: 180, left: 180, right: 180 }, width: { size: 7200, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: label, font: REPORT_FONT, bold: true, size: H2_SIZE, color: STATUS_COLORS[status] || HEADER_GREY })], spacing: { after: PARA_AFTER } }), new Paragraph({ children: [new TextRun({ text: `Critical alerts: ${D.health.critical_alerts || 0}    Warnings: ${D.health.warning_alerts || 0}    Total active alerts: ${D.health.total_alerts || 0}`, font: REPORT_FONT, size: REPORT_FONT_SIZE })] }), new Paragraph({ children: [new TextRun({ text: "Score is calculated from active alerts, section statuses, and capacity/performance thresholds.", font: REPORT_FONT, size: REPORT_FONT_SIZE, italics: true, color: HEADER_GREY })] })] }),
      ]}),
    ],
  });
}

function healthCategoryTable() {
  const categories = [
    ["Cluster Health", D.health.status],
    ["Configuration", D.security.status],
    ["Security", (D.health.alert_details || []).some(a => /default password|password based ssh|password-based ssh/i.test(a.title || "")) ? "Critical" : D.security.status],
    ["Capacity", D.storage.status],
    ["Performance", (D.cpu.status === "Critical" || D.memory.status === "Critical") ? "Critical" : (D.cpu.status === "Recommended" || D.memory.status === "Recommended" ? "Recommended" : "Healthy")],
    ["Availability", D.protection.status],
  ];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [5400, 5400], rows: [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Category", 5400), hdrCell("Status", 5400)] }), ...categories.map(([cat, st], i) => new TableRow({ cantSplit: true, children: [cell(cat, { width: 5400, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }), new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 5400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [statusRun(st)] })] })] }))] });
}

function alertSeveritySummaryTable() {
  const critical = D.health.critical_alerts || 0;
  const warning = D.health.warning_alerts || 0;
  const info = Math.max(0, (D.health.total_alerts || 0) - critical - warning);
  const max = Math.max(critical, warning, info, 1);
  const rows = [["Critical", critical], ["Warning", warning], ["Info", info]];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [2500, 8300], rows: [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Severity", 2500), hdrCell("Active Alert Count", 8300)] }), ...rows.map(([label, count], i) => new TableRow({ cantSplit: true, children: [cell(label, { width: 2500, shading: severityColor(label.toUpperCase()), bold: true }), cell(countBar(count, max), { width: 8300, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT })] }))] });
}

function alertObservationText() {
  const total = D.health.total_alerts || 0;
  const critical = D.health.critical_alerts || 0;
  const warning = D.health.warning_alerts || 0;
  if (critical > 0 || warning > 0) {
    const noun = total === 1 ? "active alert was" : "active alerts were";
    return `${total} ${noun} detected during the health assessment. Review the alert details and recommended actions below.`;
  }
  return "No Critical or Warning alerts were detected during the health assessment.";
}

function alertOverviewTable() {
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2300, 8500],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Item", 2300), hdrCell("Comments", 8500)] }),
      new TableRow({ cantSplit: true, children: [
        cell("Status", { width: 2300, shading: NUTANIX_LIGHT, bold: true }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 8500, type: WidthType.DXA },
          children: [new Paragraph({ children: [statusRun(D.health.status)], spacing: { before: 0, after: 0 } })],
        }),
      ]}),
      new TableRow({ cantSplit: true, children: [
        cell("Observation", { width: 2300, shading: NUTANIX_LIGHT, bold: true }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 8500, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: alertObservationText(), font: REPORT_FONT, size: REPORT_FONT_SIZE })], spacing: { before: 0, after: 0 } })],
        }),
      ]}),
    ],
  });
}

function capacityDashboardTable() {
  const rows = [
    ["CPU", D.cpu.average_cpu_usage_pct, D.cpu.status],
    ["Memory", D.memory.average_memory_usage_pct, D.memory.memory_utilization_status || D.memory.status],
    ["Storage", D.storage.disk_utilization_pct, D.storage.storage_utilization_status || D.storage.status],
  ];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [2200, 6200, 2400], rows: [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Resource", 2200), hdrCell("Utilization", 6200), hdrCell("Status", 2400)] }), ...rows.map(([name, pct, st], i) => new TableRow({ cantSplit: true, children: [cell(name, { width: 2200, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }), cell(barText(pct), { width: 6200, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }), new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 2400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [statusRun(st)] })] })] }))] });
}


function cpuSummaryTable() {
  return twoColTable([
    ["Physical Hosts", D.cpu.physical_hosts || "N/A"],
    ["Total CPU Sockets", D.cpu.total_cpu_sockets || "N/A"],
    ["Physical Cores", D.cpu.physical_cores || "N/A"],
    ["Logical CPUs", D.cpu.logical_cpus || "N/A"],
  ], [4200, 6600]);
}

function cpuOversubscriptionTable() {
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [3600, 3600, 3600],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Metric", 3600), hdrCell("Value", 3600), hdrCell("Status", 3600)] }),
      new TableRow({ cantSplit: true, children: [
        cell("Cluster vCPU:pCore Ratio", { width: 3600 }),
        cell(D.cpu.vcpu_pcore_ratio || "N/A", { width: 3600 }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 3600, type: WidthType.DXA }, children: [new Paragraph({ children: [statusRun(D.cpu.oversubscription_status || "Healthy")], spacing: { before: 0, after: 0 } })] }),
      ]}),
    ],
  });
}

function cpuHeadroomTable() {
  return twoColTable([
    ["Average Utilization", pct2(D.cpu.average_cpu_usage_pct)],
    ["Peak Utilization", pct2(D.cpu.peak_cpu_usage_pct)],
    ["Available CPU Headroom", pct2(D.cpu.cpu_headroom_pct)],
  ], [4200, 6600]);
}

function cpuDistributionTable() {
  const rows = D.cpu.host_cpu_distribution || [];
  if (!rows.length) {
    return body("No host-level CPU allocation data available.");
  }
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2600, 1500, 1400, 1900, 1700, 1700],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Host", 2600), hdrCell("User VM vCPUs", 1500), hdrCell("CVM vCPUs", 1400), hdrCell("Total vCPUs Allocated", 1900), hdrCell("Physical Cores Per Host", 1700), hdrCell("vCPU:pCore Per Host", 1700)] }),
      ...rows.map((r, i) => new TableRow({ cantSplit: true, children: [
        cell(r.host || "Unknown", { width: 2600, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
        cell(r.user_vm_vcpus ?? "N/A", { width: 1500, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.cvm_vcpus ?? "N/A", { width: 1400, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.vcpus ?? "N/A", { width: 1900, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.physical_cores ?? "N/A", { width: 1700, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.vcpu_pcore_ratio || "N/A", { width: 1700, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      ]})),
    ],
  });
}


function memorySummaryTable() {
  return twoColTable([
    ["Physical Hosts", D.memory.physical_hosts || "N/A"],
    ["Total Physical Memory", D.memory.total_physical_memory_gib !== "N/A" ? D.memory.total_physical_memory_gib + " GiB" : "N/A"],
    ["Total User VM Allocated Memory", D.memory.total_user_vm_memory_gib !== undefined ? D.memory.total_user_vm_memory_gib + " GiB" : "N/A"],
    ["Total CVM Allocated Memory", D.memory.total_cvm_memory_gib !== undefined ? D.memory.total_cvm_memory_gib + " GiB" : "N/A"],
    ["Total Allocated Memory", D.memory.total_vm_memory_gib !== undefined ? D.memory.total_vm_memory_gib + " GiB" : "N/A"],
      ], [4200, 6600]);
}

function memoryAllocationTable() {
  const rows = D.memory.host_memory_distribution || [];
  if (!rows.length) {
    return body("No host-level memory allocation data available.");
  }
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2200, 1000, 1700, 1400, 1700, 1800, 1000],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Host", 2200), hdrCell("User VMs", 1000), hdrCell("User VM Allocation", 1700), hdrCell("CVM Memory", 1400), hdrCell("Total Memory Allocated", 1700), hdrCell("Physical Memory Per Host", 1800), hdrCell("Percent Allocated", 1000)] }),
      ...rows.map((r, i) => new TableRow({ cantSplit: true, children: [
        cell(r.host || "Unknown", { width: 2200, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
        cell(r.vm_count ?? "N/A", { width: 1000, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.user_vm_memory_gib !== undefined ? r.user_vm_memory_gib + " GiB" : "N/A", { width: 1700, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.cvm_memory_gib !== undefined ? r.cvm_memory_gib + " GiB" : "N/A", { width: 1400, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.allocated_memory_gib !== undefined ? r.allocated_memory_gib + " GiB" : "N/A", { width: 1700, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.physical_memory_gib !== "N/A" ? r.physical_memory_gib + " GiB" : "N/A", { width: 1800, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        cell(r.allocation_pct || "N/A", { width: 1000, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      ]})),
    ],
  });
}

function memoryAllocationStatusTable() {
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [3600, 3600, 3600],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Metric", 3600), hdrCell("Value", 3600), hdrCell("Status", 3600)] }),
      new TableRow({ cantSplit: true, children: [
        cell("Total Allocated Memory / Physical Memory", { width: 3600 }),
        cell(D.memory.memory_allocation_pct || "N/A", { width: 3600 }),
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 3600, type: WidthType.DXA }, children: [new Paragraph({ children: [statusRun(D.memory.memory_allocation_status || "Healthy")], spacing: { before: 0, after: 0 } })] }),
      ]}),
    ],
  });
}

function memoryHeadroomTable() {
  return twoColTable([
    ["Average Utilization", D.memory.average_memory_usage_pct !== "N/A" ? D.memory.average_memory_usage_pct + "%" : "N/A"],
    ["Peak Utilization", D.memory.peak_memory_usage_pct !== "N/A" ? D.memory.peak_memory_usage_pct + "%" : "N/A"],
    ["Available Memory Headroom", D.memory.memory_headroom_pct !== "N/A" ? D.memory.memory_headroom_pct + "%" : "N/A"],
  ], [4200, 6600]);
}

function assessmentStatusTable() {
  const rows = [
    ["Healthy", "No recommendations required."],
    ["Recommended", "Apply during the next maintenance window."],
    ["Critical", "Apply as soon as possible."],
  ];
  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [3000, 7800],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Status", 3000), hdrCell("Guidance", 7800)] }),
      ...rows.map(([status, guidance], i) => new TableRow({ cantSplit: true, children: [
        new TableCell({ borders: BORDERS, margins: { top: 60, bottom: 60, left: 120, right: 120 }, width: { size: 3000, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "FFFFFF" : ROW_ALT, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [statusRun(status)], spacing: { before: 0, after: 0 } })] }),
        cell(guidance, { width: 7800, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
      ]}))
    ],
  });
}

function clusterConfigSummaryTable() {
  const sd = D.storage.storage_detail || {};
  return twoColTable([
    ["Nodes", String(C.node_count || "N/A")],
    ["Hypervisor", C.hypervisor || "N/A"],
    ["AOS Version", C.aos_version || "N/A"],
    ["AHV Version", C.ahv_version || "N/A"],
    ["NCC Version", C.ncc_version || "N/A"],
    ["Redundancy Factor", "RF" + String(C.redundancy_factor || "N/A")],
    ["Fault Tolerance", C.fault_tolerance || "N/A"],
    ["Storage Capacity", fmt(sd.max_capacity_tib || D.storage.pool_capacity_tib, "TiB")],
    ["Storage Utilization", D.storage.disk_utilization_pct !== "N/A" ? D.storage.disk_utilization_pct + "%" : "N/A"],
    ["Pulse Enabled", boolText(C.pulse_enabled)],
    ["Password Remote Login", boolText(C.password_remote_login_enabled)],
    ["License", `${D.licensing.license_name} (${D.licensing.license_type})`],
  ], [3600, 7200]);
}

function recommendationFromAlert(a) {
  const title = String(a.title || "Alert");
  const sev = String(a.severity || "INFO").toUpperCase();
  let priority = sev.includes("CRITICAL") ? "High" : sev.includes("WARNING") ? "Medium" : "Low";
  let recommendation = "Review and remediate the active alert in Prism Central.";
  if (/default password/i.test(title)) recommendation = "Change default credentials immediately.";
  else if (/password.*ssh|ssh.*password/i.test(title)) recommendation = "Disable password-based SSH where appropriate and use key-based access.";
  else if (/DIMM/i.test(title)) recommendation = "Validate DIMM population against the platform hardware guide and correct during maintenance.";
  else if (/disk|unqualified|mounted/i.test(title)) recommendation = "Investigate disk qualification, mount state, and hardware health before adding workload.";
  else if (/snapshot/i.test(title)) recommendation = "Review snapshot and protection policy configuration for the affected entity.";
  else if (/latency/i.test(title)) recommendation = "Review network path and remote availability zone latency with Nutanix Support if persistent.";
  return { priority, recommendation, reason: title };
}

function recommendedActions() {
  const out = [];
  const seen = new Set();
  const priorityOrder = { High: 0, Medium: 1, Low: 2 };

  for (const a of (D.health.alert_details || []).sort((a,b) => severityRank(b.severity) - severityRank(a.severity))) {
    const item = recommendationFromAlert(a);
    // Deduplicate by action so repeated alerts, such as multiple hosts using
    // default credentials, produce one clean remediation item.
    const key = item.priority + "|" + item.recommendation;
    if (!seen.has(key)) {
      out.push(item);
      seen.add(key);
    }
  }

  if (out.length === 0) {
    out.push({ priority: "Low", recommendation: "No immediate alert remediation required. Continue routine monitoring and scheduled NCC reviews.", reason: "No active alert findings were detected." });
  }

  return out
    .sort((a, b) => (priorityOrder[a.priority] ?? 99) - (priorityOrder[b.priority] ?? 99) || String(a.recommendation).localeCompare(String(b.recommendation)))
    .slice(0, 15);
}

function recommendedActionsTable() {
  const actions = recommendedActions();
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [1600, 5200, 4000], rows: [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Priority", 1600), hdrCell("Recommendation", 5200), hdrCell("Reason", 4000)] }), ...actions.map((a, i) => new TableRow({ cantSplit: true, children: [cell(a.priority, { width: 1600, shading: priorityColor(a.priority), bold: true }), cell(a.recommendation, { width: 5200, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }), cell(a.reason, { width: 4000, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT })] }))] });
}

function protectionPolicySummaryTable() {
  const rows = Array.isArray(D.protection.policies) ? D.protection.policies : [];
  const CW = [3000, 1800, 1600, 1600, 1300, 1500];
  const dataRows = rows.length ? rows.map((p, i) => {
    const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    return new TableRow({ cantSplit: true, children: [
      cell(p.name || "Unnamed", { width: CW[0], shading: bg, bold: true }),
      cell(p.role || "Applicable", { width: CW[1], shading: bg }),
      cell(String(p.category_count ?? 0), { width: CW[2], shading: bg, center: true }),
      cell(String(p.schedule_count ?? 0), { width: CW[3], shading: bg, center: true }),
      cell(String(p.paused_count ?? 0), { width: CW[4], shading: bg, center: true }),
      securityStatusCell(p.status, CW[5], bg),
    ]});
  }) : [new TableRow({ cantSplit: true, children: [
    cell("N/A", { width: CW[0], bold: true }),
    cell("No applicable Protection Policies found", { width: CW[1] + CW[2] + CW[3] + CW[4] }),
    securityStatusCell("Recommended", CW[5], "FFFFFF"),
  ]})];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Policy", CW[0]), hdrCell("Cluster Role", CW[1]), hdrCell("Categories", CW[2]), hdrCell("Schedules", CW[3]), hdrCell("Paused", CW[4]), hdrCell("Status", CW[5])] }),
    ...dataRows,
  ]});
}

function protectionScheduleSummaryTable() {
  const rows = Array.isArray(D.protection.schedules) ? D.protection.schedules : [];
  const CW = [1700, 2800, 1200, 2400, 1200, 1500];
  const dataRows = rows.length ? rows.map((s, i) => {
    const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    return new TableRow({ cantSplit: true, children: [
      cell(s.policy || "Unnamed", { width: CW[0], shading: bg, bold: true }),
      cell(s.direction || "N/A", { width: CW[1], shading: bg }),
      cell(s.rpo || "N/A", { width: CW[2], shading: bg, center: true }),
      cell(s.retention || "N/A", { width: CW[3], shading: bg }),
      cell(s.recovery_point_type || "N/A", { width: CW[4], shading: bg }),
      securityStatusCell(s.status, CW[5], bg),
    ]});
  }) : [new TableRow({ cantSplit: true, children: [
    cell("N/A", { width: CW[0], bold: true }),
    cell("No applicable replication schedules found", { width: CW[1] + CW[2] + CW[3] + CW[4] }),
    securityStatusCell("N/A", CW[5], "FFFFFF"),
  ]})];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Policy", CW[0]), hdrCell("Direction", CW[1]), hdrCell("RPO", CW[2]), hdrCell("Retention", CW[3]), hdrCell("Type", CW[4]), hdrCell("Status", CW[5])] }),
    ...dataRows,
  ]});
}

function recoveryPlanSummaryTable() {
  const rows = Array.isArray(D.protection.recovery_plans) ? D.protection.recovery_plans : [];
  const CW = [2500, 2300, 2300, 1100, 1100, 1500];
  const dataRows = rows.length ? rows.map((p, i) => {
    const bg = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    return new TableRow({ cantSplit: true, children: [
      cell(p.name || "Unnamed", { width: CW[0], shading: bg, bold: true }),
      cell(p.primary_location || "N/A", { width: CW[1], shading: bg }),
      cell(p.recovery_location || "N/A", { width: CW[2], shading: bg }),
      cell(String(p.stage_count ?? 0), { width: CW[3], shading: bg, center: true }),
      cell(String(p.network_mapping_count ?? 0), { width: CW[4], shading: bg, center: true }),
      securityStatusCell(p.status, CW[5], bg),
    ]});
  }) : [new TableRow({ cantSplit: true, children: [
    cell("N/A", { width: CW[0], bold: true }),
    cell("No v4 Recovery Plans configured", { width: CW[1] + CW[2] + CW[3] + CW[4] }),
    securityStatusCell("N/A", CW[5], "FFFFFF"),
  ]})];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Recovery Plan", CW[0]), hdrCell("Primary", CW[1]), hdrCell("Recovery", CW[2]), hdrCell("Stages", CW[3]), hdrCell("Mappings", CW[4]), hdrCell("Status", CW[5])] }),
    ...dataRows,
  ]});
}

function ngtSummaryTable() {
  const vms = D.vms.vm_list || [];
  const ignoredNames = new Set(D.vms.system_vm_ignored || []);
  const counts = {
    totalUserVms: 0,
    installed: 0,
    needsNgt: 0,
    ignored: 0,
  };

  for (const vm of vms) {
    const isIgnored = ignoredNames.has(vm.name) || vm.os_support === "Ignored";
    const s = vm.ngt_status || "Unknown";

    if (isIgnored) {
      counts.ignored++;
      continue;
    }

    counts.totalUserVms++;
    if (s === "Enabled") counts.installed++;
    else counts.needsNgt++;
  }

  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2700, 2700, 2700, 2700],
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          hdrCell("Total User VMs", 2700),
          hdrCell("NGT Installed", 2700),
          hdrCell("Needs NGT", 2700),
          hdrCell("Ignored", 2700),
        ],
      }),
      new TableRow({
        cantSplit: true,
        children: [
          cell(counts.totalUserVms, { width: 2700, shading: "FFFFFF", bold: true }),
          cell(counts.installed, { width: 2700, shading: "DFF0D8", bold: true }),
          cell(counts.needsNgt, { width: 2700, shading: counts.needsNgt ? "FFF3CD" : "FFFFFF", bold: true }),
          cell(counts.ignored, { width: 2700, shading: counts.ignored ? "E2E3E5" : "FFFFFF", bold: true }),
        ],
      }),
    ],
  });
}

function storageEncryptionSummaryTable() {
  const e = D.storage.encryption || {};
  const rows = [
    ["Encryption Status", e.status || "N/A"],
    ["Encryption Type", e.encryption_type || "N/A"],
    ["Key Management", e.key_management || "N/A"],
    ["External KMS", e.external_kms || "N/A"],
  ];
  return twoColTable(rows, [4200, 6600]);
}

function storageEncryptionContainerTable() {
  const e = D.storage.encryption || {};
  const rows = e.container_rows || [];
  if (!rows.length) return null;
  const CW = [6200, 4600];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: CW, rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Storage Container", CW[0]), hdrCell("Software Encryption", CW[1])] }),
    ...rows.map((r, i) => new TableRow({ cantSplit: true, children: [
      cell(r.name || "N/A", { width: CW[0], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT, bold: true }),
      cell(r.software_encryption || "N/A", { width: CW[1], shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
    ]})),
  ]});
}

function diskHealthSummaryTable() {
  const diskAlerts = (D.health.alert_details || []).filter(a => /disk|drive|unqualified|mounted/i.test((a.title || "") + " " + (a.message || "")));
  const critical = diskAlerts.filter(a => String(a.severity).toUpperCase().includes("CRITICAL")).length;
  const warning = diskAlerts.filter(a => String(a.severity).toUpperCase().includes("WARNING")).length;
  const rows = [
    ["Active Disk/Drive Alerts", String(diskAlerts.length)],
    ["Critical Disk Alerts", String(critical)],
    ["Warning Disk Alerts", String(warning)],
    ["User Storage Containers Reviewed", String(D.storage.container_count || 0)],
  ];
  return twoColTable(rows, [4200, 6600]);
}

function complianceTable() {
  const alerts = D.health.alert_details || [];
  const hasDefaultPw = alerts.some(a => /default password/i.test(a.title || ""));
  const checks = [
    ["NTP configured", C.ntp_servers && C.ntp_servers !== "N/A", C.ntp_servers || "N/A"],
    ["DNS configured", C.dns_servers && C.dns_servers !== "N/A", C.dns_servers || "N/A"],
    ["Redundancy Factor configured", C.redundancy_factor && C.redundancy_factor !== "N/A", "RF" + String(C.redundancy_factor || "N/A")],
    ["Pulse enabled", C.pulse_enabled === true || C.pulse_enabled === "true", boolText(C.pulse_enabled)],
    ["Password SSH disabled", C.password_remote_login_enabled === false || C.password_remote_login_enabled === "false", boolText(C.password_remote_login_enabled)],
    ["Default passwords not detected", !hasDefaultPw, hasDefaultPw ? "Default password alerts active" : "No active default password alerts"],
    ["No critical alerts", (D.health.critical_alerts || 0) === 0, `${D.health.critical_alerts || 0} critical`],
    ["Protection policies configured", (D.protection.policy_count || 0) > 0, `${D.protection.policy_count || 0} policies`],
    ["License violations clear", (D.licensing.violations || []).length === 0, (D.licensing.violations || []).length ? "Violations detected" : "No violations"],
  ];
  return new Table({ layout: TableLayoutType.AUTOFIT, width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [900, 5200, 4700], rows: [new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Result", 900), hdrCell("Best Practice", 5200), hdrCell("Evidence", 4700)] }), ...checks.map(([label, pass, evidence], i) => new TableRow({ cantSplit: true, children: [cell(pass ? "PASS" : "REVIEW", { width: 900, shading: pass ? "DFF0D8" : "FFF3CD", bold: true }), cell(label, { width: 5200, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }), cell(evidence, { width: 4700, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT })] }))] });
}

function clusterTimelineTable() {
  const alerts = D.health.alert_details || [];
  const lastAlert = alerts.map(a => a.last_occurred).filter(Boolean).sort().reverse()[0] || "N/A";
  const latestPoint = (arr) => (Array.isArray(arr) && arr.length ? arr[arr.length - 1].timestamp || "N/A" : "N/A");
  return twoColTable([
    ["Report Generated", D.date],
    ["Last Active Alert Occurrence", lastAlert !== "N/A" ? lastAlert.replace("T", " ").replace("Z", " UTC") : "N/A"],
    ["Latest CPU Sample", latestPoint(D.cpu.cpu_history)],
    ["Latest Memory Sample", latestPoint(D.memory.mem_history)],
    ["Latest Storage Sample", latestPoint(D.storage.storage_history)],
    ["Cluster Upgrade Status", C.upgrade_status || "N/A"],
    ["License Expiry", D.licensing.expiry_date || "N/A"],
  ], [4200, 6600]);
}

function alertTable() {
  if (!D.health.alert_details || D.health.alert_details.length === 0) return body("No active alerts.");

  const rows = [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      hdrCell("Severity", 1100),
      hdrCell("Title", 3900),
      hdrCell("Source Host", 1500),
      hdrCell("Last Occurred", 1500),
      hdrCell("Classification", 1300),
      hdrCell("Impact Type", 1500),
    ]}),
  ];

  D.health.alert_details.forEach((a, i) => {
    const shade = i % 2 === 0 ? "FFFFFF" : ROW_ALT;
    rows.push(new TableRow({ cantSplit: true, children: [
      cell(a.severity || "UNKNOWN", { width: 1100, shading: severityColor(a.severity), bold: true }),
      cell(a.title || "Unknown alert", { width: 3900, shading: shade }),
      cell(a.source_host || "N/A", { width: 1500, shading: shade }),
      cell(a.last_occurred ? a.last_occurred.replace("T", " ").replace("Z", " UTC").substring(0, 20) : "N/A", { width: 1500, shading: shade }),
      cell(a.classification || "N/A", { width: 1300, shading: shade }),
      cell(a.impact_type || "N/A", { width: 1500, shading: shade }),
    ]}));
  });

  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [1100, 3900, 1500, 1500, 1300, 1500],
    rows,
  });
}

const children = [
  // COVER
  new Paragraph({ spacing: { before: 2880 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nutanix Cluster Health Check", font: REPORT_FONT, bold: true, size: 56, color: NUTANIX_BLUE })], spacing: { before: 0, after: 200 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: C.cluster_name, font: REPORT_FONT, size: 40, color: HEADER_GREY })], spacing: { before: 0, after: 120 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: D.customer, font: REPORT_FONT, size: 32, color: HEADER_GREY })], spacing: { before: 0, after: 120 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: D.date, font: REPORT_FONT, size: 26, color: "888888" }) ] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "[Confidential]", font: REPORT_FONT, size: REPORT_FONT_SIZE, italics: true, color: "AAAAAA" })], spacing: { before: 240 } }),
  pageBreak(),

  // EXECUTIVE SUMMARY
  heading1("Executive Overview"),
  compactBody(`${D.customer} engaged Professional Services to conduct a Nutanix Health Check of cluster ${C.cluster_name}. This report documents the discovery, analysis, and recommendations from the assessment.`),
  compactBody(`The Nutanix environment is a critical part of ${D.customer}'s virtual machine operations. This review confirms whether the environment is configured in a supported state aligned to Nutanix best practices.`),
  compactBody("Key Objectives:"),
  bulletItem("Confirmation of best practices implemented during deployment."),
  bulletItem("Implementation of Nutanix operational best practices."),
  bulletItem("Highlight any operational improvements."),
  bulletItem("Provide recommendations for improvement."),
  compactHeading2("Assessment Summary"),
  summaryTable(),
  compactHeading2("Overall Health Score"),
  executiveHealthScoreTable(),
  compactHeading2("Health Assessment by Category"),
  healthCategoryTable(),
  compactHeading2("Resource Utilization Summary"),
  capacityDashboardTable(),
  pageBreak(),

  // HEALTH CHECK BACKGROUND
  heading1("Health Check Background"),
  heading2("Scope"),
  body(`The Nutanix Health Check for ${D.customer} assesses the cluster ${C.cluster_name}'s current health and architecture, focusing on technical and operational aspects.`),
  twoColTable([
    ["Cluster Name",       C.cluster_name],
    ["Cluster UUID",       C.cluster_uuid],
    ["Cluster VIP",        C.cluster_vip || "N/A"],
    ["Data Services IP",   C.data_svc_ip || "N/A"],
    ["AOS Version",        C.aos_version],
    ["AHV Version",        C.ahv_version && C.ahv_version !== "N/A" ? C.ahv_version : "N/A"],
    ["NCC Version",        C.ncc_version !== "N/A" ? C.ncc_version : "Not retrieved"],
    ["Hypervisor",         C.hypervisor],
    ["Nodes",              String(C.node_count)],
    ["Cluster Type",       C.cluster_type || "N/A"],
    ["Arch",               C.cluster_arch || "N/A"],
    ["Fault Tolerance",    C.fault_tolerance || "N/A"],
    ["Redundancy Factor",  "RF" + String(C.redundancy_factor)],
    ["Time Zone",          C.timezone],
    ["DNS Servers",        C.dns_servers || "N/A"],
    ["NTP Servers",        C.ntp_servers || "N/A"],
    ["License",            D.licensing.license_name + " (" + D.licensing.license_type + ")"],
    ["License Expiry",     D.licensing.expiry_date],
    ["Report Date",        D.date],
  ]),
  heading2("Document History"),
  new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [2500, 3500, 4800],
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Date", 2500), hdrCell("Author", 3500), hdrCell("Note", 4800)] }),
      new TableRow({ cantSplit: true, children: [cell(D.date, { width: 2500 }), cell("Health Check Script", { width: 3500 }), cell("Auto-generated report", { width: 4800 })] }),
    ],
  }),
  heading2("Health Check Assessment and Recommendations"),
  assessmentStatusTable(),
  pageBreak(),


  // ALERTS
  heading1("Alerts Summary"),
  alertOverviewTable(),
  heading2("Alert Severity Summary"),
  alertSeveritySummaryTable(),
  heading2("Recommended Actions"),
  body("The following action list consolidates active alerts into an operational checklist."),
  recommendedActionsTable(),
  heading2("Active Alert Details"),
  alertTable(),
  pageBreak(),

  // VMs
  heading1("Virtual Machines Summary"),
  sectionTable("Virtual Machines", D.vms.status,
    "• User VMs: " + D.vms.total +
    "\n• Controller VM Summary (CVMs): " + ((D.vm_counts || {}).cvm_count || 0) +
    "\n• Total Powered On (VMs + CVMs): " + ((D.vm_counts || {}).total_on || D.vms.powered_on) +
    "\n• User VMs Powered Off: " + D.vms.powered_off +
    "\n• Supported Guest OS: " + ((D.vms.os_supported || []).length) +
    "\n• Legacy Guest OS: " + ((D.vms.os_legacy || []).length) +
    "\n• Unsupported Guest OS: " + ((D.vms.os_unsupported || []).length) +
    "\n• Nutanix System VMs Ignored: " + ((D.vms.system_vm_ignored || []).length),
    D.vms.recommendations),
  heading2("NGT Summary"),
  ngtSummaryTable(),
  ...vmDetailTable(),
  ...cvmTable(),
  pageBreak(),

  // DATA PROTECTION
  heading1("Data Protection Summary"),
  sectionTable("Data Protection", D.protection.status,
    `• Applicable Protection Policies: ${D.protection.policy_count || 0}\n` +
    `• Applicable Replication Schedules: ${D.protection.schedule_count || 0}\n` +
    `• Paused Replication Schedules: ${D.protection.paused_schedule_count || 0}\n` +
    `• Recovery Plans: ${D.protection.recovery_plan_count || 0}`,
    D.protection.recommendations),
  heading2("Protection Policy Summary"),
  protectionPolicySummaryTable(),
  heading2("Replication Schedule Summary"),
  protectionScheduleSummaryTable(),
  heading2("Recovery Plan Summary"),
  recoveryPlanSummaryTable(),
  pageBreak(),

  // CPU
  heading1("Cluster CPU Summary"),
  sectionTable("CPU", D.cpu.status,
    `• Average Cluster CPU Usage: ${pct2(D.cpu.average_cpu_usage_pct)}\n` +
    `• Peak Cluster CPU Usage: ${pct2(D.cpu.peak_cpu_usage_pct)}\n` +
    `• Available CPU Headroom: ${pct2(D.cpu.cpu_headroom_pct)}\n` +
    `• Cluster vCPU:pCore Ratio: ${D.cpu.vcpu_pcore_ratio || "N/A"}`,
    D.cpu.recommendations),
  ...cpuChart(D.cpu.cpu_history),
  heading2("CPU Allocation by Host"),
  cpuDistributionTable(),
  heading2("CPU Summary"),
  cpuSummaryTable(),
  heading2("CPU Oversubscription"),
  cpuOversubscriptionTable(),
  heading2("CPU Headroom"),
  cpuHeadroomTable(),
  pageBreak(),

  // MEMORY
  heading1("Cluster Memory Summary"),
  sectionTable("Memory", D.memory.status,
    `• Average Cluster Memory Usage: ${D.memory.average_memory_usage_pct !== "N/A" ? D.memory.average_memory_usage_pct + "%" : "N/A"}
` +
    `• Peak Cluster Memory Usage: ${D.memory.peak_memory_usage_pct !== "N/A" ? D.memory.peak_memory_usage_pct + "%" : "N/A"}
` +
    `• Available Memory Headroom: ${D.memory.memory_headroom_pct !== "N/A" ? D.memory.memory_headroom_pct + "%" : "N/A"}
` +
    `• Memory Allocation Percentage: ${D.memory.memory_allocation_pct || "N/A"}` +
    ((D.memory.memory_alert_count || 0) > 0 ? `\n• Active Memory Alerts: ${D.memory.memory_alert_count}` : ""),
    D.memory.recommendations),
  ...(((D.memory.memory_alerts || []).length > 0) ? [heading2("Active Memory Alerts"), correlatedAlertTable(D.memory.memory_alerts)] : []),
  ...memChart(D.memory.mem_history),
  heading2("Memory Allocation by Host"),
  memoryAllocationTable(),
  heading2("Memory Summary"),
  memorySummaryTable(),
  heading2("Memory Allocation Review"),
  memoryAllocationStatusTable(),
  heading2("Memory Headroom"),
  memoryHeadroomTable(),
  pageBreak(),

  // STORAGE
  heading1("Cluster Storage Summary"),
  sectionTable("Storage", D.storage.status,
    `• Disk Utilization: ${D.storage.disk_utilization_pct !== "N/A" ? D.storage.disk_utilization_pct + "%" : "Not available via stats API"}\n• Storage Pool Capacity: ${D.storage.pool_capacity_tib ? D.storage.pool_capacity_tib + " TiB" : "N/A"}\n• User Storage Containers: ${D.storage.container_count}` +
    ((D.storage.storage_alert_count || 0) > 0 ? `\n• Active Storage Alerts: ${D.storage.storage_alert_count}` : ""),
    D.storage.recommendations),
  ...(((D.storage.storage_alerts || []).length > 0) ? [heading2("Active Storage Alerts"), correlatedAlertTable(D.storage.storage_alerts)] : []),
  heading2("Storage Health Summary"),
  diskHealthSummaryTable(),
  ...storageChart(D.storage.storage_history),
  ...containerDetailTables(),
  pageBreak(),

  // NETWORK
  heading1("Network Summary"),
  sectionTable("Network", D.network.status,
    `• Hosts: ${(D.network.host_ip_summary || []).length}\n` +
    `• VLANs: ${D.network.vlan_count || 0}\n` +
    `• Virtual Switches: ${D.network.virtual_switch_count || 0}\n` +
    `• OVS Bridges: ${D.network.bridge_count || 0}\n` +
    `• Network Bonds: ${D.network.bond_count || 0}\n` +
    `• Physical NICs: ${D.network.nic_count || 0}\n` +
    ((D.network.network_alert_count || 0) > 0 ? `• Active Network Alerts: ${D.network.network_alert_count}` : `• Active Network Alerts: 0`),
    D.network.recommendations),
  ...(((D.network.network_alerts || []).length > 0) ? [heading2("Active Network Alerts"), correlatedAlertTable(D.network.network_alerts)] : []),
  heading2("Network Health Summary"),
  networkHealthSummaryTable(),
  heading2("Host Network Summary"),
  hostNetworkSummaryTable(),
  heading2("VLAN Summary"),
  vlanSummaryTable(),
  heading2("OVS Bridge Summary"),
  bridgeSummaryTable(),
  heading2("Bond Summary"),
  bondSummaryTable(),
  heading2("Physical NIC Summary"),
  nicSummaryTable(),
  heading2("DNS / NTP Summary"),
  dnsNtpSummaryTable(),
  pageBreak(),

  // LICENSING
  heading1("Licensing Summary"),
  sectionTable("Licensing", D.licensing.status,
    `• License: ${D.licensing.license_name} (${D.licensing.license_type})\n• Support/License Expiry: ${D.licensing.expiry_date}\n• Violations: ${D.licensing.violations.length === 0 ? "None" : D.licensing.violations.join(", ")}`,
    D.licensing.recommendations),
  pageBreak(),

  // SECURITY
  heading1("Security Summary"),
  sectionTable("Security", D.security.status,
    `• Security Configuration Checks: ${(D.security.configuration_items || []).length}\n` +
    `• Security Hardening Checks: ${(D.security.security_hardening_items || []).length}\n` +
    ((D.security.secure_boot_host_count || 0) > 0
      ? `• Eligible Hosts with Secure Boot Enabled: ${D.security.secure_boot_enabled_count || 0} of ${D.security.secure_boot_host_count}\n`
      : `• Secure Boot Eligibility: No G8-or-newer hosts detected\n`) +
    `• Data-at-Rest Encryption: ${D.security.storage_encryption || "N/A"}\n` +
    `• Active Security Alerts: ${D.security.security_alert_count || 0}\n` +
    `• Critical Security Alerts: ${D.security.critical_security_alert_count || 0}`,
    D.security.recommendations),
  ...(((D.security.security_alerts || []).length > 0) ? [heading2("Active Security Alerts"), correlatedAlertTable(D.security.security_alerts)] : []),
  heading2("Security Configuration Summary"),
  securityConfigurationSummaryTable(),
  pageBreak(),
  heading2("Security Hardening Summary"),
  securityHardeningSummaryTable(),
  heading2("Host Secure Boot Summary"),
  body(D.security.secure_boot_note || "Secure Boot support begins with Nutanix G8 platforms.", { italic: true }),
  hostSecureBootSummaryTable(),
  heading2("Data-at-Rest Encryption Summary"),
  encryptionSecuritySummaryTable(),
  pageBreak(),

  // SOFTWARE LIFECYCLE
  heading1("Software Lifecycle Summary"),
  sectionTable("Software Lifecycle", D.software_lifecycle.status,
    `• Current AOS Version: ${D.software_lifecycle.aos_version}\n` +
    `• AOS Lifecycle Status: ${(D.software_lifecycle.aos_lifecycle || {}).lifecycle_status || "Unknown"}\n` +
    `• Latest AOS Version: ${(D.software_lifecycle.aos_lifecycle || {}).latest_version || "N/A"}\n` +
    `• End of Maintenance: ${(D.software_lifecycle.aos_lifecycle || {}).end_of_maintenance || "N/A"}\n` +
    `• End of Support Life: ${(D.software_lifecycle.aos_lifecycle || {}).end_of_support_life || "N/A"}\n` +
    `• Upgrade Method: Life Cycle Manager (LCM)`,
    D.software_lifecycle.recommendations),
  pageBreak(),

  // NCC
  heading1("NCC Health Checks Summary"),
  sectionTable("NCC", D.ncc.status, `• NCC checks requiring review: ${D.ncc.check_count}`, D.ncc.recommendations),
  ...(D.ncc.checks.length > 0 ? [
    heading2("NCC Check Details"),
    new Table({
    layout: TableLayoutType.AUTOFIT,
      width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [8800, 2000],
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [hdrCell("Check", 8800), hdrCell("Severity", 2000)] }),
        ...D.ncc.checks.map((ch, i) => new TableRow({ cantSplit: true, children: [
          cell(ch.title,    { width: 8800, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
          cell(ch.severity, { width: 2000, shading: i % 2 === 0 ? "FFFFFF" : ROW_ALT }),
        ]})),
      ],
    }),
  ] : []),
];

const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  styles: {
    default: {
      document: { run: { font: REPORT_FONT, size: REPORT_FONT_SIZE } },
      paragraph: { spacing: { after: PARA_AFTER }, line: 240 },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: H1_SIZE, bold: true, font: REPORT_FONT, color: NUTANIX_BLUE }, paragraph: { spacing: { before: 360, after: PARA_AFTER }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: H2_SIZE, bold: true, font: REPORT_FONT, color: HEADER_GREY }, paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: PAGE_MARGIN, right: PAGE_MARGIN, bottom: PAGE_MARGIN, left: PAGE_MARGIN, header: HEADER_FOOTER_DISTANCE, footer: HEADER_FOOTER_DISTANCE } } },
    headers: { default: { options: { children: [new Paragraph({
      children: [
        new TextRun({ text: `Nutanix Health Check – ${C.cluster_name} – ${D.customer}`, font: REPORT_FONT, size: TABLE_FONT_SIZE, color: HEADER_GREY }),
        new TextRun({ text: "\t" }), new TextRun({ text: D.date, font: REPORT_FONT, size: TABLE_FONT_SIZE, color: HEADER_GREY }),
      ],
      tabStops: [{ type: "right", position: 10800 }],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NUTANIX_BLUE, space: 4 } },
    })] } } },
    footers: { default: { options: { children: [new Paragraph({
      children: [
        new TextRun({ text: "Confidential", font: REPORT_FONT, size: TABLE_FONT_SIZE, italics: true, color: "888888" }),
        new TextRun({ text: "\t" }),
        new TextRun({ text: "Page ", font: REPORT_FONT, size: TABLE_FONT_SIZE, color: HEADER_GREY }),
        new TextRun({ children: [PageNumber.CURRENT], font: REPORT_FONT, size: 18 }),
        new TextRun({ text: " of ", font: REPORT_FONT, size: TABLE_FONT_SIZE, color: HEADER_GREY }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: REPORT_FONT, size: 18 }),
      ],
      tabStops: [{ type: "right", position: 10800 }],
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: NUTANIX_BLUE, space: 4 } },
    })] } } },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[3], buf); console.log("Report written to", process.argv[3]); });
"""


# Persistent directory next to this script where docx is installed locally.
# Using a local node_modules avoids the Windows global-npm PATH lookup issue
# where Node cannot find globally installed modules when running scripts from
# an arbitrary temp directory.
_SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
_NODE_MODULES = os.path.join(_SCRIPT_DIR, "node_modules")
_REPORT_JS    = os.path.join(_SCRIPT_DIR, "_report_builder.js")


def _ensure_docx_installed() -> bool:
    """
    Install the docx package locally (next to this script) if not already
    present.  Returns True on success, False on failure.
    """
    docx_marker = os.path.join(_NODE_MODULES, "docx", "package.json")
    if os.path.isfile(docx_marker):
        return True  # already installed

    print("    Installing docx npm package locally (one-time setup) ...")
    npm_cmd = "npm.cmd" if sys.platform == "win32" else "npm"
    result  = subprocess.run(
        [npm_cmd, "install", "docx"],
        cwd=_SCRIPT_DIR,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"    [ERROR] npm install failed:\n{result.stderr.strip()}")
        return False
    print("    docx installed successfully.")
    return True


def _write_report_js() -> None:
    """Write the report builder JS next to this script (overwrite each run)."""
    with open(_REPORT_JS, "w", encoding="utf-8") as f:
        f.write(REPORT_JS)


def _make_chart_png(history: list, title: str, ylabel: str,
                     suffix: str, color: str = "#1f77b4") -> Optional[str]:
    """Shared chart generator for CPU and Memory history."""
    if not isinstance(history, list) or len(history) < 2:
        return None
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
    except Exception as exc:
        print(f"    WARNING: Chart skipped ({suffix}). Install matplotlib. Details: {exc}")
        return None

    xs, ys = [], []
    for point in history:
        try:
            ts = str(point.get("timestamp", "")).replace("Z", "+00:00")
            xs.append(datetime.fromisoformat(ts))
            ys.append(float(point.get("value")))
        except Exception:
            continue

    if len(xs) < 2:
        return None

    fd, chart_path = tempfile.mkstemp(suffix=f"_{suffix}.png")
    os.close(fd)

    avg  = sum(ys) / len(ys)
    peak = max(ys)

    fig, ax = plt.subplots(figsize=(10.0, 4.7), dpi=150)
    ax.plot(xs, ys, linewidth=1.8, color=color)
    ax.axhline(avg, linestyle="--", linewidth=1.0, color=color, alpha=0.6)
    ax.set_title(title, fontsize=13, weight="bold")
    ax.set_ylabel(ylabel)
    ax.set_xlabel("Time")
    ax.set_ylim(bottom=0, top=max(10, min(100, peak * 1.25)))
    ax.grid(True, linewidth=0.5, alpha=0.4)
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=4, maxticks=8))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
    fig.autofmt_xdate(rotation=0)
    ax.text(0.99, 0.92, f"Avg: {avg:.2f}%   Peak: {peak:.2f}%",
            transform=ax.transAxes, ha="right", va="top", fontsize=10)
    fig.tight_layout()
    fig.savefig(chart_path, bbox_inches="tight")
    plt.close(fig)
    return chart_path


def _generate_cpu_chart_png(findings: dict) -> Optional[str]:
    """Create a temporary PNG chart for the CPU history using matplotlib."""
    return _make_chart_png(
        findings.get("cpu", {}).get("cpu_history") or [],
        title="CPU Usage Over the Last 7 Days",
        ylabel="CPU Usage (%)",
        suffix="cpu_7d",
        color="#1f77b4",
    )


def _generate_memory_chart_png(findings: dict) -> Optional[str]:
    """Create a temporary PNG chart for the Memory history using matplotlib."""
    return _make_chart_png(
        findings.get("memory", {}).get("mem_history") or [],
        title="Memory Usage Over the Last 7 Days",
        ylabel="Memory Usage (%)",
        suffix="mem_7d",
        color="#d62728",
    )


def _generate_storage_chart_png(findings: dict) -> Optional[str]:
    """Create a temporary PNG chart for Storage usage history using matplotlib."""
    return _make_chart_png(
        findings.get("storage", {}).get("storage_history") or [],
        title="Storage Usage Over the Last 7 Days",
        ylabel="Storage Used (%)",
        suffix="storage_7d",
        color="#2ca02c",
    )


def generate_report(findings: dict, output_path: str) -> None:
    # The Node builder runs from the script directory so local node_modules can
    # be resolved. Normalize the requested report path first so relative
    # --output-dir values still write to the caller's intended directory.
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Ensure docx is available locally
    if not _ensure_docx_installed():
        print("    [ERROR] Cannot generate report — docx package unavailable.")
        return

    # Build CPU and Memory chart images before handing data to the Node report builder.
    temp_chart_path = _generate_cpu_chart_png(findings)
    if temp_chart_path:
        findings.setdefault("cpu", {})["cpu_chart_path"] = temp_chart_path

    temp_mem_chart_path = _generate_memory_chart_png(findings)
    if temp_mem_chart_path:
        findings.setdefault("memory", {})["mem_chart_path"] = temp_mem_chart_path

    temp_storage_chart_path = _generate_storage_chart_png(findings)
    if temp_storage_chart_path:
        findings.setdefault("storage", {})["storage_chart_path"] = temp_storage_chart_path

    # Write the JS builder next to this script so node_modules is in scope
    _write_report_js()

    # Write findings to a temp JSON file
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as tf:
        json.dump(findings, tf, indent=2)
        data_file = tf.name

    try:
        env = os.environ.copy()
        # Also set NODE_PATH to the local node_modules as a belt-and-suspenders
        # fallback in case any edge-case lookup skips cwd resolution.
        env["NODE_PATH"] = _NODE_MODULES

        result = subprocess.run(
            ["node", _REPORT_JS, data_file, output_path],
            cwd=_SCRIPT_DIR,          # <-- run from script dir so node_modules is found
            env=env,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            print(f"    [ERROR] Node.js: {result.stderr.strip()}")
        else:
            print(f"    {result.stdout.strip()}")
    finally:
        try:
            os.unlink(data_file)
        except OSError:
            pass
        for _tmp in [temp_chart_path, temp_mem_chart_path, temp_storage_chart_path]:
            if _tmp:
                try:
                    os.unlink(_tmp)
                except OSError:
                    pass


def safe_filename(name: str) -> str:
    """Strip characters unsafe for filenames."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


class _TeeStream:
    """Write console output to both the terminal and the execution log."""

    def __init__(self, terminal, log_file):
        self.terminal = terminal
        self.log_file = log_file

    def write(self, text):
        self.terminal.write(text)
        self.log_file.write(text)
        return len(text)

    def flush(self):
        self.terminal.flush()
        self.log_file.flush()

    def isatty(self):
        return self.terminal.isatty()

    @property
    def encoding(self):
        return getattr(self.terminal, "encoding", "utf-8")


def start_execution_log(output_dir: str, timestamp: Optional[str] = None) -> str:
    """Capture console output in an output-dir/logs timestamped log file."""
    output_dir = os.path.abspath(output_dir or ".")
    log_dir = os.path.join(output_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)

    timestamp = timestamp or datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_name = f"Nutanix_Health_Check_{timestamp}"
    log_path = os.path.join(log_dir, f"{base_name}.log")
    sequence = 1
    while os.path.exists(log_path):
        log_path = os.path.join(log_dir, f"{base_name}_{sequence:02d}.log")
        sequence += 1

    log_file = open(log_path, "x", encoding="utf-8", buffering=1)
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    sys.stdout = _TeeStream(original_stdout, log_file)
    sys.stderr = _TeeStream(original_stderr, log_file)

    def close_execution_log():
        finished = datetime.now().astimezone().isoformat(timespec="seconds")
        try:
            log_file.write(f"\nExecution finished: {finished}\n")
            log_file.flush()
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            log_file.close()

    atexit.register(close_execution_log)
    started = datetime.now().astimezone().isoformat(timespec="seconds")
    print(f"Execution started: {started}")
    print(f"Execution log: {log_path}")
    return log_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Nutanix Health Check – Prism Central Edition (APIv4)"
    )
    p.add_argument("--host",      help="Prism Central IP or FQDN (skips interactive prompt)")
    p.add_argument("--port",      type=int, default=9440)
    p.add_argument("--user",      help="Prism username (skips interactive prompt)")
    p.add_argument("--password",  help="Prism password (skips interactive prompt)")
    p.add_argument("--customer",  default="", help="Customer name for the report")
    p.add_argument("--output-dir", default=".", help="Directory for output files (default: current dir)")
    p.add_argument("--data-only", action="store_true", help="Save raw JSON only; skip report generation")
    p.add_argument("--from-json", help="Re-generate report from saved raw JSON (no cluster connection)")
    p.add_argument("--os-compat-csv", default="", help="Path to OS Compatibility Matrix CSV (required for report generation if not in script/current folder)")
    p.add_argument("--aos-eol-csv", default="", help="Path to NOS/AOS EOL information CSV (required for report generation if not in script/current folder)")
    return p.parse_args()


def preflight_required_support_files(args: argparse.Namespace) -> None:
    """Validate required CSV support files before prompting for Prism Central details."""
    print()
    print("------------------------------------------------------------")
    print("Nutanix Health Check - Preflight Validation")
    print("------------------------------------------------------------")
    print()
    print("Checking output directories...")
    print()

    log_dir = os.path.abspath(os.path.join(args.output_dir or ".", "logs"))
    if os.path.isdir(log_dir) and os.access(log_dir, os.W_OK):
        print(f"  [OK] Logs directory: {log_dir}")
    else:
        print(f"  [ERROR] Logs directory is missing or not writable: {log_dir}")
        print()
        sys.exit(1)

    if getattr(args, "data_only", False):
        print()
        print("Support-file validation skipped in data-only mode.")
        return

    print()
    print("Checking required support files...")
    print()

    checks = [
        ("OS Compatibility Matrix CSV", OS_COMPAT_CSV_FILENAMES, "os_compat_csv"),
        ("AOS/NOS EOL information CSV", AOS_EOL_CSV_FILENAMES, "aos_eol_csv"),
    ]

    missing = []
    for label, filenames, attr in checks:
        explicit_path = getattr(args, attr, "")
        found = _find_optional_file(filenames, explicit_path)
        if found:
            setattr(args, attr, found)
            print(f"  [OK] {os.path.basename(found)}")
        else:
            missing.append((label, filenames))
            print(f"  [MISSING] {filenames[0]}")

    if missing:
        print()
        print("ERROR: Required support files are missing.")
        print()
        print("Place the missing CSV files in the same folder as this script and run the health check again.")
        print()
        print("Expected files:")
        for _, filenames in missing:
            print(f"  - {filenames[0]}")
        print()
        sys.exit(1)

    print()
    print("All required support files found.")
    print("Proceeding to Prism Central connection...")


def main() -> None:
    args = parse_args()
    run_timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    start_execution_log(args.output_dir, run_timestamp)

    # Validate required CSV files before prompting for Prism Central details.
    preflight_required_support_files(args)

    # ── offline mode: re-generate from saved JSON ────────────────────────
    if args.from_json:
        print(f"\nLoading data from {args.from_json} ...")
        with open(args.from_json, encoding="utf-8") as f:
            raw = json.load(f)
        customer = args.customer or "CUSTOMER_NAME"
        cluster_name = raw.get("cluster_info", {}).get("data", {}).get("name", "Cluster")
        if isinstance(cluster_name, list):
            cluster_name = cluster_name[0].get("name", "Cluster") if cluster_name else "Cluster"
        findings = HealthAnalyser(raw, customer, cluster_name, args.os_compat_csv, args.aos_eol_csv).analyse_all()
        out_path = os.path.join(
            args.output_dir,
            f"{safe_filename(cluster_name)}_Health_Check_{run_timestamp}.docx",
        )
        print("Generating report ...")
        generate_report(findings, out_path)
        print(f"\nDone! Report: {out_path}")
        return

    # ── interactive or argument-based connection setup ───────────────────
    if args.host and args.user:
        host     = args.host
        port     = args.port
        username = args.user
        password = args.password or getpass.getpass("  Password: ")
        customer = args.customer or "CUSTOMER_NAME"
        print(BANNER)
    else:
        host, port, username, password, customer = prompt_connection()

    os.makedirs(args.output_dir, exist_ok=True)

    # ── connect & verify ─────────────────────────────────────────────────
    print(f"  Connecting to Prism Central at {host}:{port} ...")
    client = PrismCentralClient(host, username, password, port)

    if not client.test_connection():
        print("\n  [ERROR] Cannot reach Prism Central. Check host, port, and credentials.")
        sys.exit(1)
    print("  Connection successful.\n")

    # ── discover clusters ────────────────────────────────────────────────
    print("  Discovering registered clusters ...")
    try:
        clusters = client.list_clusters()
    except Exception as exc:
        print(f"\n  [ERROR] Failed to list clusters: {exc}")
        sys.exit(1)

    if not clusters:
        print("  [WARNING] No AOS clusters found registered to this Prism Central.")
        sys.exit(0)

    print(f"  Found {len(clusters)} cluster(s):\n")
    for i, c in enumerate(clusters, 1):
        name = c.get("name", "Unknown")
        uuid = c.get("extId", "N/A")
        print(f"    [{i}] {name}  ({uuid})")

    print()

    # ── process each cluster ─────────────────────────────────────────────
    results = []
    for idx, cluster_meta in enumerate(clusters, 1):
        cluster_name = cluster_meta.get("name", f"Cluster-{idx}")
        cluster_uuid = cluster_meta.get("extId", "")

        print(f"  [{idx}/{len(clusters)}] Processing cluster: {cluster_name}")
        print(f"          UUID: {cluster_uuid}")

        # Collect
        collector = ClusterDataCollector(client, cluster_uuid, cluster_name)
        raw       = collector.collect_all()
        # Preserve the PC cluster catalog so cross-cluster protection-policy
        # references can be rendered with names instead of UUIDs.
        raw["cluster_catalog"] = clusters

        # Save raw JSON
        safe_name    = safe_filename(cluster_name)
        raw_json_path = os.path.join(
            args.output_dir,
            f"{safe_name}_raw_{run_timestamp}.json",
        )
        with open(raw_json_path, "w", encoding="utf-8") as f:
            json.dump(raw, f, indent=2)
        print(f"      Raw data saved: {raw_json_path}")

        if getattr(args, "debug_raw", False):
            debug_raw_summary(cluster_name, raw)

        if args.data_only:
            results.append({"cluster": cluster_name, "status": "JSON saved", "file": raw_json_path})
            print()
            continue

        # Analyse
        findings  = HealthAnalyser(raw, customer, cluster_name, args.os_compat_csv, args.aos_eol_csv).analyse_all()
        out_path = os.path.join(
            args.output_dir,
            f"{safe_name}_Health_Check_{run_timestamp}.docx",
        )

        # Generate report
        print(f"      Generating report ...")
        generate_report(findings, out_path)
        results.append({"cluster": cluster_name, "status": "Report generated", "file": out_path})
        print()

    # ── summary ──────────────────────────────────────────────────────────
    print("=" * 65)
    print("  Health Check Complete")
    print("=" * 65)
    for r in results:
        print(f"  {r['cluster']:<30}  {r['status']}")
        print(f"    -> {r['file']}")
    print()


if __name__ == "__main__":
    main()
