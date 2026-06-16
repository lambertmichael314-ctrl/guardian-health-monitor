#!/usr/bin/env python3
"""
Guardian Health Monitor — Backend API Server
Serves system health data via REST API to the dashboard frontend.
GitHub: https://github.com/YOUR_USER/guardian-health-monitor
"""

from __future__ import annotations

import os
import sys
import json
import datetime
import time
import logging
import logging.handlers
import tempfile
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Any

# Suppress COM noise from wmi/pywin32
_SAVED_STDERR = sys.stderr
sys.stderr = open(os.devnull, 'w')

from flask import Flask, jsonify, render_template, send_from_directory

app = Flask(__name__, static_folder='.', static_url_path='')

# ---- Audit Engine (same core as pc_health_audit.py) ----

LOG_DIR = Path(tempfile.gettempdir()) / "GuardianHealthLogs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("Guardian")
log.setLevel(logging.DEBUG)
log.handlers.clear()
handler = logging.handlers.RotatingFileHandler(
    LOG_DIR / f"guardian_{datetime.datetime.now():%Y%m%d_%H%M%S}.log",
    maxBytes=2 * 1024 * 1024, backupCount=5
)
handler.setLevel(logging.DEBUG)
handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)-8s %(message)s"))
log.addHandler(handler)

DISK_FREE_PCT_WARN = 10
DISK_FREE_PCT_CRIT = 5
MEM_USAGE_PCT_WARN = 85
MEM_USAGE_PCT_CRIT = 95
CPU_LOAD_WARN = 85
BATTERY_WEAR_WARN = 30
EVENT_ERROR_WARN = 10


@dataclass
class HealthCheckResult:
    category: str
    status: str
    score: int
    summary: str
    details: list[str] = field(default_factory=list)
    recommendations: list[str] = field(default_factory=list)


@dataclass
class AuditReport:
    timestamp: str = ""
    computer_name: str = ""
    os_version: str = ""
    uptime_days: float = 0.0
    overall_score: int = 0
    overall_grade: str = ""
    checks: list[HealthCheckResult] = field(default_factory=list)


def _init_com():
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _uninit_com():
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _read_wmi(query: str, cls_name: str) -> list:
    """Generic WMI query helper — returns list of result objects."""
    import wmi
    c = wmi.WMI()
    results = list(getattr(c, cls_name)())
    del c
    return results


def check_driver_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    total = 0
    unsigned = 0
    has_updates = 0

    _init_com()
    try:
        drivers = _read_wmi("SELECT * FROM Win32_PnPSignedDriver", "Win32_PnPSignedDriver")
        total = len(drivers)
        for d in drivers:
            if not bool(getattr(d, "IsSigned", False)):
                unsigned += 1
                details.append(f"  Unsigned driver: {str(getattr(d, 'DeviceName', '') or 'Unknown')[:60]}")
        details.append(f"  Total drivers: {total}")
        details.append(f"  Unsigned: {unsigned}")

        try:
            import win32com.client
            session = win32com.client.Dispatch("Microsoft.Update.Session")
            searcher = session.CreateUpdateSearcher()
            result = searcher.Search("IsInstalled=0 and Type='Driver'")
            has_updates = result.Updates.Count
            if has_updates > 0:
                details.append(f"  Driver updates available: {has_updates}")
                recommendations.append("Run Windows Update to install driver updates")
        except Exception:
            details.append("  Windows Update check: unavailable")
    except ImportError:
        _uninit_com()
        return HealthCheckResult("Driver Health", "INFO", 50, "WMI not available — install: pip install wmi", details)
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("Driver Health", "INFO", 50, f"Driver check error: {type(exc).__name__}", [f"  Error: {exc}"])
    _uninit_com()

    score = 100
    if unsigned > 0:
        score -= min(unsigned * 10, 40)
    if has_updates > 0:
        score -= min(has_updates * 5, 30)
    status = "PASS" if (score >= 80 and unsigned == 0) else "WARN" if score >= 50 else "FAIL"
    summary = f"All {total} drivers signed" if unsigned == 0 else f"{total} drivers — {unsigned} unsigned"
    return HealthCheckResult("Driver Health", status, score, summary, details, recommendations)


def check_storage_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    _init_com()
    try:
        disks = _read_wmi("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3", "Win32_LogicalDisk")
        score = 100
        disk_issues = 0
        for disk in disks:
            name = str(getattr(disk, "DeviceID", "") or "?")
            size = int(getattr(disk, "Size", 0) or 0)
            free = int(getattr(disk, "FreeSpace", 0) or 0)
            if size == 0:
                continue
            total_gb = size / (1024 ** 3)
            free_gb = free / (1024 ** 3)
            free_pct = (free / size) * 100 if size > 0 else 0
            details.append(f"  {name}: {free_gb:.1f}GB free / {total_gb:.1f}GB total ({free_pct:.1f}% free)")
            if free_pct < DISK_FREE_PCT_CRIT:
                score -= 30
                disk_issues += 1
                recommendations.append(f"{name} critically low")
            elif free_pct < DISK_FREE_PCT_WARN:
                score -= 15
                disk_issues += 1
                recommendations.append(f"{name} low on space")
        _uninit_com()
        status = "PASS" if score >= 90 else "WARN" if score >= 60 else "FAIL"
        summary = "All disks healthy" if score >= 90 else f"{disk_issues} disk(s) below threshold"
        return HealthCheckResult("Storage Health", status, score, summary, details, recommendations)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("Storage Health", "INFO", 50, "WMI not available", details)
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("Storage Health", "INFO", 50, f"Error: {type(exc).__name__}", [f"  {exc}"])


def check_memory_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    _init_com()
    try:
        os_list = _read_wmi("SELECT * FROM Win32_OperatingSystem", "Win32_OperatingSystem")
        score = 100
        for os_info in os_list:
            tv = int(getattr(os_info, "TotalVisibleMemorySize", 0) or 0)
            fp = int(getattr(os_info, "FreePhysicalMemory", 0) or 0)
            if tv > 0:
                pct = ((tv - fp) / tv) * 100
                details.append(f"  Physical RAM: {(tv-fp)/1024:.0f}MB used / {tv/1024:.0f}MB total ({pct:.1f}%)")
                if pct > MEM_USAGE_PCT_CRIT:
                    score -= 30
                    recommendations.append("RAM usage critically high")
                elif pct > MEM_USAGE_PCT_WARN:
                    score -= 15
                    recommendations.append("High RAM usage")
        _uninit_com()
        status = "PASS" if score >= 90 else "WARN" if score >= 60 else "FAIL"
        return HealthCheckResult("Memory Health", status, score, "RAM usage is healthy" if score >= 90 else "Elevated memory usage", details, recommendations)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("Memory Health", "INFO", 50, "WMI not available", details)
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("Memory Health", "INFO", 50, f"Error: {type(exc).__name__}", [f"  {exc}"])


def check_cpu_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    _init_com()
    try:
        os_list = _read_wmi("SELECT * FROM Win32_OperatingSystem", "Win32_OperatingSystem")
        score = 100
        for os_info in os_list:
            lb = getattr(os_info, "LastBootUpTime", None)
            if lb:
                try:
                    bd = datetime.datetime.strptime(str(lb).split(".")[0], "%Y%m%d%H%M%S")
                    up = datetime.datetime.now() - bd
                    details.append(f"  Uptime: {up.days}d {up.seconds//3600}h")
                    if up.days > 30:
                        score -= 10
                        recommendations.append("Uptime over 30 days")
                except Exception:
                    pass
        cpus = _read_wmi("SELECT * FROM Win32_Processor", "Win32_Processor")
        if cpus:
            load = int(getattr(cpus[0], "LoadPercentage", 0) or 0)
            name = str(getattr(cpus[0], "Name", "") or "Unknown").strip()
            details.append(f"  CPU: {name}")
            details.append(f"  Load: {load}%")
            if load > CPU_LOAD_WARN:
                score -= 20
                recommendations.append(f"CPU at {load}%")
            elif load > 70:
                score -= 10
        _uninit_com()
        status = "PASS" if score >= 90 else "WARN" if score >= 60 else "FAIL"
        return HealthCheckResult("CPU & Uptime", status, score, "CPU healthy" if score >= 90 else "CPU warnings", details, recommendations)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("CPU & Uptime", "INFO", 50, "WMI not available", details)
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("CPU & Uptime", "INFO", 50, f"Error: {type(exc).__name__}", [f"  {exc}"])


def check_battery_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    _init_com()
    try:
        import wmi
        c = wmi.WMI()
        batteries = list(c.Win32_Battery())
        del c
        _uninit_com()
        if not batteries:
            return HealthCheckResult("Battery Health", "INFO", 100, "No battery (desktop system)", ["  Desktop — no battery to check"])
        return HealthCheckResult("Battery Health", "PASS", 100, "Battery healthy", details)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("Battery Health", "INFO", 50, "WMI not available", details)
    except Exception:
        _uninit_com()
        return HealthCheckResult("Battery Health", "INFO", 100, "No battery data", details)


def check_windows_update_status() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    _init_com()
    try:
        import win32com.client
        session = win32com.client.Dispatch("Microsoft.Update.Session")
        searcher = session.CreateUpdateSearcher()
        score = 100
        result = searcher.Search("IsInstalled=0")
        total = result.Updates.Count
        important = 0
        optional = 0
        driver_count = 0
        for update in result.Updates:
            try:
                if any(c.Name == "Drivers" for c in update.Categories if hasattr(c, "Name")):
                    driver_count += 1
                elif update.AutoSelectOnWebSites:
                    important += 1
                else:
                    optional += 1
            except Exception:
                optional += 1
        details.append(f"  Pending: {total} (Important: {important}, Optional: {optional}, Drivers: {driver_count})")
        if important > 0:
            score -= min(important * 15, 50)
            recommendations.append(f"{important} important updates pending")
        if optional > 0:
            score -= min(optional * 3, 20)
        _uninit_com()
        status = "PASS" if total == 0 else "FAIL" if important > 0 else "WARN"
        summary = "Fully up to date" if total == 0 else f"{important} important updates pending" if important > 0 else f"{total} updates available"
        return HealthCheckResult("Windows Update", status, score, summary, details, recommendations)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("Windows Update", "INFO", 50, "pywin32 not installed", ["  Install: pip install pywin32"])
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("Windows Update", "INFO", 50, f"Error: {type(exc).__name__}", [f"  {exc}"])


def check_event_log_health() -> HealthCheckResult:
    details: list[str] = []
    recommendations: list[str] = []
    try:
        import win32evtlog
        score = 100
        errors = 0
        for log_name in ["System", "Application"]:
            try:
                hand = win32evtlog.OpenEventLog(None, log_name)
                flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
                events = win32evtlog.ReadEventLog(hand, flags, 0)
                cutoff = datetime.datetime.now() - datetime.timedelta(hours=24)
                for event in events[:200]:
                    try:
                        tg = event.TimeGenerated
                        if isinstance(tg, datetime.datetime) and tg < cutoff:
                            continue
                        if event.EventType == 1 and log_name == "System":
                            errors += 1
                    except Exception:
                        continue
                win32evtlog.CloseEventLog(hand)
            except Exception:
                details.append(f"  Could not read {log_name} log")
        details.append(f"  Critical errors in last 24h: {errors}")
        if errors > EVENT_ERROR_WARN:
            score -= min(errors * 5, 40)
            recommendations.append("Review System event log")
        elif errors > 0:
            score -= min(errors * 2, 20)
        status = "PASS" if score >= 90 else "WARN" if score >= 60 else "FAIL"
        summary = f"Event log clean — {errors} errors" if errors == 0 else f"{errors} errors in 24h"
        return HealthCheckResult("Event Log Health", status, score, summary, details, recommendations)
    except ImportError:
        return HealthCheckResult("Event Log Health", "INFO", 50, "pywin32 not installed", ["  Install: pip install pywin32"])
    except Exception as exc:
        return HealthCheckResult("Event Log Health", "INFO", 50, f"Error: {type(exc).__name__}", [f"  {exc}"])


def check_system_info() -> HealthCheckResult:
    details: list[str] = []
    manufacturer = "Unknown"
    caption = "Windows"
    _init_com()
    try:
        import wmi
        c = wmi.WMI()
        for os_info in c.Win32_OperatingSystem():
            caption = str(getattr(os_info, "Caption", "") or "").strip()
            build = str(getattr(os_info, "BuildNumber", "") or "")
            arch = str(getattr(os_info, "OSArchitecture", "") or "")
            details.append(f"  OS: {caption} (Build {build}, {arch})")
        for cs in c.Win32_ComputerSystem():
            manufacturer = str(getattr(cs, "Manufacturer", "") or "")
            model = str(getattr(cs, "Model", "") or "")
            total_ram = int(getattr(cs, "TotalPhysicalMemory", 0) or 0)
            details.append(f"  System: {manufacturer} {model}")
            details.append(f"  RAM: {total_ram / (1024**3):.1f} GB")
        del c
        _uninit_com()
        return HealthCheckResult("System Info", "INFO", 100, f"{manufacturer} — {caption}", details)
    except ImportError:
        _uninit_com()
        return HealthCheckResult("System Info", "INFO", 100, "System info unavailable", details)
    except Exception as exc:
        _uninit_com()
        return HealthCheckResult("System Info", "INFO", 100, f"Error: {type(exc).__name__}", [f"  {exc}"])


def compute_overall_score(checks: list[HealthCheckResult]) -> tuple[int, str]:
    scored = [c for c in checks if c.status in ("PASS", "WARN", "FAIL") and c.score >= 0]
    if not scored:
        return 0, "NO DATA"
    avg = sum(c.score for c in scored) // len(scored)
    if avg >= 85:
        return avg, "EXCELLENT"
    elif avg >= 70:
        return avg, "GOOD"
    elif avg >= 50:
        return avg, "FAIR"
    elif avg >= 30:
        return avg, "POOR"
    return avg, "CRITICAL"


def run_health_scan() -> AuditReport:
    """Execute all health checks and return a complete report."""
    checks: list[HealthCheckResult] = []
    checks.append(check_system_info())
    checks.append(check_driver_health())
    checks.append(check_storage_health())
    checks.append(check_memory_health())
    checks.append(check_cpu_health())
    checks.append(check_battery_health())
    checks.append(check_windows_update_status())
    checks.append(check_event_log_health())

    computer_name = os.environ.get("COMPUTERNAME", "Unknown")
    os_version = "Windows"
    uptime_days = 0.0

    _init_com()
    try:
        import wmi
        c = wmi.WMI()
        for os_info in c.Win32_OperatingSystem():
            os_version = str(getattr(os_info, "Caption", "") or "Windows").strip()[:60]
            lb = getattr(os_info, "LastBootUpTime", None)
            if lb:
                try:
                    bd = datetime.datetime.strptime(str(lb).split(".")[0], "%Y%m%d%H%M%S")
                    uptime_days = (datetime.datetime.now() - bd).total_seconds() / 86400
                except Exception:
                    pass
        del c
    except Exception:
        pass
    _uninit_com()

    overall_score, overall_grade = compute_overall_score(checks)

    return AuditReport(
        timestamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        computer_name=computer_name,
        os_version=os_version,
        uptime_days=uptime_days,
        overall_score=overall_score,
        overall_grade=overall_grade,
        checks=checks,
    )


# ---- Flask Routes ----

@app.route('/')
def index():
    return send_from_directory('.', 'dashboard.html')


@app.route('/api/scan')
def api_scan():
    """Run a health scan and return JSON results."""
    report = run_health_scan()
    return jsonify(asdict(report))


@app.route('/api/health')
def api_health():
    """Health check for the API itself."""
    return jsonify({"status": "online", "version": "1.0.0"})


# ---- Entry Point ----

if __name__ == '__main__':
    print("=" * 55)
    print("  Guardian Health Monitor — Backend API")
    print("  Starting server at http://localhost:5000")
    print("  Press Ctrl+C to stop")
    print("=" * 55)
    log.info("Server starting on http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
