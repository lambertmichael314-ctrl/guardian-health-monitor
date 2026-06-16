**# Guardian Health Monitor**



> Real-time Windows system health diagnostics with a professional cyber dashboard. Scans drivers, storage, memory, CPU, battery, event logs, and Windows Update status — then serves the results through a clean REST API and a bad-ass web interface.



!\[Dashboard Preview](https://img.shields.io/badge/status-stable-green)

!\[Python](https://img.shields.io/badge/python-3.10+-blue)

!\[Flask](https://img.shields.io/badge/flask-3.0+-lightgrey)

!\[License](https://img.shields.io/badge/license-MIT-green)



\---



**## Features**



\- \*\*8 diagnostic checks\*\* — Driver Health, Storage, Memory, CPU \& Uptime, Battery, Windows Update, Event Logs, System Info

\- \*\*REST API backend\*\* — Flask server exposes `/api/scan` and `/api/health` endpoints

\- \*\*Professional cyber dashboard\*\* — Dark theme, animated score gauge, color-coded cards, scanline overlay

\- \*\*Live terminal feed\*\* — See every step of the scan in real-time

\- \*\*No dependencies on external services\*\* — All checks run locally via WMI and the Windows Update COM API



**## Quick Start**



\### 1. Install dependencies



```bash

pip install flask pywin32 wmi

