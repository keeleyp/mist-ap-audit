# Mist AP Audit Report

A Python script that generates a comprehensive Excel report of all Juniper Mist access points across an organisation, categorised by status and health.

## What It Does

The script connects to the Juniper Mist API, retrieves all AP statistics, site information, and recent device events, then produces a multi-sheet Excel workbook with the APs grouped by operational state.

Before running, it displays a summary of the organisation (name, AP count, site count) and an estimate of the API calls required, then asks for confirmation to proceed.

### Output Sheets

#### Sheet 1 — Online APs (green header)

All APs currently connected and operational.

| Column | Description |
|---|---|
| Name | AP hostname |
| Site ID | Mist site UUID |
| Site Name | Human-readable site name |
| MAC | AP MAC address |
| Serial | Hardware serial number |
| Model | AP model (e.g. AP32, AP43) |
| HW Rev | Hardware revision |
| Device Profile | Assigned Mist device profile |
| Status | Connection status (`connected`) |
| Uptime (days) | How long the AP has been running since last reboot |

#### Sheet 2 — Offline < N Days (orange header)

APs that have recently gone offline (within the configurable threshold, default 4 days). These are the APs most likely to need immediate attention — they may indicate a recent outage, power issue, or network problem.

Includes all columns from Sheet 1 plus:

| Column | Description |
|---|---|
| EPOC Last Seen | Unix timestamp of last contact |
| Date Last Seen | Human-readable UTC date/time |
| Days Since Last Seen | Decimal number of days since last contact |
| AP_RESTARTED (24h) | Number of restart events in the last 24 hours |
| AP_DISCONNECTED (24h) | Number of disconnect events in the last 24 hours |
| AP_CONNECTED (24h) | Number of connect events in the last 24 hours |

The event columns help identify whether an AP is flapping (repeatedly connecting and disconnecting) or experienced a single outage.

#### Sheet 3 — Offline > N Days (red header)

APs that have been offline for longer than the threshold. These are likely decommissioned, in storage, or at sites with long-term issues.

Includes the same columns as Sheet 2 but without the 24-hour event counts (since there will be no recent events for APs that have been offline for this long).

#### Sheet 4 — Uptime < N Hours (blue header)

APs that are currently online but have been running for less than the configurable threshold (default 24 hours). These APs have recently rebooted, which may indicate power cycling, firmware upgrades, or instability.

Includes all columns from Sheet 1 plus:

| Column | Description |
|---|---|
| Uptime (seconds) | Raw uptime in seconds |
| Uptime (hours) | Uptime converted to hours |
| AP_RESTARTED (24h) | Number of restart events in the last 24 hours |
| AP_DISCONNECTED (24h) | Number of disconnect events in the last 24 hours |
| AP_CONNECTED (24h) | Number of connect events in the last 24 hours |

An AP with multiple restart events in 24 hours is likely experiencing a problem (power instability, crash loop, etc.).

## Output File

The Excel file is saved to the configured output directory with the naming format:

```
Mist_AP_Report_<OrgName>_<YYYY-MM-DD_HHMMSS>.xlsx
```

## Prerequisites

- Python 3.6+
- A Juniper Mist API token with read access to the organisation

### Python Dependencies

```bash
pip install requests openpyxl
```

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/keeleyp/mist-ap-audit.git
   cd mist-ap-audit
   ```

2. Copy the example config and edit it with your details:
   ```bash
   cp mist_ap_report.ini.example mist_ap_report.ini
   ```

3. Edit `mist_ap_report.ini`:
   ```ini
   [mist]
   api_base = https://api.eu.mist.com/api/v1
   org_id = your-org-uuid-here
   api_token = your-api-token-here

   [thresholds]
   offline_days = 4
   low_uptime_hours = 24

   [output]
   directory = ~
   ```

   **Note:** Use `https://api.mist.com/api/v1` for US-hosted organisations.

## Usage

```bash
python3 mist_ap_report.py
```

The script will:

1. Fetch organisation info and display a pre-flight summary
2. Ask for confirmation before proceeding
3. Fetch all AP statistics with live progress
4. Resolve site names
5. Fetch the last 24 hours of AP events
6. Generate the Excel report
7. Display a summary of results and API usage

### Example Output

```
Fetching organisation info...

=======================================================
  Mist AP Report
=======================================================
  Organisation:     Contoso Ltd
  Org ID:           79017ec2-8f9c-4dbf-a270-18de1cbe9dab
  Total APs:        ~31461
  Total Sites:      ~3242
=======================================================
  This report will:
    1. Fetch all AP stats  (~32 API calls)
    2. Fetch site names    (~4 API calls)
    3. Fetch 24h events    (~50 API calls)
  Estimated total API calls: ~88
=======================================================

  Proceed? (y/n): y
```

## Configuration

| Section | Key | Description | Default |
|---|---|---|---|
| `mist` | `api_base` | Mist API base URL | `https://api.eu.mist.com/api/v1` |
| `mist` | `org_id` | Organisation UUID | — |
| `mist` | `api_token` | API authentication token | — |
| `thresholds` | `offline_days` | Boundary between "recently offline" and "long-term offline" | `4` |
| `thresholds` | `low_uptime_hours` | Uptime threshold for Sheet 4 | `24` |
| `output` | `directory` | Where to save the Excel file (`~` expands to home) | `~` |

## API Calls

The script is designed to minimise API usage:

- AP stats are fetched in pages of 1000
- Site names are fetched in bulk via the org-level sites endpoint (not per-site)
- Events use cursor-based pagination via the `next` field
- A typical run for ~30,000 APs uses approximately 90 API calls

## Security

The `mist_ap_report.ini` file containing your API token is excluded from version control via `.gitignore`. Never commit this file to a public repository.

## License

MIT
