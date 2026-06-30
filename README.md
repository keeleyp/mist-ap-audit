# Mist AP Tools

A collection of Python scripts for querying the Juniper Mist API and generating Excel reports on access point health and inventory.

| Script | Description |
|---|---|
| `mist_ap_report.py` | Multi-sheet health audit — APs grouped by status, offline duration, and recent reboots |
| `mist_ap_details.py` | Single-sheet inventory detail — per-AP port stats, LLDP neighbours, PoE draw, IP info |

Both scripts share the same `mist_ap_report.ini` config file.

---

## mist_ap_report.py — AP Health Audit

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

---

## mist_ap_details.py — AP Inventory Detail

Queries the Mist API **site by site** and produces a single-sheet Excel workbook with deep per-AP detail — port stats, LLDP neighbour information, PoE draw, IP configuration, and more. Useful for network audits, switch port mapping, and PoE capacity planning.

### What It Does

1. Fetches org info and site list
2. Calls `/sites/{site_id}/stats/devices?type=ap&limit=1000` once per site
3. Checks the API rate limit before starting and every 50 sites during the run — pausing automatically if calls run low
4. Produces a single filtered, frozen Excel sheet with one row per AP

### Output Sheet — AP Details (dark indigo header)

#### Identity

| Column | Description |
|---|---|
| Site Name | Human-readable site name |
| AP Name | AP hostname |
| MAC | AP MAC address |
| Status | `connected` or `disconnected` |
| Power Constrained | Whether the AP is operating below full power due to PoE budget |

#### Network

| Column | Description |
|---|---|
| IP | AP management IP address |
| Gateway | Default gateway |
| Netmask | Subnet mask |
| DNS Servers | Configured DNS servers (comma-separated) |
| DHCP Server | IP of the DHCP server that issued the lease |
| External IP | Public/NAT IP seen from the Mist cloud |
| Mount | Physical orientation (`facedown`, `faceup`, `wall`, etc.) |

#### eth0 Port Statistics

| Column | Description |
|---|---|
| eth0 Up | Whether the uplink port is active |
| eth0 Speed (Mbps) | Negotiated port speed |
| eth0 Full Duplex | Duplex mode |
| eth0 TX / RX Bytes | Cumulative bytes transmitted / received |
| eth0 TX / RX Pkts | Cumulative packets transmitted / received |
| eth0 RX Errors | Receive error count |
| eth0 RX / TX Peak (bps) | Peak observed throughput |

#### LLDP Neighbour (switch port the AP is connected to)

| Column | Description |
|---|---|
| Switch Name | LLDP system name of the upstream switch |
| Switch Description | Full system description including software version |
| Switch Mgmt Addr | Switch management MAC/IP |
| Switch Port Desc | Description of the switch port |
| Switch Port ID | Interface identifier (e.g. `ge-0/0/42`) |
| Switch Chassis ID | Switch chassis MAC |
| LLDP MED Supported | Whether the switch port supports LLDP-MED |
| PoE Req Count | Number of PoE power negotiation attempts |
| PoE Allocated (mW) | Power budget allocated by the switch |
| PoE Requested (mW) | Power requested by the AP |
| PoE Draw (mW) | Actual power being drawn |

#### Other

| Column | Description |
|---|---|
| Inactive Wired VLANs | VLANs configured but with no active clients |
| ESL Stat | Electronic shelf label state (JSON, blank if none) |

### Output File

```
Mist_AP_Details_<OrgName>_<YYYY-MM-DD_HHMMSS>.xlsx
```

### Rate-Limit Handling

The Mist API allows **5,000 calls per hour**. This script makes one call per site, so large organisations (e.g. 3,000+ sites) may approach that limit.

**Pre-flight:** before starting, the script fetches current usage from `/self/usage` and compares calls remaining against the estimated run cost. If there are insufficient calls:

```
  *** INSUFFICIENT API CALLS ***
  Calls still needed this run: ~3242
  Calls available (with 10 headroom): 312
  Shortfall: ~2930 calls
  Window resets in ~47m 12s

  Options:
    w — wait for the window to reset, then start automatically
    q — quit
```

**Mid-run:** the check repeats every 50 sites. If the window rolls over during a long run, the script pauses, shows a live countdown, then continues automatically.

### Usage

```bash
python3 mist_ap_details.py
```

### Example Pre-flight Output

```
============================================================
  Mist AP Details Report
============================================================
  Organisation:        Contoso Ltd
  Org ID:              79017ec2-8f9c-4dbf-a270-18de1cbe9dab
  Total APs:           ~31461
  Total Sites:         ~3242
============================================================
  This report will:
    1. Fetch all sites         (~4 API calls)
    2. Fetch AP stats per site (~3242 API calls, one per site)
  Estimated total API calls:   ~3249
============================================================
  API rate limit:      5000 calls/hour
  Used this hour:      522
  Remaining:           4478  (window resets in ~58m 42s)
  Sufficient calls available.  OK to proceed.
============================================================

  Proceed? (y/n):
```

### Configuration

Uses the same `mist_ap_report.ini` as `mist_ap_report.py`. The `[thresholds]` section is not used by this script.

| Section | Key | Description |
|---|---|---|
| `mist` | `api_base` | Mist API base URL |
| `mist` | `org_id` | Organisation UUID |
| `mist` | `api_token` | API authentication token |
| `output` | `directory` | Where to save the Excel file |

### API Calls

- 1 call per site (no pagination needed — sites have ≤ 1,000 APs)
- 3 preflight calls (org info, org stats, usage check)
- A mid-run usage re-check every 50 sites (counts against the limit)
- Typical run for 3,000 sites ≈ 3,060 API calls

---

## Shared Setup

### Prerequisites

- Python 3.6+
- A Juniper Mist API token with read access to the organisation

```bash
pip install requests openpyxl
```

### Config File

```bash
cp mist_ap_report.ini.example mist_ap_report.ini
```

Edit `mist_ap_report.ini`:

```ini
[mist]
api_base = https://api.eu.mist.com/api/v1
org_id = YOUR_ORG_ID_HERE
api_token = YOUR_API_TOKEN_HERE

[thresholds]
offline_days = 4
low_uptime_hours = 24

[output]
directory = ~
```

Use `https://api.mist.com/api/v1` for US-hosted organisations.

## Security

`mist_ap_report.ini` is excluded from version control via `.gitignore`. Never commit it to a public repository.

## License

MIT
