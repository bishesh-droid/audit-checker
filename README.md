# Audit Checker

Cross-checks assets listed in a Google Sheet against your local drives and produces a colour-coded Excel availability report.

---

## Table of Contents

- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output Report](#output-report)
- [What NOT to Commit](#what-not-to-commit)
- [Dependencies](#dependencies)

---

## Project Structure

```
audit-checker/
├── audit_checker.py         # Main audit script
├── config.example.json      # Config template — copy to config.json
├── requirements.txt         # Python dependencies
└── README.md

# Created at runtime — gitignored:
├── config.json              # Your personal config
├── gsheet_cache/            # Cached Google Sheet data
├── availability_report.xlsx # Audit report output
├── audit_checker.log        # Runtime log
└── .drive_index_cache.pkl   # Cached local drive scan index
```

---

## Requirements

- Python 3.10+
- One or more local drives to scan

---

## Installation

```bash
# 1. Clone or download the repo
cd audit-checker

# 2. Install Python dependencies
pip install -r requirements.txt
```

---

## Configuration

```bash
# Copy the example config
cp config.example.json config.json
```

Open `config.json` and fill in your details:

```json
{
  "gsheet_url": "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit",

  "drives": [
    "/path/to/drive1",
    "/path/to/drive2"
  ],

  "output": "./availability_report.xlsx",

  "google_drive": {
    "enabled": true
  }
}
```

| Key | Description |
|---|---|
| `gsheet_url` | Full URL of the Google Sheet |
| `drives` | Mount paths of local drives to scan |
| `output` | Where to save the Excel report |
| `google_drive.enabled` | `true` to verify each Drive link is live |
| `scanning.fuzzy_threshold` | Name-match sensitivity 0–100 (default: 75) |
| `scanning.cache_max_age_hours` | How long the drive scan cache is valid (default: 24h) |
| `scanning.gsheet_cache_hours` | How long the sheet cache is valid (default: 1h) |

---

## Usage

```bash
# Standard audit — scan drives + check Drive links + save report
python3 audit_checker.py

# Fresh audit — ignore all caches (use after the sheet has been updated)
python3 audit_checker.py --no_cache

# Download assets that are on Drive but missing locally
python3 audit_checker.py --download

# Download ALL Drive assets regardless of local presence (full refresh)
python3 audit_checker.py --download_all

# Fresh audit + download missing
python3 audit_checker.py --no_cache --download

# Save report to a custom path
python3 audit_checker.py --output /path/to/report.xlsx

# Scan specific drives for this run only
python3 audit_checker.py --drives "/path/to/drive1" "/path/to/drive2"

# Use a different Google Sheet for this run
python3 audit_checker.py --gsheet_url "https://docs.google.com/spreadsheets/d/NEW_ID/edit"

# Relax name matching (catches folders with slightly different names)
python3 audit_checker.py --fuzzy_threshold 60

# Strict name matching (reduces false positives)
python3 audit_checker.py --fuzzy_threshold 90

# Full debug log output
python3 audit_checker.py --log_level DEBUG

# Use a different config file
python3 audit_checker.py --config /path/to/config.json
```

---

## Output Report

`availability_report.xlsx` has one row per entry with columns for each asset type:

| Column | Description |
|---|---|
| `<Asset>_Local` | `Yes` / `Yes (Downloaded)` / `No` |
| `<Asset>_Local_Path` | Full path to the matched folder on disk |
| `<Asset>_Drive` | `Available` / `Missing` / `Broken Link` / `No Link` |

**Row colour coding:**

| Colour | Meaning |
|---|---|
| Green | All assets found locally and all Drive links are live |
| Yellow | Some assets found or some Drive links live |
| Red | Nothing found locally and no Drive links accessible |

---

## What NOT to Commit

These files are in `.gitignore` and must never be pushed:

| File | Why |
|---|---|
| `config.json` | Contains your personal drive paths and sheet URL |
| `credentials.json` | Google OAuth credentials — treat like a password |
| `mycreds.txt` | Cached OAuth token |
| `gsheet_cache/` | Downloaded sheet data |
| `availability_report.xlsx` | Generated output |
| `.drive_index_cache.pkl` | Local drive scan cache |

---

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | DataFrame handling and Excel writing |
| `openpyxl` | Read Excel hyperlinks + write styled reports |
| `rapidfuzz` | Fuzzy name matching for folder lookup |
| `tqdm` | Progress bars |
| `pydrive2` | Optional authenticated Drive API access |

```bash
pip install -r requirements.txt
```

---

## License

MIT
