# Audit Checker + Downloader

Two-tool suite for Coursera course content management:

- **`audit_checker.py`** — cross-checks assets listed in a Google Sheet against your local drives and produces a colour-coded Excel availability report.
- **`downloader.py`** — downloads all course asset folders from Google Drive onto two external disks, balanced and resumable, with a colour-coded Excel progress report.

---

## Table of Contents

- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Tool 1: audit\_checker.py](#tool-1-audit_checkerpy)
- [Tool 2: downloader.py](#tool-2-downloaderpy)
- [Output Reports](#output-reports)
- [What NOT to Commit](#what-not-to-commit)
- [Dependencies](#dependencies)

---

## Project Structure

```
audit-checker/
├── audit_checker.py         # Audit tool — scans drives + checks Drive links
├── downloader.py            # Downloader — fetches course folders from Drive to disk
├── config.example.json      # Config template — copy to config.json and fill in
├── requirements.txt         # Python dependencies
└── README.md

# Created at runtime — gitignored:
├── config.json              # Your personal config (never commit)
├── gsheet_cache/            # Cached Google Sheet download (auto-managed)
├── availability_report.xlsx # Audit report output
├── audit_checker.log        # Audit tool runtime log
├── downloader.log           # Downloader runtime log
├── download_progress.xlsx   # Downloader progress report
├── disk_assignment.json     # Which course goes to which disk (runtime state)
├── download_results.json    # Per-course download outcomes (runtime state)
└── .drive_index_cache.pkl   # Cached local drive scan index
```

---

## Requirements

- Python 3.10+
- **rclone** configured with a Google Drive remote named `gdrive` (required by `downloader.py`)
- One or more local drives to scan / download to

---

## Installation

```bash
# 1. Clone or download the repo
cd audit-checker

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Set up rclone (required for downloader.py only)
rclone config
# → Add a new remote, name it "gdrive", type "Google Drive", follow the prompts
```

---

## Configuration

```bash
# Copy the example config (used by audit_checker.py)
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
    "enabled": true,
    "download_dest": "/path/to/drive"
  },

  "scanning": {
    "fuzzy_threshold": 75,
    "cache_max_age_hours": 24,
    "gsheet_cache_hours": 1
  }
}
```

| Key | Description |
|---|---|
| `gsheet_url` | Full URL of the Google Sheet (must be shared publicly) |
| `drives` | Mount paths of local drives to scan |
| `output` | Where to save the audit Excel report |
| `google_drive.enabled` | `true` to verify each Drive link is live |
| `google_drive.download_dest` | Default drive for `--download` downloads |
| `scanning.fuzzy_threshold` | Name-match sensitivity 0–100 (default: 75) |
| `scanning.cache_max_age_hours` | How long the drive scan cache is valid (default: 24h) |
| `scanning.gsheet_cache_hours` | How long the sheet cache is valid (default: 1h) |

> **Note:** `downloader.py` has its own hardcoded constants at the top of the file (disk paths, sheet ID, sheet tab name, column indices). Edit those directly to match your setup.

---

## Tool 1: audit\_checker.py

Reads a Google Sheet (or local Excel/CSV files), extracts hidden Drive hyperlinks from six asset columns, scans local drives using fuzzy name matching, checks whether each Drive link is publicly accessible, and writes a colour-coded `availability_report.xlsx`.

### Usage

```bash
# Standard audit — fetch sheet + scan drives + check Drive links
python3 audit_checker.py

# Fresh run — ignore all caches (use after the sheet has been updated)
python3 audit_checker.py --no_cache

# Audit then download everything missing onto a drive
python3 audit_checker.py --download --download_dest "/run/media/you/Drive Name"

# Download ALL Drive assets regardless of local presence (full sync)
python3 audit_checker.py --download_all --download_dest "/run/media/you/Drive Name"

# Fresh audit + download missing
python3 audit_checker.py --no_cache --download

# Save report to a custom path
python3 audit_checker.py --output /path/to/report.xlsx

# Scan specific drives for this run only
python3 audit_checker.py --drives "/path/to/drive1" "/path/to/drive2"

# Use a different Google Sheet for this run
python3 audit_checker.py --gsheet_url "https://docs.google.com/spreadsheets/d/SHEET_ID/edit"

# Use a local Excel folder instead of a Google Sheet
python3 audit_checker.py --excel_dir ./excel

# Relax name matching (catches folders with slightly different names)
python3 audit_checker.py --fuzzy_threshold 60

# Strict name matching (reduces false positives)
python3 audit_checker.py --fuzzy_threshold 90

# Enable authenticated Google Drive API checking (requires credentials.json)
python3 audit_checker.py --gdrive

# Full debug log output
python3 audit_checker.py --log_level DEBUG

# Use a different config file
python3 audit_checker.py --config /path/to/config.json
```

### CLI Reference (audit\_checker.py)

| Flag | Description |
|---|---|
| `--gsheet_url URL` | Google Sheets URL to fetch and audit |
| `--excel_dir DIR` | Folder with .xlsx/.csv files (fallback if no sheet URL) |
| `--config FILE` | Path to a custom config.json |
| `--drives PATH…` | One or more local drive paths to scan |
| `--output FILE` | Path for the Excel audit report |
| `--download` | Download assets that are on Drive but missing locally |
| `--download_all` | Download ALL linked Drive assets (full sync) |
| `--download_dest DIR` | Root folder where downloaded assets are saved |
| `--min_free_gb GB` | Minimum free space; skip a course if below threshold (default: 5.0) |
| `--no_cache` | Ignore all caches — re-download sheet and re-scan drives |
| `--gdrive` | Enable authenticated Drive API checking via pydrive2 |
| `--fuzzy_threshold N` | Name-match sensitivity 0–100 (default: 75) |
| `--log_level LEVEL` | Logging verbosity: DEBUG, INFO, WARNING, ERROR |

---

## Tool 2: downloader.py

Dedicated bulk downloader. Reads the master Google Sheet, assigns each completed course to one of two external disks (balanced by free space), and downloads every asset folder (Course Outline, PPTs, Written Assets, Final Videos, Raw Videos, Course Artifacts) using **rclone**.

### Key behaviours

- Each course lives entirely on **one disk** — no cross-disk splits.
- If a course folder already exists on a disk, that disk is reused.
- New courses go to the disk with more free space; if space is similar, assignments alternate to balance load.
- **Resumable** — already-downloaded folders are skipped automatically.
- Duplicate Drive folder IDs within a course are detected and skipped.
- State is persisted in `disk_assignment.json` and `download_results.json`.

### Usage

```bash
# Full run — download all 'Completed' courses with Drive links
python3 downloader.py

# Download a single course (case-insensitive partial match)
python3 downloader.py --course "Writing Practice"

# Download multiple specific courses
python3 downloader.py --course "Discrete Math" --course "Linear Algebra"

# Dry run — show what would be downloaded without doing anything
python3 downloader.py --dry-run

# Report only — regenerate the Excel without downloading
python3 downloader.py --report-only

# Force re-download the Google Sheet (ignore 1-hour cache)
python3 downloader.py --refresh-sheet
```

### CLI Reference (downloader.py)

| Flag | Description |
|---|---|
| `--course NAME` | Download only courses whose name contains NAME (repeatable) |
| `--dry-run` | Preview what would be downloaded without downloading |
| `--report-only` | Regenerate `download_progress.xlsx` without downloading |
| `--refresh-sheet` | Force re-download the Google Sheet even if cache is fresh |

### rclone setup (required)

`downloader.py` uses rclone under the hood. Run this once:

```bash
rclone config
# Add remote → name it "gdrive" → type "drive" → follow auth prompts
```

Then verify it works:

```bash
rclone lsd gdrive:
```

---

## Output Reports

### availability\_report.xlsx (audit\_checker.py)

One row per course. For each of the six asset types:

| Column | Description |
|---|---|
| `<Asset>_Local` | `Yes` / `Yes (Downloaded)` / `No` |
| `<Asset>_Local_Path` | Full path to the matched folder on disk |
| `<Asset>_Drive` | `Available` / `Missing` / `Broken Link` / `No Link` / `Link Present (not checked)` |

**Row colour coding:**

| Colour | Meaning |
|---|---|
| Green | All assets found locally and all Drive links are live |
| Yellow | Some assets found or some Drive links live |
| Red | Nothing found locally and no Drive links accessible |

### download\_progress.xlsx (downloader.py)

One row per course with per-asset download status:

| Colour | Meaning |
|---|---|
| Green | Downloaded successfully |
| Yellow | Already present — skipped |
| Red | Download failed |
| Orange | No Drive link / In Production |
| Grey | Not yet started |

A **Legend** sheet is included in the workbook.

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
| `download_progress.xlsx` | Generated output |
| `.drive_index_cache.pkl` | Local drive scan cache |
| `disk_assignment.json` | Runtime state (disk ↔ course mapping) |
| `download_results.json` | Runtime state (per-course download outcomes) |
| `*.log` | Runtime logs |

---

## Dependencies

### Python packages

| Package | Used by | Purpose |
|---|---|---|
| `pandas` | audit_checker | DataFrame handling and Excel writing |
| `openpyxl` | both | Read Excel hyperlinks + write styled reports |
| `rapidfuzz` | audit_checker | Fuzzy name matching for folder lookup |
| `tqdm` | audit_checker | Progress bars |
| `pydrive2` | audit_checker | Optional: authenticated Drive API access |
| `gdown` | audit_checker | Optional: folder downloads via `--download` |

```bash
pip install -r requirements.txt
```

### System tool

| Tool | Used by | Purpose |
|---|---|---|
| `rclone` | downloader | Authenticated Google Drive folder downloads |

Install rclone from [rclone.org](https://rclone.org/install/) or:

```bash
sudo apt install rclone        # Debian/Ubuntu/Kali
# or
curl https://rclone.org/install.sh | sudo bash
```

---

## License

MIT
