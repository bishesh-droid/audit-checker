# Coursera Audit Checker

> Automatically cross-check your Coursera course materials on Google Drive against files stored on local hard drives — and download anything that's missing.

Reads a Google Sheet that lists courses and their Google Drive folder links across six asset types, scans your connected drives for matching content, checks whether every Drive folder is still live, and produces a colour-coded Excel report. Missing assets can be downloaded directly to the correct course folder on your drive.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Command Examples](#command-examples)
- [Terminal Output](#terminal-output)
- [Output Report](#output-report)
- [Project Structure](#project-structure)
- [What NOT to Commit](#what-not-to-commit)
- [Dependencies](#dependencies)

---

## Features

- **Google Sheets as input** — paste a sharing URL, the sheet is fetched automatically (no manual downloads)
- **Six asset types per course** — Course Outline, PPTs, Written Assets, Final Videos, Raw Videos, Course Artifacts
- **Live Drive link checking** — each folder URL is verified as Available, Missing, or Broken
- **Local drive scanning** — recursively indexes connected hard drives using fuzzy name matching
- **Auto-download missing assets** — downloads entire Google Drive folders to the correct course subfolder via `gdown`
- **Smart caching** — drive index cached 24h, sheet cached 1h — repeated runs are fast
- **Colour-coded Excel report** — green / yellow / red per course row with per-asset columns
- **Fully configurable** — all settings in `config.json`, everything overridable via CLI flags

---

## Requirements

- Python 3.10+
- A Google Sheet shared as **Anyone with the link can view**
- One or more local drives mounted as filesystem paths

---

## Installation

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/audit-checker.git
cd audit-checker

# 2. Create a virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Make the script executable (Linux / macOS)
chmod +x audit_checker.py
```

---

## Configuration

### Step 1 — Copy the example config

```bash
cp config.example.json config.json
```

### Step 2 — Fill in your values

Open `config.json` and set at minimum these three things:

```json
{
  "gsheet_url": "https://docs.google.com/spreadsheets/d/1ABC123XYZ/edit",

  "drives": [
    "/run/media/duffer/One Touch A",
    "/run/media/duffer/One Touch B"
  ],

  "google_drive": {
    "enabled": true,
    "download_dest": "/run/media/duffer/One Touch A"
  }
}
```

| Key | Description |
|-----|-------------|
| `gsheet_url` | Full URL of your Google Sheet (must be set to *Anyone with link can view*) |
| `drives` | Mount paths of your connected hard drives to scan |
| `google_drive.enabled` | `true` to verify each Drive link live (recommended) |
| `google_drive.download_dest` | Drive path where missing files will be downloaded |
| `scanning.fuzzy_threshold` | Name-match sensitivity 0–100 (default: 75) |
| `scanning.cache_max_age_hours` | How long drive index cache is valid (default: 24) |
| `scanning.gsheet_cache_hours` | How long the sheet cache is valid (default: 1) |
| `google_drive.min_free_gb` | Min free disk space before skipping a course download (default: 5.0 GB) |

### Google Sheet format

The sheet must have these column headers (names are configurable in `config.json`):

| Course | Sem | Term | Status | Course Outline | PPTs | Written Assets (PQ, GQ, DP) | Final Videos | Raw Videos | Course Artifacts Link |
|--------|-----|------|--------|----------------|------|------------------------------|--------------|------------|-----------------------|
| Intro to Programming | S1 | T1 | Active | [link] | [link] | [link] | [link] | [link] | [link] |

Each asset cell should contain a **hyperlinked label** pointing to a Google Drive folder — the tool extracts the real URL automatically.

---

## Usage

```
python audit_checker.py [OPTIONS]
```

Or if made executable:

```
./audit_checker.py [OPTIONS]
```

### All Flags

```
INPUT
  --gsheet_url URL        Google Sheets URL (overrides config for this run)
  --excel_dir DIR         Folder with local .xlsx/.csv files (fallback input)
  --config FILE           Path to a custom config.json

DRIVES
  --drives PATH [PATH …]  Drive paths to scan (overrides config for this run)

OUTPUT
  --output FILE           Report save path (default: ./availability_report.xlsx)

DOWNLOAD
  --download              Download assets Available on Drive but missing locally
  --download_all          Download ALL linked assets regardless of local presence
  --download_dest DIR     Drive to save downloads to (overrides config)
  --min_free_gb GB        Min free disk space (GB) before skipping a course (default: 5.0)

CACHE
  --no_cache              Ignore all cached data — re-download sheet + rescan drives

ADVANCED
  --gdrive                Force Drive link checking on (if disabled in config)
  --fuzzy_threshold N     Name-match sensitivity 0–100 (default: 75)
  --workers N             Parallel scan workers (default: CPU count)
  --log_level LEVEL       DEBUG | INFO | WARNING | ERROR (default: INFO)
```

---

## Command Examples

### Run a standard audit

```bash
python audit_checker.py
```

Fetches the sheet (uses cache if fresh), scans drives (uses cache if fresh), checks all Drive links, saves the report.

---

### Run with a fresh sheet and drive rescan

```bash
python audit_checker.py --no_cache
```

Forces re-download of the Google Sheet and full re-scan of all drives. Use this after the spreadsheet has been updated.

---

### Download everything that is missing locally

```bash
python audit_checker.py --download
```

After auditing, downloads every asset that is **Available on Google Drive** but **not found on your local drives**. Files are saved as:

```
<download_dest>/
  <Course Name>/
    Course_Outline/
    PPTs/
    Written_Assets/
    Final_Videos/
    Raw_Videos/
    Course_Artifacts/
```

---

### Full sync — download all Drive assets regardless of local presence

```bash
python audit_checker.py --download_all
```

Downloads **all** linked Drive assets, even ones already found locally. Useful for a full refresh.

---

### Fresh run + download missing assets

```bash
python audit_checker.py --no_cache --download
```

Re-downloads sheet, re-scans drives, audits, then downloads everything missing.

---

### Override drives for a single run

```bash
python audit_checker.py --drives "/run/media/duffer/One Touch A" "/run/media/duffer/One Touch B"
```

Scans specific drives without editing `config.json`.

---

### Override the Google Sheet URL for a single run

```bash
python audit_checker.py --gsheet_url "https://docs.google.com/spreadsheets/d/1XYZ_NEW_ID/edit"
```

---

### Save the report to a custom path

```bash
python audit_checker.py --output /home/duffer/reports/march_audit.xlsx
```

---

### Download to a specific drive

```bash
python audit_checker.py --download --download_dest "/run/media/duffer/One Touch B"
```

---

### Relax fuzzy matching (catch more courses with slightly different folder names)

```bash
python audit_checker.py --fuzzy_threshold 60
```

Lower value = more permissive matching. Default is 75. Try 60–70 if many courses show "No local match".

---

### Strict fuzzy matching (reduce false positives)

```bash
python audit_checker.py --fuzzy_threshold 90
```

Higher value = stricter — only folders with nearly identical names are matched.

---

### Enable verbose debug logging

```bash
python audit_checker.py --log_level DEBUG
```

Full debug output written to `audit_checker.log`.

---

### Use a custom config file

```bash
python audit_checker.py --config /home/duffer/configs/semester2.json
```

---

## Terminal Output

This is what a typical run looks like in the terminal:

```
[0/4] Fetching Google Sheet …
      Sheet ID : 1Kb7AcEmVZDLg5lgV6pHJQ8lE40sJMIL0
      Using cached sheet (14 min old) → gsheet_cache/gsheet_1Kb7AcEmVZDLg5lgV6pHJQ8lE40sJMIL0.xlsx

[1/4] Reading Excel / CSV input files …
      56 unique courses loaded.
      239 Drive hyperlinks extracted across all asset columns.

[2/4] Scanning 2 drive(s) …
Reading Excel/CSV files: 100%|██████████| 1/1 [00:01<00:00,  1.23s/file]
Scanning drives: 100%|██████████| 2/2 [00:18<00:00,  9.04s/drive]
      142,837 total paths indexed (files + folders).

[3/4] Connecting to Google Drive …
      Using public HTTP fallback.

[4/4] Auditing 56 course(s) across 6 asset types …
Auditing: 100%|██████████| 56/56 [00:42<00:00,  1.31course/s]

[5/4] Download step skipped (use --download or --download_all to enable).

  Courses with NO local match (check folder names on your drive):
  --------------------------------------------------------------------------
  • Advanced Python for Data Science
    Links in Excel : 4/6
    Drive folder   : https://drive.google.com/drive/folders/1AbC...
  --------------------------------------------------------------------------

+-------------------------------------------+
|            AUDIT SUMMARY                  |
+-------------------------------------------+
|  Total courses          :       56        |
|  Courses with links     :       51        |
|  All assets found local :       38        |
|  Some assets found local:       11        |
|  No local assets found  :        7        |
+-------------------------------------------+

  Report saved → ./availability_report.xlsx
```

---

### With `--download` active

```
[5/4] 23 missing asset(s) across 7 course(s) → '/run/media/duffer/One Touch A'
      Min free space : 5.0 GB  (whole course is skipped if space is low — no cross-disk splits)
      Disk [████████████████░░░░░░░░░░░░░░] 53.2% used  |  186.40 GB free / 931.51 GB total

Courses: 100%|██████████| 7/7 [04:13<00:00, 36.1s/course]

      Disk [█████████████████░░░░░░░░░░░░░] 55.8% used  |  162.10 GB free / 931.51 GB total

+-------------------------------------------+
|            AUDIT SUMMARY                  |
+-------------------------------------------+
|  Total courses          :       56        |
|  Courses with links     :       51        |
|  All assets found local :       38        |
|  Some assets found local:       11        |
|  No local assets found  :        7        |
|  Downloaded OK          :       21        |
|  Download failures      :        2        |
|  Courses skipped (low space):    0        |
+-------------------------------------------+

  Report saved → ./availability_report.xlsx
```

---

### With `--no_cache` (fresh scan)

```
[0/4] Fetching Google Sheet …
      Sheet ID : 1Kb7AcEmVZDLg5lgV6pHJQ8lE40sJMIL0
      Fetching : https://docs.google.com/spreadsheets/d/1Kb7.../export?format=xlsx
      Saved 184,320 bytes → gsheet_cache/gsheet_1Kb7AcEmVZDLg5lgV6pHJQ8lE40sJMIL0.xlsx

[2/4] Scanning 2 drive(s) …
Scanning drives: 100%|██████████| 2/2 [01:34<00:00, 47.1s/drive]
Building index: 100%|██████████| 290,441/290,441 [00:03<00:00, 95,102path/s]
      290,441 total paths indexed (files + folders).
```

---

## Output Report

The generated `availability_report.xlsx` has one row per course.

### Columns

| Column | Description |
|--------|-------------|
| Course | Course name from the sheet |
| Semester | Semester (e.g. S1) |
| Term | Term (e.g. T1) |
| Status | Active / Inactive etc. |
| `<Asset>_Local` | `Yes` / `Yes (Downloaded)` / `No` |
| `<Asset>_Local_Path` | Full path to matched folder on disk |
| `<Asset>_Drive` | Drive link status (see below) |

The six asset types are: **Course_Outline**, **PPTs**, **Written_Assets**, **Final_Videos**, **Raw_Videos**, **Course_Artifacts** — each with its own three columns above.

### Drive status values

| Status | Meaning |
|--------|---------|
| `Available` | Folder is accessible and publicly shared |
| `Missing` | Folder is private, deleted, or login-protected |
| `Broken Link` | URL could not be parsed or the request failed |
| `No Link` | No Google Drive link in this spreadsheet cell |
| `Not Checked` | Link present but Drive checking is disabled |

### Row colours

| Colour | Meaning |
|--------|---------|
| Green (#C6EFCE) | All assets found locally **and** all Drive links are live |
| Yellow (#FFEB9C) | Some assets found or some Drive links are live |
| Red (#FFC7CE) | Nothing found locally and no Drive links accessible |

### Example report snapshot

```
| Course                   | Sem | Term | Status | Course_Outline_Local | Course_Outline_Drive | PPTs_Local | PPTs_Drive | …
|--------------------------|-----|------|--------|----------------------|----------------------|------------|------------|---
| Intro to Python          | S1  | T1   | Active | Yes                  | Available            | Yes        | Available  | … ← GREEN row
| Data Science Basics      | S1  | T2   | Active | No                   | Available            | Yes        | Missing    | … ← YELLOW row
| Advanced ML              | S2  | T1   | Active | No                   | Missing              | No         | Missing    | … ← RED row
| Web Dev Fundamentals     | S1  | T1   | Active | Yes (Downloaded)     | Available            | Yes        | Available  | … ← GREEN row
```

---

## How It Works

```
 Google Sheets URL
        │
        ▼
 [0] Download .xlsx ──► cached for 1 hour (--no_cache to refresh)
        │
        ▼
 [1] Parse courses + Drive folder links from all asset columns
        │
        ▼
 [2] Scan local drives ──► index all files + folders (cached 24h)
        │
        ▼
 [3] For each course × 6 asset types:
      ├─ Fuzzy-match course name → local folder path
      └─ HTTP check Drive folder URL → Available / Missing / Broken
        │
        ▼
 [4] Generate colour-coded Excel report
        │
        ▼
 [5] (optional --download / --download_all)
      └─ For every qualifying asset with a Drive link:
           Check free disk space (skip whole course if below threshold)
           Download entire Drive folder → <drive>/<Course Name>/<Asset>/
```

### Fuzzy matching explained

The tool uses `rapidfuzz` to match course names to folder names on disk. The `--fuzzy_threshold` (default 75) controls how strict the match must be:

| Threshold | Effect |
|-----------|--------|
| 90–100 | Very strict — folder name must almost exactly match the course name |
| 75 (default) | Balanced — handles minor differences in spacing, punctuation |
| 50–65 | Permissive — useful when folder names are abbreviated or reordered |

If a course shows up in the "NO local match" list, try lowering the threshold:

```bash
python audit_checker.py --fuzzy_threshold 60
```

---

## Project Structure

```
audit-checker/
├── audit_checker.py         # Main script — run this
├── config.example.json      # Configuration template — copy to config.json
├── requirements.txt         # Python dependencies
├── settings.yaml            # pydrive2 OAuth settings (optional)
└── README.md

# Created at runtime — gitignored:
├── config.json              # Your personal config (copy from config.example.json)
├── gsheet_cache/            # Cached Google Sheet downloads
├── availability_report.xlsx # Generated audit report
├── audit_checker.log        # Runtime log
└── .drive_index_cache.pkl   # Cached drive scan index
```

---

## What NOT to Commit

The following are listed in `.gitignore` and should **never** be pushed:

| File | Why |
|------|-----|
| `config.json` | Contains your personal drive paths and sheet URL |
| `credentials.json` | Google OAuth credentials — treat like a password |
| `mycreds.txt` | Cached OAuth token |
| `gsheet_cache/` | Downloaded sheet data |
| `availability_report.xlsx` | Generated output |
| `.drive_index_cache.pkl` | Local drive scan cache |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | DataFrame handling and Excel writing |
| `openpyxl` | Read Excel hyperlinks + write styled reports |
| `rapidfuzz` | Fast fuzzy matching for course name → folder name |
| `tqdm` | Progress bars |
| `gdown` | Download entire Google Drive folders |
| `pydrive2` | *(optional)* Authenticated Google Drive API access |

Install all at once:

```bash
pip install -r requirements.txt
```

---

## License

MIT
