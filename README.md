# Coursera Audit Checker

> Automatically cross-check your Coursera course materials on Google Drive against files stored on local hard drives — and download anything that's missing.

Reads a Google Sheet that lists courses and their Google Drive folder links across six asset types, scans your connected drives for matching content, checks whether every Drive folder is still live, and produces a colour-coded Excel report. Missing assets can be downloaded directly to the correct course folder on your drive.

---

## Features

- **Google Sheets as input** — paste a sharing URL, the sheet is fetched automatically (no manual downloads)
- **Six asset types per course** — Course Outline, PPTs, Written Assets, Final Videos, Raw Videos, Course Artifacts
- **Live Drive link checking** — each folder URL is verified as Available, Missing, or Broken
- **Local drive scanning** — recursively indexes connected hard drives using fuzzy name matching
- **Auto-download missing assets** — downloads entire Google Drive folders to the correct course subfolder on your drive via `gdown`
- **Smart caching** — drive index and sheet download are cached so repeated runs are fast
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
# Clone the repo
git clone https://github.com/YOUR_USERNAME/audit-checker.git
cd audit-checker

# (Recommended) create a virtual environment
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Make the script executable (Linux / macOS)
chmod +x audit_checker.py
```

---

## Configuration

Copy the example config and fill in your values:

```bash
cp config.example.json config.json
```

Then open `config.json` and set three things:

```json
{
  "gsheet_url": "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit",

  "drives": [
    "/run/media/yourname/One Touch A",
    "/run/media/yourname/One Touch B"
  ],

  "google_drive": {
    "enabled": true,
    "download_dest": "/run/media/yourname/One Touch A"
  }
}
```

| Key | Description |
|-----|-------------|
| `gsheet_url` | Full URL of your Google Sheet (must be set to *Anyone with link can view*) |
| `drives` | Mount paths of your connected hard drives to scan |
| `google_drive.enabled` | `true` to check each Drive link live (recommended) |
| `google_drive.download_dest` | Drive path where missing files will be downloaded |

### Google Sheet format

The sheet must have these column headers (names are configurable in `config.json`):

| Course | Sem | Term | Status | Course Outline | PPTs | Written Assets (PQ, GQ, DP) | Final Videos | Raw Videos | Course Artifacts Link |
|--------|-----|------|--------|---------------|------|-----------------------------|-------------|------------|----------------------|
| Intro to Programming | S1 | T1 | Active | [link] | [link] | [link] | [link] | [link] | [link] |

Each asset cell should contain a **hyperlinked label** pointing to a Google Drive folder — the tool extracts the real URL automatically.

---

## Usage

```bash
./audit_checker.py --help      # show all commands
```

### Common commands

| Command | What it does |
|---------|-------------|
| `./audit_checker.py` | Standard audit — fetch sheet, scan drives, check all links, save report |
| `./audit_checker.py --download` | Audit + download every missing asset from Google Drive |
| `./audit_checker.py --no_cache` | Force fresh sheet download and full drive rescan |
| `./audit_checker.py --no_cache --download` | Full fresh run and download everything missing |

### All flags

```
INPUT
  --gsheet_url URL        Google Sheets URL (overrides config for this run)
  --excel_dir DIR         Folder with local .xlsx/.csv files (fallback input)
  --config FILE           Path to a custom config.json

DRIVES
  --drives PATH [PATH …]  Drive paths to scan (overrides config for this run)

OUTPUT
  --output FILE           Report save path  (default: ./availability_report.xlsx)

DOWNLOAD
  --download              Download missing assets after auditing
  --download_dest DIR     Drive to save downloads to (overrides config)

CACHE
  --no_cache              Ignore all cached data — re-download sheet + rescan drives

ADVANCED
  --fuzzy_threshold N     Name-match sensitivity 0–100  (default: 75)
  --workers N             Parallel scan workers  (default: CPU count)
  --log_level LEVEL       DEBUG | INFO | WARNING | ERROR  (default: INFO)
```

---

## Output Report

The generated `availability_report.xlsx` has one row per course:

### Columns

| Column | Description |
|--------|-------------|
| Course, Semester, Term, Status | Pulled directly from the sheet |
| `<Asset>_Local` | `Yes` / `No` — found on a local drive |
| `<Asset>_Local_Path` | Full path to the matched folder on disk |
| `<Asset>_Drive` | Drive link status (see below) |

### Drive status values

| Status | Meaning |
|--------|---------|
| `Available` | Folder is accessible and publicly shared |
| `Missing` | Folder is private, deleted, or login-protected |
| `Broken Link` | URL could not be parsed or the request failed |
| `No Link` | No Google Drive link in this spreadsheet cell |

### Row colours

| Colour | Meaning |
|--------|---------|
| 🟢 Green | All assets found locally **and** all Drive links are live |
| 🟡 Yellow | Some assets found or some Drive links are live |
| 🔴 Red | Nothing found locally and no Drive links accessible |

---

## How it works

```
 Google Sheets URL
        │
        ▼
 [0] Download .xlsx ──► cached for 1 hour (--no_cache to refresh)
        │
        ▼
 [1] Parse 56 courses + 239 Drive folder links
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
 [5] (optional --download)
      └─ For every asset that is Available on Drive but missing locally:
           Download entire Drive folder → <drive>/<Course Name>/<Asset>/
```

---

## Project Structure

```
audit-checker/
├── audit_checker.py         # Main script — run this
├── config.example.json      # Configuration template — copy to config.json
├── requirements.txt         # Python dependencies
├── settings.yaml            # pydrive2 OAuth settings (optional)
├── excel/                   # Drop fallback .xlsx/.csv files here
│   └── .gitkeep
└── README.md

# These are created at runtime and are gitignored:
├── config.json              # Your personal config (copy from config.example.json)
├── gsheet_cache/            # Cached Google Sheet downloads
├── availability_report.xlsx # Generated audit report
├── audit_checker.log        # Runtime log
└── .drive_index_cache.pkl   # Cached drive scan index
```

---

## What NOT to commit

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

---

## License

MIT
