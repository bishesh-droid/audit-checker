# Audit Checker

A production-ready file/program availability auditor.

Reads program names and Google Drive links from Excel/CSV files, scans local drives recursively (with multiprocessing + caching), checks Google Drive availability, and produces a colour-coded Excel report.

---

## Features

- Scans one or more local drives/folders recursively using parallel workers
- Fuzzy filename matching via `rapidfuzz` (configurable threshold)
- Google Drive availability checks (authenticated via pydrive2 or public HTTP fallback)
- Persistent drive index cache (pickle) — skips expensive rescans within a configurable time window
- Colour-coded Excel report output (green = found both, yellow = found one, red = missing)
- Fully configurable via `config.json` and/or CLI arguments

---

## Project Structure

```
audit-checker/
├── audit_checker.py        # Main script
├── config.json             # Runtime configuration
├── settings.yaml           # pydrive2 OAuth settings
├── requirements.txt        # Python dependencies
├── excel/                  # Drop your input .xlsx / .csv files here
│   └── .gitkeep
└── README.md
```

> **Do not commit** `credentials.json` or `mycreds.txt` — they are listed in `.gitignore`.

---

## Setup

### 1. Install dependencies

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Prepare input files

Place your Excel (`.xlsx`) or CSV (`.csv`) files inside the `excel/` folder.
Each file must have at least a **Program_Name** column. An optional **Drive_Link** column enables Google Drive checks.

| Program_Name   | Drive_Link                                      |
|----------------|-------------------------------------------------|
| MyApp          | https://drive.google.com/file/d/ABC123.../view  |
| AnotherTool    |                                                 |

### 3. Configure `config.json`

Edit `config.json` to point to your local drives and tune other settings:

```json
{
  "excel_dir": "./excel",
  "drives": ["/mnt/hdd1", "/mnt/hdd2"],
  "output": "./availability_report.xlsx"
}
```

### 4. (Optional) Set up Google Drive access

1. Go to [Google Cloud Console](https://console.cloud.google.com/) → APIs & Services → Credentials
2. Create an **OAuth 2.0 Client ID** (Desktop app) and download `credentials.json`
3. Place `credentials.json` in the project root
4. On the first run you will be prompted to authorise in the browser; a `mycreds.txt` token is then cached automatically

---

## Usage

```bash
# Basic — scan drives defined in config.json
python audit_checker.py

# Override drives and excel dir from CLI
python audit_checker.py --excel_dir ./excel --drives /mnt/hdd1 /mnt/hdd2

# Windows drives with custom output
python audit_checker.py --excel_dir ./data --drives D:\\ E:\\ --output ./report.xlsx

# Force rescan (ignore cache) and disable Google Drive checking
python audit_checker.py --excel_dir ./excel --drives /data --no_cache --no_gdrive

# Verbose debug logging
python audit_checker.py --excel_dir ./excel --drives /data --log_level DEBUG

# Use a custom config file
python audit_checker.py --config my_config.json
```

### CLI Arguments

| Argument            | Description                                            |
|---------------------|--------------------------------------------------------|
| `--config FILE`     | Path to a JSON config file                             |
| `--excel_dir DIR`   | Directory containing input Excel/CSV files             |
| `--drives PATH...`  | Local drive or folder paths to scan                    |
| `--output FILE`     | Output Excel report path                               |
| `--program_col COL` | Excel column name for program names (default: `Program_Name`) |
| `--drive_col COL`   | Excel column name for Drive links (default: `Drive_Link`) |
| `--fuzzy_threshold` | Minimum fuzzy-match score 0–100 (default: 80)          |
| `--workers N`       | Parallel scan processes (default: CPU count)           |
| `--no_cache`        | Ignore cache and force a full rescan                   |
| `--no_gdrive`       | Disable Google Drive checks                            |
| `--log_level LEVEL` | Logging verbosity: DEBUG / INFO / WARNING / ERROR      |

---

## Output Report

The generated `availability_report.xlsx` contains one row per program with:

| Column                  | Description                                      |
|-------------------------|--------------------------------------------------|
| `Program_Name`          | Program/file name from the input                 |
| `Found_in_Hard_Drive`   | Yes / No                                         |
| `Found_in_Google_Drive` | Yes / No                                         |
| `Local_File_Path`       | Full path of the matched local file              |
| `Drive_Status`          | Available / Missing / Broken Link / Not Checked  |
| `Match_Confidence`      | Fuzzy match score (100 = exact)                  |

Row colours:
- **Green** — found on both local drives and Google Drive
- **Yellow** — found in one location only
- **Red** — not found anywhere

---

## License

MIT
