# Coursera Audit Checker & Downloader

> Two tools in one repo:
> - **`audit_checker.py`** — cross-checks course materials on Google Drive against your local hard drives and produces a colour-coded Excel report
> - **`downloader.py`** — downloads every course from Google Drive, organised course-by-course with one folder per asset type, spread across two external drives

---

## Table of Contents

- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [rclone Setup (Required for Downloader)](#rclone-setup-required-for-downloader)
- [Downloader — Complete Guide](#downloader--complete-guide)
  - [What it downloads](#what-it-downloads)
  - [Folder structure on disk](#folder-structure-on-disk)
  - [Disk assignment logic](#disk-assignment-logic)
  - [Step-by-step commands](#step-by-step-commands)
  - [All downloader commands](#all-downloader-commands)
  - [Progress report](#progress-report)
  - [Resuming after interruption](#resuming-after-interruption)
  - [Troubleshooting](#troubleshooting)
- [Audit Checker — Complete Guide](#audit-checker--complete-guide)
  - [Configuration](#configuration)
  - [All audit commands](#all-audit-commands)
  - [Output report](#output-report)
- [What NOT to Commit](#what-not-to-commit)
- [Dependencies](#dependencies)

---

## Project Structure

```
audit-checker/
├── audit_checker.py         # Audit script
├── downloader.py            # Downloader script
├── config.example.json      # Config template — copy to config.json
├── requirements.txt         # Python dependencies
├── settings.yaml            # pydrive2 OAuth settings (optional)
└── README.md

# Created at runtime — gitignored:
├── config.json              # Your personal config
├── gsheet_cache/            # Cached Google Sheet downloads
├── availability_report.xlsx # Audit report output
├── download_progress.xlsx   # Downloader progress report
├── download_results.json    # Per-course download state
├── disk_assignment.json     # Which course is assigned to which disk
├── audit_checker.log        # Audit runtime log
├── downloader.log           # Downloader runtime log
└── .drive_index_cache.pkl   # Cached local drive scan index
```

---

## Requirements

- Python 3.10+
- rclone (for downloading — already installed on this system)
- Two external drives mounted:
  - `/run/media/duffer/One Touch`
  - `/run/media/duffer/One Touch A`

---

## Installation

```bash
# 1. Go to the project folder
cd /home/duffer/Gemini/Project_1\(complete\)/audit-checker

# 2. Install Python dependencies
pip install -r requirements.txt
```

---

## rclone Setup (Required for Downloader)

The downloader uses **rclone** to download from Google Drive. This is needed because the Drive folders are shared with your account specifically (not public), so a browser login is required. rclone handles this via a one-time OAuth login.

**You only need to do this once.**

### Step 1 — Run rclone config

```bash
rclone config
```

### Step 2 — Follow the prompts exactly

```
No remotes found, make a new one?
→ Press: n  (new remote)

name> gdrive
→ Type: gdrive  then press Enter

Storage>
→ Type: drive  then press Enter  (this is Google Drive)

Google Application Client Id - leave blank normally.
client_id>
→ Just press Enter

Google Application Client Secret - leave blank normally.
client_secret>
→ Just press Enter

Scope that rclone should use when requesting access from drive.
scope>
→ Type: 1  then press Enter  (full access — needed to download shared files)

root_folder_id>
→ Just press Enter

service_account_file>
→ Just press Enter

Edit advanced config?
→ Press: n

Use web browser to authenticate rclone?
→ Press: y
→ A browser window will open — log in with the Google account that has access to the course files
→ Click "Allow"
→ Come back to the terminal

Configure this as a Shared Drive (Team Drive)?
→ Press: n

Keep this "gdrive" remote?
→ Press: y

→ Press: q  to quit config
```

### Step 3 — Test rclone is working

```bash
rclone lsd gdrive:
```

This should list the folders in your Google Drive. If you see output, rclone is correctly authenticated and ready.

---

## Downloader — Complete Guide

### What it downloads

The Google Sheet (`RCA` tab) has **56 courses** total:
- **41 Completed** — all have Google Drive folder links → these get downloaded
- **15 In Production** — no links yet → automatically skipped

For each completed course, it downloads **6 asset types**:

| Asset Folder | Contents |
|---|---|
| `Course Outline` | COD / course outline document |
| `PPTs` | Presentation slides |
| `Written Assets` | Practice questions, graded questions, discussion prompts |
| `Final Videos` | Final edited video files |
| `Raw Videos` | Raw footage |
| `Course Artifacts` | Full production folder (Premiere files, AEP, exports, etc.) |

### Folder structure on disk

Every course gets its own folder, with each asset type as a subfolder inside it:

```
/run/media/duffer/One Touch/
  Downloaded Courses/
    Writing Practice/
      Course Outline/
      PPTs/
      Written Assets/
      Final Videos/
      Raw Videos/
      Course Artifacts/
    Introduction to Programming/
      Course Outline/
      PPTs/
      Written Assets/
      Final Videos/
      Raw Videos/
      Course Artifacts/
    ...

/run/media/duffer/One Touch A/
  Downloaded Courses/
    Discrete Mathematics/
      Course Outline/
      PPTs/
      Written Assets/
      Final Videos/
      Raw Videos/
      Course Artifacts/
    ...
```

### Disk assignment logic

- Courses alternate between the two disks (~20–21 per disk) to balance load
- Once a course is assigned to a disk, it **never moves** — it stays on that disk forever
- If a course folder already exists on a disk from a previous run, it uses that same disk
- A course is **never split across two disks** — all 6 asset folders for one course live on one disk
- Assignments are saved in `disk_assignment.json`

### Step-by-step commands

**Step 1 — Make sure both drives are plugged in**
```bash
ls "/run/media/duffer/One Touch"
ls "/run/media/duffer/One Touch A"
```
Both should list files. If either is missing, plug in the drive and try again.

**Step 2 — Go to the project folder**
```bash
cd /home/duffer/Gemini/Project_1\(complete\)/audit-checker
```

**Step 3 — Test with one course first**
```bash
python3 downloader.py --course "Writing Practice"
```
Wait for it to finish. Then check the result:
```bash
ls "/run/media/duffer/One Touch/Downloaded Courses/Writing Practice/"
```
You should see: `Course Outline/  PPTs/  Written Assets/  Final Videos/  Raw Videos/  Course Artifacts/`

**Step 4 — If the test looks good, run all 41 courses**
```bash
python3 downloader.py
```
This will run for a long time (many GBs of video). You can stop it at any time with `Ctrl+C` and resume later — it picks up where it left off.

**Step 5 — Check progress at any time**

Open `download_progress.xlsx` to see which courses are done, which failed, and which are still pending.

Or regenerate the report without downloading:
```bash
python3 downloader.py --report-only
```

### All downloader commands

```bash
# Download a single course (test before full run)
python3 downloader.py --course "Writing Practice"

# Download by partial name match (case-insensitive)
python3 downloader.py --course "discrete"

# Download TWO or more courses at once — repeat --course for each
python3 downloader.py --course "Discrete Mathematics" --course "Introduction to Programming"

# Download all 41 completed courses
python3 downloader.py

# Preview what would be downloaded — no actual downloads
python3 downloader.py --dry-run

# Force re-download the Google Sheet before running
python3 downloader.py --refresh-sheet

# Regenerate the Excel progress report without downloading anything
python3 downloader.py --report-only

# Combine flags — refresh sheet + download one course
python3 downloader.py --refresh-sheet --course "Writing Practice"
```

### Progress report

`download_progress.xlsx` is automatically generated/updated after every run.

| Column | Possible values |
|---|---|
| Course | Course name |
| Status | `Completed` / `In Production` |
| Assigned Disk | `One Touch` / `One Touch A` |
| Disk Path | Full path to the course folder on disk |
| Course Outline | `ok` / `skipped` / `no_link` / `failed` / `not started` |
| PPTs | same as above |
| Written Assets | same as above |
| Final Videos | same as above |
| Raw Videos | same as above |
| Course Artifacts | same as above |
| Overall | `Complete` / `Partial` / `Failed` / `Not Started` / `In Production` |

**Colour coding:**

| Colour | Meaning |
|---|---|
| Green | Downloaded successfully |
| Yellow | Already present on disk — skipped to avoid re-downloading |
| Red | Download failed (check `downloader.log` for details) |
| Orange | No Drive link in the sheet, or course is In Production |
| Grey | Not yet started |

### Resuming after interruption

The script saves progress after every completed course. If you stop it mid-run (power cut, `Ctrl+C`, etc.), just run it again:

```bash
python3 downloader.py
```

It will automatically skip every course whose folder already has files on disk and continue from where it left off.

If a specific course partially downloaded and you want to re-download it cleanly:

```bash
# Delete the partial folder
rm -rf "/run/media/duffer/One Touch/Downloaded Courses/Writing Practice"

# Remove it from the results file so the script tries again
# (or just edit download_results.json and delete that course's entry)

# Re-run
python3 downloader.py --course "Writing Practice"
```

### Troubleshooting

**"rclone: command not found"**
```bash
sudo apt install rclone
```

**"Failed to copy: directory not found" or similar rclone error**

Your rclone remote is not set up. Run:
```bash
rclone config
```
and follow the [rclone setup steps](#rclone-setup-required-for-downloader) above.

**"NOTICE: gdrive: 'root_folder_id' is not empty, disabling teamdrives"**

This is just an info message, not an error. Downloads will still work.

**A course shows as `failed` in the progress report**

1. Check the log for details:
   ```bash
   tail -50 downloader.log
   ```
2. Make sure rclone is authenticated:
   ```bash
   rclone lsd gdrive:
   ```
3. Re-run just that course:
   ```bash
   python3 downloader.py --course "Course Name Here"
   ```

**Drive not found / not mounted**

Plug in the drive and check it appears:
```bash
ls /run/media/duffer/
```
Both `One Touch` and `One Touch A` must be present before running.

**Google Sheet failed to download**

Force a fresh fetch:
```bash
python3 downloader.py --refresh-sheet
```

---

## Audit Checker — Complete Guide

`audit_checker.py` reads the Google Sheet, scans your local drives for matching content, checks whether every Drive link is still live, and produces a colour-coded Excel report (`availability_report.xlsx`).

### Configuration

```bash
# Copy the example config
cp config.example.json config.json
```

Open `config.json` and fill in:

```json
{
  "gsheet_url": "https://docs.google.com/spreadsheets/d/1Kb7AcEmVZDLg5lgV6pHJQ8lE40sJMIL0/edit",

  "drives": [
    "/run/media/duffer/One Touch",
    "/run/media/duffer/One Touch A"
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
| `drives` | Mount paths of your external drives to scan |
| `output` | Where to save the Excel report |
| `google_drive.enabled` | `true` to verify each Drive link is live |
| `scanning.fuzzy_threshold` | Name-match sensitivity 0–100 (default: 75) |
| `scanning.cache_max_age_hours` | How long the drive scan cache is valid (default: 24h) |
| `scanning.gsheet_cache_hours` | How long the sheet cache is valid (default: 1h) |

### All audit commands

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
python3 audit_checker.py --output /home/duffer/reports/march_audit.xlsx

# Scan specific drives for this run only
python3 audit_checker.py --drives "/run/media/duffer/One Touch" "/run/media/duffer/One Touch A"

# Use a different Google Sheet for this run
python3 audit_checker.py --gsheet_url "https://docs.google.com/spreadsheets/d/NEW_ID/edit"

# Relax name matching (catches folders with slightly different names)
python3 audit_checker.py --fuzzy_threshold 60

# Strict name matching (reduces false positives)
python3 audit_checker.py --fuzzy_threshold 90

# Full debug log output
python3 audit_checker.py --log_level DEBUG

# Use a different config file
python3 audit_checker.py --config /home/duffer/configs/semester2.json
```

### Output report

`availability_report.xlsx` has one row per course with these columns for each of the 6 asset types:

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

These files are in `.gitignore` and must never be pushed to git:

| File | Why |
|---|---|
| `config.json` | Contains your personal drive paths and sheet URL |
| `credentials.json` | Google OAuth credentials — treat like a password |
| `mycreds.txt` | Cached OAuth token |
| `gsheet_cache/` | Downloaded sheet data |
| `availability_report.xlsx` | Generated output |
| `download_progress.xlsx` | Generated downloader report |
| `download_results.json` | Downloader state |
| `disk_assignment.json` | Disk assignment state |
| `.drive_index_cache.pkl` | Local drive scan cache |

---

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | DataFrame handling and Excel writing |
| `openpyxl` | Read Excel hyperlinks + write styled reports |
| `rapidfuzz` | Fuzzy matching for course name → folder name |
| `tqdm` | Progress bars |
| `pydrive2` | Optional authenticated Drive API access |
| `rclone` | Primary downloader — authenticated via OAuth, handles files shared with your account |

Install Python packages:
```bash
pip install -r requirements.txt
```

Install rclone (already installed on this system):
```bash
sudo apt install rclone
```

---

## License

MIT
