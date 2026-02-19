#!/usr/bin/env python3
"""
audit_checker.py
================
Production-ready file/program availability auditor.

Reads program names and Google Drive links from Excel/CSV files,
scans local drives recursively (with multiprocessing + caching),
checks Google Drive availability, and produces a colour-coded Excel report.

Author  : Senior Python Engineer
Python  : 3.10+
License : MIT
"""

from __future__ import annotations

import argparse
import json
import logging
import multiprocessing
import os
import pickle
import re
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl.styles import Font, PatternFill
from rapidfuzz import fuzz, process as fz_process
from tqdm import tqdm

# ── Optional pydrive2 import (graceful fallback) ──────────────────────────────
try:
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive

    GDRIVE_AVAILABLE = True
except ImportError:
    GDRIVE_AVAILABLE = False

# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING SETUP
# ═══════════════════════════════════════════════════════════════════════════════

def setup_logging(log_file: str = "audit_checker.log", level: str = "INFO") -> logging.Logger:
    """
    Configure structured logging to a rotating file and to the console
    (console shows WARNING+ only to keep stdout clean during progress bars).

    Args:
        log_file: Path to the log file.
        level:    Root log level string (DEBUG, INFO, WARNING, ERROR).

    Returns:
        Configured Logger instance.
    """
    _logger = logging.getLogger("audit_checker")
    _logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    _logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler — full verbosity
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    _logger.addHandler(fh)

    # Console handler — warnings only (so tqdm progress bars stay clean)
    ch = logging.StreamHandler(sys.stderr)
    ch.setFormatter(fmt)
    ch.setLevel(logging.WARNING)
    _logger.addHandler(ch)

    return _logger


logger = setup_logging()  # default; reconfigured in main() once args are parsed

# ═══════════════════════════════════════════════════════════════════════════════
#  DATA STRUCTURES
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class ProgramEntry:
    """One row extracted from the input Excel / CSV files."""
    program_name: str
    drive_link:   Optional[str] = None
    source_file:  Optional[str] = None
    row_index:    Optional[int] = None


@dataclass
class AuditResult:
    """Consolidated audit result for a single program entry."""
    program_name:         str
    found_in_hard_drive:  str   = "No"
    found_in_google_drive: str  = "No"
    local_file_path:      str   = ""
    drive_status:         str   = "Not Checked"
    match_confidence:     float = 0.0


# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_CONFIG: dict = {
    "excel_dir": "./excel",
    "drives": [],
    "output": "./availability_report.xlsx",
    "columns": {
        "program_name": "Program_Name",
        "drive_link":   "Drive_Link",
    },
    "scanning": {
        "fuzzy_threshold":      80,
        "cache_file":           ".drive_index_cache.pkl",
        "cache_max_age_hours":  24,
        "max_workers":          multiprocessing.cpu_count(),
        "extensions_filter":    [],   # e.g. ["exe","apk","dmg"] — empty = all
    },
    "google_drive": {
        "enabled":          True,
        "credentials_file": "credentials.json",
        "settings_file":    "settings.yaml",
    },
    "logging": {
        "log_file": "audit_checker.log",
        "level":    "INFO",
    },
}


def _deep_merge(base: dict, override: dict) -> None:
    """Recursively merge *override* into *base* in-place (base is mutated)."""
    for key, value in override.items():
        if key in base and isinstance(base[key], dict) and isinstance(value, dict):
            _deep_merge(base[key], value)
        else:
            base[key] = value


def load_config(config_path: Optional[str] = None) -> dict:
    """
    Load configuration from a JSON file, deep-merging with DEFAULT_CONFIG.

    Search order:
        1. ``config_path`` argument (if given)
        2. ``config.json`` in the current working directory
        3. ``audit_config.json`` in the current working directory

    Args:
        config_path: Explicit path to a config JSON file.

    Returns:
        Merged configuration dictionary.
    """
    import copy
    config = copy.deepcopy(DEFAULT_CONFIG)

    candidates = [config_path] if config_path else ["config.json", "audit_config.json"]

    for candidate in candidates:
        if not candidate:
            continue
        p = Path(candidate)
        if p.exists():
            try:
                with open(p, encoding="utf-8") as fh:
                    user_cfg = json.load(fh)
                _deep_merge(config, user_cfg)
                logger.info("Loaded config from: %s", p)
                break
            except (json.JSONDecodeError, OSError) as exc:
                logger.warning("Could not load config '%s': %s", p, exc)

    return config


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL / CSV PARSING
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_single_file(
    filepath: Path,
    program_col: str,
    drive_col: Optional[str],
) -> list[ProgramEntry]:
    """
    Parse one Excel (.xlsx) or CSV file and return a list of ProgramEntry objects.

    Args:
        filepath:    Path to the file.
        program_col: Column name that holds program / file names.
        drive_col:   Column name that holds Google Drive links (may be None).

    Returns:
        List of ProgramEntry instances (empty if the file cannot be parsed).
    """
    try:
        if filepath.suffix.lower() == ".csv":
            df = pd.read_csv(filepath, dtype=str, keep_default_na=False)
        else:
            df = pd.read_excel(filepath, dtype=str, keep_default_na=False, engine="openpyxl")
    except Exception as exc:
        logger.error("Cannot read '%s': %s", filepath, exc)
        return []

    # Normalise column names (strip leading/trailing whitespace)
    df.columns = [str(c).strip() for c in df.columns]

    if program_col not in df.columns:
        logger.warning(
            "Column '%s' not found in '%s'. Available: %s",
            program_col, filepath.name, list(df.columns),
        )
        return []

    entries: list[ProgramEntry] = []
    for idx, row in df.iterrows():
        name = str(row[program_col]).strip()
        # Skip empty / sentinel values
        if not name or name.lower() in {"nan", "none", "n/a", ""}:
            continue

        drive_link: Optional[str] = None
        if drive_col and drive_col in df.columns:
            raw = str(row[drive_col]).strip()
            if raw and raw.lower() not in {"nan", "none", "n/a", ""}:
                drive_link = raw

        entries.append(
            ProgramEntry(
                program_name=name,
                drive_link=drive_link,
                source_file=filepath.name,
                row_index=int(idx),  # type: ignore[arg-type]
            )
        )

    return entries


def read_excel_files(
    excel_dir: str,
    program_col: str,
    drive_col: Optional[str] = None,
) -> list[ProgramEntry]:
    """
    Discover and parse all ``.xlsx`` and ``.csv`` files inside *excel_dir*
    (recursive search).

    Args:
        excel_dir:   Root directory that contains the input files.
        program_col: Column name for program names.
        drive_col:   Column name for Google Drive links.

    Returns:
        Combined list of ProgramEntry objects from all discovered files.

    Raises:
        FileNotFoundError: If *excel_dir* does not exist.
    """
    dir_path = Path(excel_dir)
    if not dir_path.exists():
        raise FileNotFoundError(f"Excel directory not found: {excel_dir!r}")

    files = sorted(
        list(dir_path.rglob("*.xlsx")) + list(dir_path.rglob("*.csv"))
    )

    if not files:
        logger.warning("No .xlsx / .csv files found in: %s", excel_dir)
        return []

    logger.info("Discovered %d input file(s) in: %s", len(files), excel_dir)

    all_entries: list[ProgramEntry] = []
    for fp in tqdm(files, desc="Reading Excel/CSV files", unit="file"):
        try:
            batch = _parse_single_file(fp, program_col, drive_col)
            logger.info("  %s → %d entries", fp.name, len(batch))
            all_entries.extend(batch)
        except Exception as exc:
            logger.error("Unexpected error parsing '%s': %s", fp, exc)

    # Deduplicate by (name, drive_link) while preserving order
    seen: set[tuple[str, Optional[str]]] = set()
    unique: list[ProgramEntry] = []
    for entry in all_entries:
        key = (entry.program_name.lower(), entry.drive_link)
        if key not in seen:
            seen.add(key)
            unique.append(entry)

    logger.info("Total unique program entries: %d", len(unique))
    return unique


# ═══════════════════════════════════════════════════════════════════════════════
#  LOCAL DRIVE SCANNER
# ═══════════════════════════════════════════════════════════════════════════════

# Directories to silently skip on any OS (avoids permission errors + noise)
_SKIP_DIRS: frozenset[str] = frozenset({
    "System Volume Information",
    "$RECYCLE.BIN",
    "$Recycle.Bin",
    "Windows",
    "WinSxS",
    "SoftwareDistribution",
    "Recovery",
    "hiberfil.sys",
    "pagefile.sys",
    "swapfile.sys",
    "proc",          # Linux virtual FS
    "sys",           # Linux virtual FS
    "dev",           # Linux virtual FS
})


def _scan_single_drive(args: tuple[str, list[str]]) -> list[str]:
    """
    Recursively walk one drive / folder and return all file paths found.

    Designed to be executed inside a subprocess (ProcessPoolExecutor).
    Silently skips directories that raise PermissionError.

    Args:
        args: Tuple of (root_path: str, extensions_filter: list[str]).
              extensions_filter is a list of lowercase extensions without dots
              (e.g. ["exe", "apk"]).  Empty list → include all files.

    Returns:
        List of absolute file path strings.
    """
    root_path, extensions_filter = args
    exts = frozenset(e.lower().lstrip(".") for e in extensions_filter) if extensions_filter else frozenset()
    found: list[str] = []

    for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=None, followlinks=False):
        # Prune directories we never want to descend into
        dirnames[:] = [
            d for d in dirnames
            if not d.startswith(".")
            and d not in _SKIP_DIRS
        ]

        for fname in filenames:
            if exts and Path(fname).suffix.lower().lstrip(".") not in exts:
                continue
            found.append(os.path.join(dirpath, fname))

    return found


def build_file_index(
    drive_paths: list[str],
    cache_file: str           = ".drive_index_cache.pkl",
    cache_max_age_hours: float = 24.0,
    max_workers: int           = 4,
    extensions_filter: Optional[list[str]] = None,
) -> dict[str, list[str]]:
    """
    Build an in-memory filename index for all files across the given drives.

    The index maps lowercased filename (and separately, lowercased stem) to a
    list of full absolute paths.  Results are cached to a pickle file so that
    re-runs within ``cache_max_age_hours`` skip the expensive disk scan.

    Args:
        drive_paths:          List of root drive / folder paths to scan.
        cache_file:           Path for the pickle cache file.
        cache_max_age_hours:  Max age of the cache before triggering rescan.
        max_workers:          Number of parallel worker processes.
        extensions_filter:    List of extensions to include (empty = all).

    Returns:
        Dict mapping ``{lowercase_name_or_stem: ["/full/path/a", ...]}``.
    """
    cache_path = Path(cache_file)
    drives_key = sorted(drive_paths)

    # ── Try loading a valid cache ──────────────────────────────────────────────
    if cache_path.exists():
        age_h = (time.time() - cache_path.stat().st_mtime) / 3600
        if age_h < cache_max_age_hours:
            try:
                with open(cache_path, "rb") as fh:
                    cached = pickle.load(fh)
                if cached.get("drives") == drives_key:
                    idx = cached["index"]
                    logger.info(
                        "Using cached index (%.1fh old, %d unique name keys).",
                        age_h, len(idx),
                    )
                    return idx
            except Exception as exc:
                logger.warning("Cache load failed (%s) — rescanning.", exc)

    # ── Full scan ─────────────────────────────────────────────────────────────
    logger.info(
        "Scanning %d drive(s) with up to %d worker(s) …",
        len(drive_paths), max_workers,
    )
    all_paths: list[str] = []
    scan_args = [(d, extensions_filter or []) for d in drive_paths]
    workers   = max(1, min(max_workers, len(drive_paths)))

    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_scan_single_drive, arg): arg[0] for arg in scan_args}

        with tqdm(total=len(futures), desc="Scanning drives", unit="drive") as pbar:
            for future in as_completed(futures):
                drv = futures[future]
                try:
                    result = future.result()
                    all_paths.extend(result)
                    logger.info("  Drive '%s': %d files.", drv, len(result))
                except Exception as exc:
                    logger.error("Scan failed for '%s': %s", drv, exc)
                finally:
                    pbar.update(1)

    logger.info("Total files indexed: %d", len(all_paths))

    # ── Build dict: lowercase name / stem → [paths] ───────────────────────────
    index: dict[str, list[str]] = {}
    for path in tqdm(all_paths, desc="Building index", unit="file", leave=False):
        p     = Path(path)
        full  = p.name.lower()          # e.g. "myapp.exe"
        stem  = p.stem.lower()          # e.g. "myapp"
        for key in (full, stem):
            index.setdefault(key, []).append(path)

    # ── Persist cache ─────────────────────────────────────────────────────────
    try:
        with open(cache_path, "wb") as fh:
            pickle.dump({"drives": drives_key, "index": index}, fh)
        logger.info("Index cached to: %s", cache_path)
    except Exception as exc:
        logger.warning("Could not write cache: %s", exc)

    return index


def find_file_locally(
    program_name: str,
    file_index:   dict[str, list[str]],
    fuzzy_threshold: int = 80,
) -> tuple[bool, str, float]:
    """
    Locate a program / file in the local file index.

    Strategy (in order):
    1. Exact match on full filename (lowercased).
    2. Exact match on stem (filename without extension, lowercased).
    3. Fuzzy match using ``rapidfuzz.fuzz.token_sort_ratio`` against a
       pre-filtered candidate pool (same first character + similar length).
    4. If step-3 yields no result, fuzzy match against ALL index keys.

    Args:
        program_name:    The filename or program name to search for.
        file_index:      Index returned by :func:`build_file_index`.
        fuzzy_threshold: Minimum rapidfuzz score (0–100) for a fuzzy match.

    Returns:
        Tuple of ``(found, matched_path, confidence)``.
        ``confidence`` is 100.0 for exact matches, 0–100 for fuzzy matches.
    """
    name_lower = program_name.lower()
    stem_lower = Path(program_name).stem.lower()

    # ── 1 & 2: Exact match ───────────────────────────────────────────────────
    for key in (name_lower, stem_lower):
        if key in file_index:
            return True, file_index[key][0], 100.0

    # ── 3: Fuzzy match with pre-filtering ────────────────────────────────────
    #   Only compare against names that share the same first character AND
    #   whose length is within ±50 % of the target.  This reduces the
    #   candidate pool dramatically on multi-TB drives (10M+ files).
    target      = stem_lower or name_lower
    target_len  = len(target)
    first_char  = target[0] if target else ""

    # Length tolerance: at least 3 chars or 50 % of target length
    tol = max(3, int(target_len * 0.5))

    candidates: list[str] = [
        k for k in file_index
        if k and k[0] == first_char
        and abs(len(k) - target_len) <= tol
    ]

    # ── 4: Fall back to full index if filtered set is too small ──────────────
    if len(candidates) < 5:
        candidates = list(file_index.keys())

    result = fz_process.extractOne(
        target,
        candidates,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=fuzzy_threshold,
    )

    if result:
        matched_name, score, _ = result
        return True, file_index[matched_name][0], float(score)

    return False, "", 0.0


# ═══════════════════════════════════════════════════════════════════════════════
#  GOOGLE DRIVE CHECKER
# ═══════════════════════════════════════════════════════════════════════════════

# Ordered list of regex patterns used to extract a Drive file/folder ID from a URL
_DRIVE_ID_PATTERNS: list[str] = [
    r"/file/d/([a-zA-Z0-9_-]{10,})",       # /file/d/<ID>/view
    r"/folders/([a-zA-Z0-9_-]{10,})",      # /folders/<ID>
    r"[?&]id=([a-zA-Z0-9_-]{10,})",        # ?id=<ID>  or  &id=<ID>
    r"/open\?id=([a-zA-Z0-9_-]{10,})",     # /open?id=<ID>
    r"^([a-zA-Z0-9_-]{25,})$",             # raw file ID (no URL)
]


def extract_drive_file_id(link: str) -> Optional[str]:
    """
    Extract the Google Drive file / folder ID from a Drive URL or raw ID.

    Supports all common Drive URL formats:
    - ``https://drive.google.com/file/d/<ID>/view``
    - ``https://drive.google.com/open?id=<ID>``
    - ``https://drive.google.com/drive/folders/<ID>``
    - ``https://docs.google.com/...?id=<ID>``
    - Raw 25+ character file IDs

    Args:
        link: A Google Drive URL string or raw file ID.

    Returns:
        Extracted file / folder ID string, or ``None`` if unrecognised.

    Examples:
        >>> extract_drive_file_id("https://drive.google.com/file/d/ABC123xyz/view")
        'ABC123xyz'
    """
    if not link:
        return None
    link = link.strip()
    for pattern in _DRIVE_ID_PATTERNS:
        m = re.search(pattern, link)
        if m:
            return m.group(1)
    return None


def authenticate_gdrive(
    credentials_file: str = "credentials.json",
    settings_file:    str = "settings.yaml",
) -> Optional[object]:
    """
    Authenticate with the Google Drive API using pydrive2.

    Attempts OAuth2 authentication.  If credentials / settings are missing
    the function returns ``None`` gracefully (public-link HTTP fallback is
    used instead).

    Args:
        credentials_file: Path to the OAuth2 client-secrets JSON file
                          downloaded from Google Cloud Console.
        settings_file:    Path to the pydrive2 ``settings.yaml`` file.

    Returns:
        Authenticated :class:`pydrive2.drive.GoogleDrive` instance,
        or ``None`` if authentication is impossible.
    """
    if not GDRIVE_AVAILABLE:
        logger.warning(
            "pydrive2 is not installed — Google Drive API unavailable. "
            "Install it with: pip install pydrive2"
        )
        return None

    creds_path    = Path(credentials_file)
    settings_path = Path(settings_file)
    saved_creds   = Path("mycreds.txt")

    if not creds_path.exists() and not settings_path.exists():
        logger.warning(
            "No Drive credentials found (%s / %s). "
            "Falling back to public-link HTTP checks only.",
            creds_path, settings_path,
        )
        return None

    try:
        gauth = GoogleAuth(
            settings_file=str(settings_path) if settings_path.exists() else None
        )

        if saved_creds.exists():
            gauth.LoadCredentialsFile(str(saved_creds))

        if gauth.credentials is None:
            gauth.LocalWebserverAuth()
        elif gauth.access_token_expired:
            gauth.Refresh()
        else:
            gauth.Authorize()

        gauth.SaveCredentialsFile(str(saved_creds))
        drive = GoogleDrive(gauth)
        logger.info("Google Drive authenticated successfully.")
        return drive

    except Exception as exc:
        logger.error("Google Drive authentication failed: %s", exc)
        return None


def _public_gdrive_check(file_id: str) -> str:
    """
    Check whether a *public* Google Drive file is accessible via HTTP HEAD.

    Does not download the file — only probes the export/redirect URL.

    Args:
        file_id: Google Drive file ID.

    Returns:
        One of ``"Available"``, ``"Missing"``, or ``"Broken Link"``.
    """
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    try:
        req = urllib.request.Request(url, method="HEAD")
        req.add_header("User-Agent", "Mozilla/5.0 (AuditChecker/1.0)")
        with urllib.request.urlopen(req, timeout=12) as resp:
            code = resp.status
            if code == 200:
                return "Available"
            if code in (403, 404):
                return "Missing"
            return "Broken Link"
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return "Missing"
        logger.debug("HTTP error for Drive ID %s: %s", file_id, exc)
        return "Broken Link"
    except Exception as exc:
        logger.debug("Network error checking Drive ID %s: %s", file_id, exc)
        return "Broken Link"


def check_gdrive_file(
    link:  Optional[str],
    drive: Optional[object] = None,
) -> tuple[str, str]:
    """
    Determine whether a Google Drive file exists and is accessible.

    Checks in this order:
    1. If an authenticated :class:`~pydrive2.drive.GoogleDrive` instance is
       provided, fetch file metadata via the Drive API.
    2. Fall back to a public HTTP HEAD request (no auth required).

    Args:
        link:  Google Drive URL or raw file ID.
        drive: Authenticated pydrive2 GoogleDrive instance (may be ``None``).

    Returns:
        Tuple of ``(found_in_google_drive, drive_status)`` where:
        - ``found_in_google_drive`` is ``"Yes"`` or ``"No"``.
        - ``drive_status`` is one of
          ``"Available"``, ``"Missing"``, ``"Broken Link"``, ``"Not Checked"``.
    """
    if not link:
        return "No", "Not Checked"

    file_id = extract_drive_file_id(link)
    if not file_id:
        logger.warning("Could not extract Drive file ID from link: %s", link)
        return "No", "Broken Link"

    # ── Authenticated Drive API check ─────────────────────────────────────────
    if drive is not None:
        try:
            file_obj = drive.CreateFile({"id": file_id})
            file_obj.FetchMetadata(fields="id,title,trashed")
            if file_obj.get("trashed", False):
                return "No", "Missing"
            return "Yes", "Available"
        except Exception as exc:
            err = str(exc).lower()
            if "404" in err or "not found" in err:
                return "No", "Missing"
            logger.warning(
                "Drive API error for ID '%s' — falling back to HTTP: %s",
                file_id, exc,
            )
            # Fall through to public HTTP check

    # ── Public HTTP fallback ──────────────────────────────────────────────────
    status = _public_gdrive_check(file_id)
    found  = "Yes" if status == "Available" else "No"
    return found, status


# ═══════════════════════════════════════════════════════════════════════════════
#  REPORT GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def generate_report(results: list[AuditResult], output_path: str) -> None:
    """
    Write the audit results to a colour-coded ``.xlsx`` Excel report.

    Row colours:
    - Green  → found on BOTH local drives AND Google Drive.
    - Yellow → found on ONE of the two locations.
    - Red    → not found anywhere.

    Args:
        results:     List of :class:`AuditResult` objects to write.
        output_path: Destination ``.xlsx`` file path.

    Raises:
        OSError: If the output file cannot be created or written.
    """
    records = [
        {
            "Program_Name":          r.program_name,
            "Found_in_Hard_Drive":   r.found_in_hard_drive,
            "Found_in_Google_Drive": r.found_in_google_drive,
            "Local_File_Path":       r.local_file_path,
            "Drive_Status":          r.drive_status,
            "Match_Confidence":      round(r.match_confidence, 2),
        }
        for r in results
    ]

    df = pd.DataFrame(records)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    # Write via pandas → openpyxl engine so we can apply formatting
    with pd.ExcelWriter(str(out), engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Audit Report")
        ws = writer.sheets["Audit Report"]

        # ── Auto-size columns ─────────────────────────────────────────────────
        for col_cells in ws.columns:
            max_len = max(
                (len(str(cell.value or "")) for cell in col_cells),
                default=10,
            )
            col_letter = col_cells[0].column_letter
            ws.column_dimensions[col_letter].width = min(max_len + 4, 80)

        # ── Colour fills ──────────────────────────────────────────────────────
        fill_green  = PatternFill("solid", fgColor="C6EFCE")  # both found
        fill_yellow = PatternFill("solid", fgColor="FFEB9C")  # one found
        fill_red    = PatternFill("solid", fgColor="FFC7CE")  # none found

        bold = Font(bold=True)

        # Style header row (row 1)
        for cell in ws[1]:
            cell.font = bold

        # Style data rows (rows 2 onwards)
        hd_col_idx = 1   # "Found_in_Hard_Drive"   — 0-based after Program_Name
        gd_col_idx = 2   # "Found_in_Google_Drive"

        for row in ws.iter_rows(min_row=2):
            hd = row[hd_col_idx].value
            gd = row[gd_col_idx].value

            if hd == "Yes" and gd == "Yes":
                fill = fill_green
            elif hd == "Yes" or gd == "Yes":
                fill = fill_yellow
            else:
                fill = fill_red

            for cell in row:
                cell.fill = fill

    logger.info("Report written to: %s", output_path)
    print(f"\n  Report saved → {output_path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  AUDIT ORCHESTRATION
# ═══════════════════════════════════════════════════════════════════════════════

def run_audit(
    excel_dir:            str,
    drives:               list[str],
    output:               str,
    program_col:          str,
    drive_col:            Optional[str],
    fuzzy_threshold:      int,
    cache_file:           str,
    cache_max_age_hours:  float,
    max_workers:          int,
    extensions_filter:    Optional[list[str]],
    gdrive_enabled:       bool,
    credentials_file:     str,
    settings_file:        str,
) -> None:
    """
    End-to-end orchestration: parse inputs → scan drives → check Drive → report.

    Args:
        excel_dir:           Directory with input Excel / CSV files.
        drives:              Validated local drive / folder paths.
        output:              Output ``.xlsx`` report path.
        program_col:         Excel column name for program names.
        drive_col:           Excel column name for Drive links (or ``None``).
        fuzzy_threshold:     Minimum fuzzy score for a local match (0–100).
        cache_file:          Path for the drive index pickle cache.
        cache_max_age_hours: Max cache age in hours before rescanning.
        max_workers:         Parallel worker count for drive scanning.
        extensions_filter:   Restrict local scan to these extensions.
        gdrive_enabled:      Whether to attempt Google Drive checks.
        credentials_file:    OAuth2 credentials path for pydrive2.
        settings_file:       pydrive2 settings YAML path.
    """
    # ── Step 1: Read input files ──────────────────────────────────────────────
    print("\n[1/4] Reading Excel / CSV input files …")
    entries = read_excel_files(excel_dir, program_col, drive_col)

    if not entries:
        logger.error("No program entries found — nothing to audit.")
        sys.exit(1)

    print(f"      {len(entries)} unique entries loaded.")

    # ── Step 2: Build local file index ───────────────────────────────────────
    file_index: dict[str, list[str]] = {}
    if drives:
        print(f"\n[2/4] Building file index for {len(drives)} drive(s) …")
        file_index = build_file_index(
            drives,
            cache_file=cache_file,
            cache_max_age_hours=cache_max_age_hours,
            max_workers=max_workers,
            extensions_filter=extensions_filter,
        )
        print(f"      {len(file_index):,} unique filenames in index.")
    else:
        print("\n[2/4] No local drives specified — skipping local scan.")

    # ── Step 3: Authenticate Google Drive ────────────────────────────────────
    gdrive = None
    if gdrive_enabled:
        print("\n[3/4] Connecting to Google Drive …")
        gdrive = authenticate_gdrive(credentials_file, settings_file)
        if gdrive:
            print("      Authenticated via pydrive2.")
        else:
            print("      Falling back to public-link HTTP checks.")
    else:
        print("\n[3/4] Google Drive checking is disabled.")

    # ── Step 4: Audit every entry ─────────────────────────────────────────────
    print(f"\n[4/4] Auditing {len(entries)} entries …")
    results: list[AuditResult] = []

    for entry in tqdm(entries, desc="Auditing", unit="entry"):
        result = AuditResult(program_name=entry.program_name)

        # Local drive check
        if file_index:
            found, path, confidence = find_file_locally(
                entry.program_name, file_index, fuzzy_threshold
            )
            if found:
                result.found_in_hard_drive = "Yes"
                result.local_file_path     = path
                result.match_confidence    = confidence

        # Google Drive check
        if entry.drive_link:
            result.found_in_google_drive, result.drive_status = check_gdrive_file(
                entry.drive_link, gdrive
            )
        else:
            result.drive_status = "No Link Provided"

        results.append(result)

    # ── Generate Excel report ─────────────────────────────────────────────────
    generate_report(results, output)

    # ── Print summary ─────────────────────────────────────────────────────────
    total        = len(results)
    hd_found     = sum(1 for r in results if r.found_in_hard_drive  == "Yes")
    gd_found     = sum(1 for r in results if r.found_in_google_drive == "Yes")
    both_found   = sum(1 for r in results if r.found_in_hard_drive == "Yes" and r.found_in_google_drive == "Yes")
    both_missing = sum(1 for r in results if r.found_in_hard_drive == "No"  and r.found_in_google_drive == "No")

    print(f"""
+-------------------------------------------+
|              AUDIT SUMMARY                |
+-------------------------------------------+
|  Total entries checked   : {total:>8,}    |
|  Found on local drives   : {hd_found:>8,}    |
|  Found on Google Drive   : {gd_found:>8,}    |
|  Found in BOTH locations : {both_found:>8,}    |
|  Missing everywhere      : {both_missing:>8,}    |
+-------------------------------------------+
""")


# ═══════════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════════

def _build_arg_parser() -> argparse.ArgumentParser:
    """Return the configured CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="audit_checker",
        description=(
            "Audit file/program availability across local drives and Google Drive\n"
            "using program names and Drive links read from Excel / CSV files."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic usage — scan two drives, read Excel from ./excel, write report
  python audit_checker.py --excel_dir ./excel --drives /mnt/hdd1 /mnt/hdd2

  # Windows drives with custom output
  python audit_checker.py --excel_dir ./data --drives D:\\ E:\\ --output ./report.xlsx

  # Custom config file
  python audit_checker.py --config my_config.json

  # Force rescan (ignore cache), disable Google Drive checking
  python audit_checker.py --excel_dir ./excel --drives /data --no_cache --no_gdrive

  # Verbose debug logging
  python audit_checker.py --excel_dir ./excel --drives /data --log_level DEBUG
        """,
    )

    parser.add_argument(
        "--config",
        metavar="FILE",
        help="Path to config.json (default: ./config.json if present)",
    )
    parser.add_argument(
        "--excel_dir",
        metavar="DIR",
        help="Directory containing Excel / CSV input files",
    )
    parser.add_argument(
        "--drives",
        nargs="+",
        metavar="PATH",
        help="Local drive or folder paths to scan (space-separated)",
    )
    parser.add_argument(
        "--output",
        metavar="FILE",
        help="Output Excel report file (default: ./availability_report.xlsx)",
    )
    parser.add_argument(
        "--program_col",
        metavar="COL",
        help="Excel column name for program / file names",
    )
    parser.add_argument(
        "--drive_col",
        metavar="COL",
        help="Excel column name for Google Drive links",
    )
    parser.add_argument(
        "--fuzzy_threshold",
        type=int,
        metavar="0-100",
        help="Minimum fuzzy-match score for local file matching (default: 80)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        metavar="N",
        help="Number of parallel scan processes (default: CPU count)",
    )
    parser.add_argument(
        "--no_cache",
        action="store_true",
        help="Ignore existing drive index cache and force a full rescan",
    )
    parser.add_argument(
        "--no_gdrive",
        action="store_true",
        help="Disable Google Drive availability checks entirely",
    )
    parser.add_argument(
        "--log_level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        metavar="LEVEL",
        help="Logging verbosity: DEBUG | INFO | WARNING | ERROR (default: INFO)",
    )

    return parser


def main() -> None:
    """
    CLI entry point.

    Parses arguments, merges with config file, validates paths, and
    delegates to :func:`run_audit`.
    """
    parser = _build_arg_parser()
    args   = parser.parse_args()

    # Load config (JSON file) then let CLI args override
    config = load_config(args.config)

    # Resolve final values (CLI > config > default)
    excel_dir    = args.excel_dir       or config["excel_dir"]
    drives       = args.drives          or config["drives"]
    output       = args.output          or config["output"]
    program_col  = args.program_col     or config["columns"]["program_name"]
    drive_col    = args.drive_col       or config["columns"].get("drive_link") or None
    fuzzy_thr    = args.fuzzy_threshold or config["scanning"]["fuzzy_threshold"]
    max_workers  = args.workers         or config["scanning"]["max_workers"]
    log_level    = args.log_level       or config["logging"]["level"]
    cache_file   = config["scanning"]["cache_file"]
    cache_age    = 0.0 if args.no_cache else float(config["scanning"]["cache_max_age_hours"])
    ext_filter   = config["scanning"].get("extensions_filter") or []
    gdrive_on    = (not args.no_gdrive) and config["google_drive"]["enabled"]
    creds_file   = config["google_drive"]["credentials_file"]
    settings_f   = config["google_drive"]["settings_file"]
    log_file     = config["logging"]["log_file"]

    # Reconfigure logger with final settings
    global logger
    logger = setup_logging(log_file, log_level)

    # Validate local drive paths
    valid_drives: list[str] = []
    for d in drives:
        p = Path(d)
        if p.exists():
            valid_drives.append(str(p.resolve()))
        else:
            logger.warning("Drive path does not exist, skipping: %s", d)

    # Log resolved configuration
    logger.info("=== Audit Checker starting ===")
    logger.info("Excel dir    : %s", excel_dir)
    logger.info("Valid drives : %s", valid_drives)
    logger.info("Output       : %s", output)
    logger.info("Program col  : %s", program_col)
    logger.info("Drive col    : %s", drive_col)
    logger.info("Fuzzy score  : %d", fuzzy_thr)
    logger.info("Max workers  : %d", max_workers)
    logger.info("GDrive on    : %s", gdrive_on)

    run_audit(
        excel_dir=excel_dir,
        drives=valid_drives,
        output=output,
        program_col=program_col,
        drive_col=drive_col,
        fuzzy_threshold=fuzzy_thr,
        cache_file=cache_file,
        cache_max_age_hours=cache_age,
        max_workers=max_workers,
        extensions_filter=ext_filter or None,
        gdrive_enabled=gdrive_on,
        credentials_file=creds_file,
        settings_file=settings_f,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Required on Windows so that spawned processes don't re-execute main()
    multiprocessing.freeze_support()
    main()
