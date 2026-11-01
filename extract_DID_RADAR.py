import re
import shutil
import sys
import threading
import json
import os
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
except ImportError:
    CellRichText = None
    TextBlock = None
    InlineFont = None
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


RESPONSE_ID = "18DAF1B0"

INPUT_EXTENSION = ".asc"

HEADER_KEYWORD_MAX = "max"
HEADER_KEYWORD_PDR = "pdr"
HEADER_KEYWORD_DESCRIPTION = "description"
HEADER_KEYWORD_SYMBOL = "symbol"
DATA_START_MARKER = "00"
SUMMARY_HEADER_COMPARE_RESULT = "Compare_result"
SUMMARY_HEADER_SUMMARY = "Summary"
SUMMARY_GAP_DEVIATION_HEADER = "Gap value deviations are limited to the following bytes; others are as expected:"
AVERAGE_GAP_HEADER = "Average gap_value"

TAG_OK = "[OK]"
TAG_NOK = "[NOK]"
TAG_ERROR = "[Error]"

MSG_NO_CUSTOMER_DATA = "Customer_response has no data"
MSG_NO_TESTER_DATA = "Tester_response has no data"
MSG_BOTH_GT_MAX = "Both Customer_response & Tester_response > Max"
MSG_CUSTOMER_GT_MAX = "Customer_response > Max"
MSG_TESTER_GT_MAX = "Tester_response > Max"
MSG_EQUAL = "Customer_response = Tester_response"
MSG_CUSTOMER_LT_TESTER = "Customer_response < Tester_response"
MSG_CUSTOMER_GT_TESTER = "Customer_response > Tester_response"

FORMULA_MSG_ERROR_NO_TESTER = f"{TAG_ERROR}: {MSG_NO_TESTER_DATA}"
FORMULA_MSG_ERROR_BOTH_GT_MAX = f"{TAG_ERROR}: {MSG_BOTH_GT_MAX}"
FORMULA_MSG_ERROR_CUSTOMER_GT_MAX = f"{TAG_ERROR}: {MSG_CUSTOMER_GT_MAX}"
FORMULA_MSG_ERROR_TESTER_GT_MAX = f"{TAG_ERROR}: {MSG_TESTER_GT_MAX}"
FORMULA_MSG_OK_EQUAL = f"{TAG_OK}: {MSG_EQUAL}"
FORMULA_MSG_NOK_CUSTOMER_LT = f"{TAG_NOK}: {MSG_CUSTOMER_LT_TESTER}"
FORMULA_MSG_NOK_CUSTOMER_GT = f"{TAG_NOK}: {MSG_CUSTOMER_GT_TESTER}"

# No-data messages written into Compare_result cells of DID sheets
FORMULA_MSG_BOTH_NO_DATA = "Error: Both customer_response & Tester_response have no data"
FORMULA_MSG_CUST_NO_DATA = "Error: Customer_response has no data"
FORMULA_MSG_TEST_NO_DATA = "Error: Tester_response has no data"

# Partial-data messages for Summary sheets (when data and no-data are mixed within the same response)
FORMULA_MSG_BOTH_PARTIAL_DATA = "Error: Both customer_response & Tester_response have partial data"
FORMULA_MSG_CUST_PARTIAL_DATA = "Error: Customer_response has partial data"
FORMULA_MSG_TEST_PARTIAL_DATA = "Error: Tester_response has partial data"

# Corresponding messages aggregated into Summary sheets
SUMMARY_MSG_BOTH_NO_DATA = FORMULA_MSG_BOTH_NO_DATA
SUMMARY_MSG_CUST_NO_DATA = FORMULA_MSG_CUST_NO_DATA
SUMMARY_MSG_TEST_NO_DATA = FORMULA_MSG_TEST_NO_DATA
SUMMARY_MSG_BOTH_PARTIAL_DATA = FORMULA_MSG_BOTH_PARTIAL_DATA
SUMMARY_MSG_CUST_PARTIAL_DATA = FORMULA_MSG_CUST_PARTIAL_DATA
SUMMARY_MSG_TEST_PARTIAL_DATA = FORMULA_MSG_TEST_PARTIAL_DATA

ANALYSIS_PROFILES = {
    "RADAR": {
        "mode": "RADAR",
        "customer_file_prefix": "Diag_customer_RADAR_",
        "tester_file_prefix": "Diag_tester_RADAR_",
        "logs_dir_name": "Diag_RADAR_logs",
        "response_dir_name": "DID_response_RADAR",
        "excel_template_name": "MaxHold_Re-sim_RADAR.xlsx",
        "results_dir_name": "MaxHold_Re-sim_RADAR_results",
        "output_excel_prefix": "MaxHold_Re-sim_RADAR_",
        "did_configs": [
            {
                "name": "A8C4",
                "target_prefix": bytes.fromhex("62 A8 C4"),
                "format_file": "Sample_format/A8C4_format.txt",
                "output_prefix": "A8C4_response_",
            },
            {
                "name": "A8C5",
                "target_prefix": bytes.fromhex("62 A8 C5"),
                "format_file": "Sample_format/A8C5_format.txt",
                "output_prefix": "A8C5_response_",
            },
        ],
    },
    "MPC": {
        "mode": "MPC",
        "customer_file_prefix": "Diag_customer_MPC_",
        "tester_file_prefix": "Diag_tester_MPC_",
        "logs_dir_name": "Diag_MPC_logs",
        "response_dir_name": "DID_response_MPC",
        "excel_template_name": "MaxHold_Re-sim_MPC.xlsx",
        "results_dir_name": "MaxHold_Re-sim_MPC_results",
        "output_excel_prefix": "MaxHold_Re-sim_MPC_",
        "did_configs": [
            {
                "name": "A8C6",
                "target_prefix": bytes.fromhex("62 A8 C6"),
                "format_file": "Sample_format/A8C6_format.txt",
                "output_prefix": "A8C6_response_",
            },
        ],
    },
}

DEFAULT_ANALYSIS_MODE = "RADAR"

TIMESTAMP_MATCH_TOLERANCE_SEC = 0.020
GAP_VALUE_PATTERN = re.compile(r"gap\s*value\s*=\s*(-?\d+)", re.IGNORECASE)
GUI_SETTINGS_FILE = ".did_analysis_gui_settings.json"


def detect_analysis_mode(base_dir: Path):
    """Detect RADAR/MPC mode from folder naming and available files."""
    dir_name_upper = base_dir.name.upper()
    if "MPC" in dir_name_upper and "RADAR" not in dir_name_upper:
        return "MPC"
    if "RADAR" in dir_name_upper and "MPC" not in dir_name_upper:
        return "RADAR"

    sample_dir = base_dir / "Sample_format"
    radar_hint = (
        (base_dir / "Diag_RADAR_logs").exists()
        or any(base_dir.glob("Diag_customer_RADAR_*.asc"))
        or any(base_dir.glob("Diag_tester_RADAR_*.asc"))
        or (sample_dir / "MaxHold_Re-sim_RADAR.xlsx").exists()
        or (sample_dir / "A8C4_format.txt").exists()
        or (sample_dir / "A8C5_format.txt").exists()
    )
    mpc_hint = (
        (base_dir / "Diag_MPC_logs").exists()
        or any(base_dir.glob("Diag_customer_MPC_*.asc"))
        or any(base_dir.glob("Diag_tester_MPC_*.asc"))
        or (sample_dir / "MaxHold_Re-sim_MPC.xlsx").exists()
        or (sample_dir / "A8C6_format.txt").exists()
    )

    if mpc_hint and not radar_hint:
        return "MPC"
    if radar_hint and not mpc_hint:
        return "RADAR"
    return DEFAULT_ANALYSIS_MODE


def get_analysis_profile(base_dir: Path):
    mode = detect_analysis_mode(base_dir)
    return ANALYSIS_PROFILES.get(mode, ANALYSIS_PROFILES[DEFAULT_ANALYSIS_MODE])


def extract_data_bytes(line: str):
    hex_pairs = re.findall(r"\b[0-9A-Fa-f]{2}\b", line)
    if len(hex_pairs) < 8:
        return None
    return [int(x, 16) for x in hex_pairs[-8:]]


def normalize_can_id(line: str):
    match = re.search(r"\b([0-9A-Fa-f]{8})x?\b", line)
    if not match:
        return None
    return match.group(1).upper()


def extract_line_timestamp(line: str):
    """Extract ASC timestamp at the beginning of a log line."""
    match = re.match(r"\s*([0-9]+(?:\.[0-9]+)?)\b", line)
    if not match:
        return None
    return match.group(1)


def complete_message_if_ready(buffer: bytearray, total_len: int):
    if total_len is None:
        return None
    if len(buffer) < total_len:
        return None
    return bytes(buffer[:total_len])


def parse_format_file(format_path: Path):
    ranges = []
    with format_path.open("r", encoding="utf-8", errors="ignore") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line:
                continue

            if "-" in line:
                left, right = [part.strip() for part in line.split("-", 1)]
                start = int(left, 16)
                end = int(right, 16)
                if end < start:
                    raise ValueError(f"Invalid format: {line}")
                ranges.append((start, end))
            else:
                idx = int(line, 16)
                ranges.append((idx, idx))
    return ranges


def build_output_lines(payload: bytes, ranges):
    output_lines = []
    for start, end in ranges:
        if end >= len(payload):
            output_lines.append(None)  # Byte range not available in partial payload
            continue
        chunk = payload[start : end + 1]
        output_lines.append("".join(f"{b:02X}" for b in chunk))
    return output_lines


def parse_timestamp_seconds(timestamp_text):
    if timestamp_text is None:
        return None
    try:
        return float(timestamp_text)
    except (TypeError, ValueError):
        return None


def natural_sort_key(text: str):
    """Return a key that sorts strings in human/natural order (e.g. #2 before #10)."""
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def align_responses_by_timestamp(customer_entries, tester_entries, tolerance_sec=TIMESTAMP_MATCH_TOLERANCE_SEC):
    """Align customer/tester responses by first-frame timestamp within tolerance."""
    aligned = []
    i = 0
    j = 0

    while i < len(customer_entries) and j < len(tester_entries):
        customer_entry = customer_entries[i]
        tester_entry = tester_entries[j]

        cust_ts = parse_timestamp_seconds(customer_entry.get("first_frame_timing"))
        test_ts = parse_timestamp_seconds(tester_entry.get("first_frame_timing"))

        if cust_ts is None and test_ts is None:
            aligned.append({"customer": customer_entry, "tester": tester_entry})
            i += 1
            j += 1
            continue

        if cust_ts is None:
            aligned.append({"customer": customer_entry, "tester": None})
            i += 1
            continue

        if test_ts is None:
            aligned.append({"customer": None, "tester": tester_entry})
            j += 1
            continue

        if abs(cust_ts - test_ts) <= tolerance_sec:
            aligned.append({"customer": customer_entry, "tester": tester_entry})
            i += 1
            j += 1
        elif cust_ts < test_ts:
            aligned.append({"customer": customer_entry, "tester": None})
            i += 1
        else:
            aligned.append({"customer": None, "tester": tester_entry})
            j += 1

    while i < len(customer_entries):
        aligned.append({"customer": customer_entries[i], "tester": None})
        i += 1

    while j < len(tester_entries):
        aligned.append({"customer": None, "tester": tester_entries[j]})
        j += 1

    return aligned


def extract_payloads_for_prefix(in_path: Path, target_prefix: bytes):
    active_total_len = None
    active_buffer = bytearray()
    active_is_target = False  # Whether the in-progress multi-frame message starts with target_prefix
    active_first_frame_timing = None
    extracted_payloads = []

    with in_path.open("r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            can_id = normalize_can_id(line)
            if can_id != RESPONSE_ID:
                continue

            timestamp = extract_line_timestamp(line)

            data = extract_data_bytes(line)
            if data is None:
                continue

            pci = data[0]
            frame_type = pci >> 4

            if frame_type == 0x0:
                sf_len = pci & 0x0F
                message = bytes(data[1 : 1 + sf_len])
                if message.startswith(target_prefix):
                    extracted_payloads.append(
                        {
                            "payload": message[len(target_prefix) :],
                            "first_frame_timing": timestamp,
                        }
                    )

            elif frame_type == 0x1:
                # If a previous in-progress message was for our target and never completed, record partial bytes.
                if active_total_len is not None and active_is_target:
                    extracted_payloads.append(
                        {
                            "payload": bytes(active_buffer)[len(target_prefix):],
                            "first_frame_timing": active_first_frame_timing,
                        }
                    )

                active_total_len = ((pci & 0x0F) << 8) | data[1]
                active_buffer = bytearray(data[2:])
                active_is_target = bytes(active_buffer).startswith(target_prefix)
                active_first_frame_timing = timestamp if active_is_target else None

                message = complete_message_if_ready(active_buffer, active_total_len)
                if message and message.startswith(target_prefix):
                    extracted_payloads.append(
                        {
                            "payload": message[len(target_prefix) :],
                            "first_frame_timing": active_first_frame_timing,
                        }
                    )
                    active_total_len = None
                    active_buffer = bytearray()
                    active_is_target = False
                    active_first_frame_timing = None

            elif frame_type == 0x2:
                if active_total_len is None:
                    continue
                active_buffer.extend(data[1:])
                message = complete_message_if_ready(active_buffer, active_total_len)
                if message and message.startswith(target_prefix):
                    extracted_payloads.append(
                        {
                            "payload": message[len(target_prefix) :],
                            "first_frame_timing": active_first_frame_timing,
                        }
                    )
                if message:
                    active_total_len = None
                    active_buffer = bytearray()
                    active_is_target = False
                    active_first_frame_timing = None

    # Handle incomplete target message at end of file.
    if active_total_len is not None and active_is_target:
        extracted_payloads.append(
            {
                "payload": bytes(active_buffer)[len(target_prefix):],
                "first_frame_timing": active_first_frame_timing,
            }
        )

    return extracted_payloads


def write_response_files(output_dir: Path, file_prefix: str, payloads, format_ranges):
    output_dir.mkdir(parents=True, exist_ok=True)

    for idx, response_entry in enumerate(payloads, start=1):
        payload = response_entry["payload"]
        first_frame_timing = response_entry.get("first_frame_timing")
        output_lines = build_output_lines(payload, format_ranges)
        out_path = output_dir / f"{file_prefix}{idx}.txt"

        with out_path.open("w", encoding="ascii") as f:
            timing_text = first_frame_timing if first_frame_timing is not None else "N/A"
            f.write(f"First-frame timestamp: {timing_text}\n")
            for line in output_lines:
                f.write(f"{line}\n" if line is not None else "\n")


def resolve_input_path(base_dir: Path, candidates):
    for filename in candidates:
        path = base_dir / filename
        if path.exists():
            return path
    return None


def discover_trace_pairs(base_dir: Path, profile=None):
    """Find all customer .asc files and all matching tester .asc files by prefix."""
    profile = profile or get_analysis_profile(base_dir)
    customer_prefix = profile["customer_file_prefix"]
    tester_prefix_root = profile["tester_file_prefix"]
    logs_dir = base_dir / profile["logs_dir_name"]
    search_dir = logs_dir if logs_dir.exists() else base_dir
    customer_files = []
    tester_files = []

    for path in search_dir.glob(f"{customer_prefix}*{INPUT_EXTENSION}"):
        customer_files.append(path)

    for path in search_dir.glob(f"{tester_prefix_root}*{INPUT_EXTENSION}"):
        tester_files.append(path)

    # Create pairs: each customer may have multiple testers with matching prefix
    pairs = []
    for customer_path in sorted(customer_files):
        customer_suffix = customer_path.stem[len(customer_prefix) :]
        tester_prefix = f"{tester_prefix_root}{customer_suffix}"
        
        # Find all tester files that start with this prefix followed by '_' or end of string.
        # This prevents e.g. customer '#1' from matching tester '#10_1_N1_run_1'.
        matching_testers = [
            t for t in sorted(tester_files)
            if t.stem.startswith(tester_prefix)
            and (len(t.stem) == len(tester_prefix) or t.stem[len(tester_prefix)] == "_")
        ]
        
        # Create one pair per matching tester, plus one without tester if none found
        if matching_testers:
            for tester_path in matching_testers:
                tester_suffix = tester_path.stem[len(tester_prefix_root) :]
                pairs.append({
                    "customer_suffix": customer_suffix,
                    "customer": customer_path,
                    "tester": tester_path,
                    "tester_suffix": tester_suffix,
                })
        else:
            pairs.append({
                "customer_suffix": customer_suffix,
                "customer": customer_path,
                "tester": None,
                "tester_suffix": None,
            })

    return pairs


def cleanup_response_folders(base_dir: Path, profile=None):
    """Delete entire DID_response_* folder with all subfolders."""
    profile = profile or get_analysis_profile(base_dir)
    response_root = base_dir / profile["response_dir_name"]
    if not response_root.exists():
        return

    try:
        shutil.rmtree(response_root)
        print(f"Cleanup: Deleted {profile['response_dir_name']} folder")
    except Exception as e:
        print(f"  Warning: Failed to delete {profile['response_dir_name']} folder: {e}")


def cleanup_results_folder(base_dir: Path, profile=None):
    """Delete mode-specific Excel results folder and all files inside it."""
    profile = profile or get_analysis_profile(base_dir)
    results_root = base_dir / profile["results_dir_name"]
    if not results_root.exists():
        return

    try:
        shutil.rmtree(results_root)
        print(f"Cleanup: Deleted {profile['results_dir_name']} folder")
    except Exception as e:
        print(f"  Warning: Failed to delete {profile['results_dir_name']} folder: {e}")


def clone_sample_excel(base_dir: Path, output_excel_path: Path, profile):
    """Clone sample Excel template from Sample/ to output_excel_path."""
    sample_path = base_dir / "Sample_format" / profile["excel_template_name"]
    dest_path = output_excel_path

    if not sample_path.exists():
        print(f"[Excel] Error: Sample file not found: {sample_path}")
        return False

    shutil.copy2(sample_path, dest_path)
    print(
        f"Cloned sample: Sample_format/{profile['excel_template_name']}"
        f" -> {profile['results_dir_name']}/{dest_path.name}"
    )
    return True


def shorten_path_for_display(path: Path, max_len: int = 85) -> str:
    text = str(path)
    if len(text) <= max_len:
        return text
    head_len = max_len // 2 - 2
    tail_len = max_len - head_len - 3
    return f"{text[:head_len]}...{text[-tail_len:]}"


def get_gui_settings_path(base_dir: Path) -> Path:
    return base_dir / GUI_SETTINGS_FILE


def load_saved_base_dir(default_base_dir: Path) -> Path:
    settings_path = get_gui_settings_path(default_base_dir)
    if not settings_path.exists():
        return default_base_dir

    try:
        with settings_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        saved = data.get("last_selected_base_dir")
        if isinstance(saved, str) and saved.strip():
            saved_path = Path(saved)
            if saved_path.exists() and saved_path.is_dir():
                return saved_path
    except Exception:
        pass

    return default_base_dir


def save_selected_base_dir(default_base_dir: Path, selected_base_dir: Path):
    settings_path = get_gui_settings_path(default_base_dir)
    payload = {"last_selected_base_dir": str(selected_base_dir)}
    try:
        with settings_path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except Exception:
        # Settings persistence failure must not block analysis flow.
        pass


def resolve_runtime_base_dir() -> Path:
    """Return runtime base directory.

    - EXE run (PyInstaller): folder containing the executable.
    - Python run: folder containing this script.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_col_by_header_contains(ws, keyword: str):
    """Find first column whose header contains keyword (case-insensitive)."""
    key = keyword.strip().lower()
    for cell in ws[1]:
        if not cell.value:
            continue
        header = str(cell.value).strip().lower().replace("\n", " ")
        if key in header:
            return cell.column
    return None


def find_Max_col_and_data_start(ws):
    """Find 'Max' column and first row where PDR column = '00'."""
    Max_col = find_col_by_header_contains(ws, HEADER_KEYWORD_MAX)
    pdr_col = find_col_by_header_contains(ws, HEADER_KEYWORD_PDR)

    if pdr_col is None:
        return Max_col, None

    data_start_row = None
    for row in ws.iter_rows(min_col=pdr_col, max_col=pdr_col):
        cell = row[0]
        if cell.value is not None and str(cell.value).strip() == DATA_START_MARKER:
            data_start_row = cell.row
            break

    return Max_col, data_start_row


def find_max_col(ws):
    """Find the column with header 'Max' (case-insensitive)."""
    return find_col_by_header_contains(ws, HEADER_KEYWORD_MAX)


def parse_max_from_description(desc_value):
    if desc_value is None:
        return None
    match = re.search(r"Max:\s*(-?\d+(?:\.\d+)?)", str(desc_value))
    if not match:
        return None
    value = float(match.group(1))
    if value.is_integer():
        return int(value)
    return value


def build_row_context(ws, data_start_row):
    """Build per-row context (PDR and Max) starting from data_start_row."""
    pdr_col = find_col_by_header_contains(ws, HEADER_KEYWORD_PDR)
    desc_col = find_col_by_header_contains(ws, HEADER_KEYWORD_DESCRIPTION)
    max_col = find_max_col(ws)
    symbol_col = find_col_by_header_contains(ws, HEADER_KEYWORD_SYMBOL)

    if pdr_col is None:
        return []

    max_rows = ws.max_row - data_start_row + 1
    contexts = []

    for i in range(max_rows):
        row = data_start_row + i
        pdr_val = ws.cell(row=row, column=pdr_col).value
        if pdr_val is None:
            contexts.append(None)
            continue

        pdr_text = str(pdr_val).strip()
        max_val = None
        if max_col is not None:
            raw_max = ws.cell(row=row, column=max_col).value
            if isinstance(raw_max, (int, float)):
                max_val = raw_max
            elif raw_max is not None:
                raw_text = str(raw_max).strip()
                # Ignore formulas here; unresolved formula text can contain row numbers (e.g. C36)
                # and cause incorrect Max parsing in Summary.
                if not raw_text.startswith("="):
                    m = re.fullmatch(r"-?\d+(?:\.\d+)?", raw_text)
                    if m:
                        parsed = float(m.group(0))
                        max_val = int(parsed) if parsed.is_integer() else parsed

        # Backward-compatible fallback for old templates where Max is embedded in description text.
        if max_val is None and desc_col is not None:
            max_val = parse_max_from_description(ws.cell(row=row, column=desc_col).value)
        symbol_val = None
        if symbol_col is not None:
            symbol_cell = ws.cell(row=row, column=symbol_col).value
            if symbol_cell is not None:
                symbol_val = str(symbol_cell).strip()
        contexts.append({"pdr": pdr_text, "max": max_val, "symbol": symbol_val})

    return contexts


def build_summary_messages_for_response(idx, did_name, excel_data, row_contexts, source_ws):
    """Aggregate Compare_result messages for one response index in one DID sheet.
    
    Priority order:
    1. Both no-data (all rows None in both customer & tester)
    2. Both partial-data (mixed data and None in both)
    3. Customer no-data
    4. Customer partial-data
    5. Tester no-data
    6. Tester partial-data
    7. Per-row comparison
    """
    messages = []
    found_issue = False

    customer_payloads = excel_data.get(did_name, {}).get("customer", [])
    tester_payloads = excel_data.get(did_name, {}).get("tester", [])
    contexts = row_contexts.get(did_name, [])

    customer_payload = customer_payloads[idx] if idx < len(customer_payloads) else None
    tester_payload = tester_payloads[idx] if idx < len(tester_payloads) else None

    # Determine data state for each side: "absent", "no_data", "partial", or "complete"
    def get_data_state(payload):
        if payload is None:
            return "absent"
        if all(line is None for line in payload):
            return "no_data"
        if any(line is None for line in payload):
            return "partial"
        return "complete"

    cust_state = get_data_state(customer_payload)
    test_state = get_data_state(tester_payload)

    # Priority-based error checking
    if cust_state in ("absent", "no_data") and test_state in ("absent", "no_data"):
        return SUMMARY_MSG_BOTH_NO_DATA
    if cust_state == "partial" and test_state == "partial":
        # Compare available data for both customer and tester with their previous responses
        if idx > 0 and idx - 1 < len(customer_payloads) and idx - 1 < len(tester_payloads):
            prev_cust_payload = customer_payloads[idx - 1]
            prev_test_payload = tester_payloads[idx - 1]
            
            # Check customer match
            cust_all_match = True
            for row_idx, current_line in enumerate(customer_payload):
                if current_line is not None:
                    prev_line = prev_cust_payload[row_idx] if row_idx < len(prev_cust_payload) else None
                    if current_line != prev_line:
                        cust_all_match = False
                        break
            
            # Check tester match
            test_all_match = True
            for row_idx, current_line in enumerate(tester_payload):
                if current_line is not None:
                    prev_line = prev_test_payload[row_idx] if row_idx < len(prev_test_payload) else None
                    if current_line != prev_line:
                        test_all_match = False
                        break
            
            # Return appropriate message based on match status
            if cust_all_match and test_all_match:
                return f"Error: Both customer_response & Tester_response have partial data & The available data from both are the same as in Customer_response_{idx} and Tester_response_{idx}"
            elif not cust_all_match and not test_all_match:
                return f"Error: Both customer_response & Tester_response have partial data & The available data from both are ***NOT*** the same as in Customer_response_{idx} and Tester_response_{idx}"
            elif not cust_all_match:
                return f"Error: Both customer_response & Tester_response have partial data & Customer available data is ***NOT*** the same as in Customer_response_{idx}"
            else:
                return f"Error: Both customer_response & Tester_response have partial data & Tester available data is ***NOT*** the same as in Tester_response_{idx}"
        return SUMMARY_MSG_BOTH_PARTIAL_DATA
    if cust_state in ("absent", "no_data"):
        return SUMMARY_MSG_CUST_NO_DATA
    if cust_state == "partial":
        # Compare available data with previous customer response
        if idx > 0 and idx - 1 < len(customer_payloads):
            prev_payload = customer_payloads[idx - 1]
            all_match = True
            for row_idx, current_line in enumerate(customer_payload):
                if current_line is not None:  # Only compare available rows
                    prev_line = prev_payload[row_idx] if row_idx < len(prev_payload) else None
                    if current_line != prev_line:
                        all_match = False
                        break
            if all_match:
                return f"Error: Customer_response has partial data & The available customer data is the same as in Customer_response_{idx}"
            else:
                return f"Error: Customer_response has partial data & The available customer data is ***NOT*** the same as in Customer_response_{idx}"
        return SUMMARY_MSG_CUST_PARTIAL_DATA
    if test_state in ("absent", "no_data"):
        return SUMMARY_MSG_TEST_NO_DATA
    if test_state == "partial":
        return SUMMARY_MSG_TEST_PARTIAL_DATA

    def clean_symbol_name(symbol_name):
        """Normalize legacy symbol prefixes for cleaner summary display."""
        if not symbol_name:
            return symbol_name
        cleaned = str(symbol_name).strip()
        while cleaned.startswith("FaileSafe value "):
            cleaned = cleaned[len("FaileSafe value "):]
        if cleaned.startswith("FS_CNT_MAX_"):
            cleaned = cleaned[len("FS_CNT_MAX_"):]
        return cleaned

    # Both have complete data → proceed with row-by-row comparison
    Max_col, data_start_row = find_Max_col_and_data_start(source_ws)
    gap_col = None if Max_col is None else Max_col + 3 + idx * 4

    customer_lines = customer_payload
    tester_lines = tester_payload
    row_count = min(len(customer_lines), len(tester_lines), len(contexts))

    for row_idx in range(row_count):
        ctx = contexts[row_idx]
        if not ctx or ctx["max"] is None:
            continue

        cust_val_str = customer_lines[row_idx] if row_idx < len(customer_lines) else None
        test_val_str = tester_lines[row_idx] if row_idx < len(tester_lines) else None
        if cust_val_str is None or test_val_str is None:
            continue  # Missing byte range — skip row comparison
        customer_val = int(cust_val_str, 16)
        tester_val = int(test_val_str, 16)
        max_val = ctx["max"]
        pdr = ctx["pdr"]
        symbol = ctx.get("symbol")
        pdr_prefix = f"PDR {pdr}"
        if symbol:
            pdr_prefix = f"PDR {pdr} ({clean_symbol_name(symbol)})"

        if customer_val > max_val and tester_val > max_val:
            found_issue = True
            messages.append(f"{pdr_prefix}: {MSG_BOTH_GT_MAX}")
        elif customer_val > max_val:
            found_issue = True
            messages.append(f"{pdr_prefix}: {MSG_CUSTOMER_GT_MAX}")
        elif tester_val > max_val:
            found_issue = True
            messages.append(f"{pdr_prefix}: {MSG_TESTER_GT_MAX}")
        elif customer_val < tester_val:
            found_issue = True
            gap_val = abs(customer_val - tester_val)
            if gap_col is not None and data_start_row is not None:
                source_ws.cell(row=data_start_row + row_idx, column=gap_col)
            messages.append(
                f"{pdr_prefix}: {MSG_CUSTOMER_LT_TESTER} with gap value = {gap_val}"
            )
        elif customer_val > tester_val:
            found_issue = True
            gap_val = abs(customer_val - tester_val)
            if gap_col is not None and data_start_row is not None:
                source_ws.cell(row=data_start_row + row_idx, column=gap_col)
            messages.append(
                f"{pdr_prefix}: {MSG_CUSTOMER_GT_TESTER} with gap value = {gap_val}"
            )

    if not found_issue:
        return MSG_EQUAL
    return f"{SUMMARY_GAP_DEVIATION_HEADER}\n" + "\n".join(messages)


def write_summary_sheets(wb, excel_data, did_configs):
    summary_sources = {f"{cfg['name']}_Summary": cfg["name"] for cfg in did_configs}

    # Build row contexts once from source sheets.
    row_contexts = {}
    for did_name in [cfg["name"] for cfg in did_configs]:
        if did_name not in wb.sheetnames:
            row_contexts[did_name] = []
            continue
        ws = wb[did_name]
        _, data_start_row = find_Max_col_and_data_start(ws)
        if data_start_row is None:
            row_contexts[did_name] = []
            continue
        row_contexts[did_name] = build_row_context(ws, data_start_row)

    for summary_name, did_name in summary_sources.items():
        # Delete old summary sheet if exists, then recreate next to source sheet.
        if summary_name in wb.sheetnames:
            del wb[summary_name]

        source_idx = wb.sheetnames.index(did_name) if did_name in wb.sheetnames else None
        insert_pos = (source_idx + 1) if source_idx is not None else None
        ws = wb.create_sheet(summary_name, insert_pos)

        max_responses = max(
            len(excel_data.get(did_name, {}).get("customer", [])),
            len(excel_data.get(did_name, {}).get("tester", [])),
        )

        ws.cell(row=1, column=1, value=SUMMARY_HEADER_COMPARE_RESULT)
        ws.cell(row=1, column=2, value=SUMMARY_HEADER_SUMMARY)
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=2).font = Font(bold=True)

        max_summary_line_len = len(SUMMARY_HEADER_SUMMARY)

        for idx in range(max_responses):
            out_row = idx + 2
            ws.cell(row=out_row, column=1, value=f"Compare_result_{idx + 1}")
            summary_text = build_summary_messages_for_response(
                idx, did_name, excel_data, row_contexts, wb[did_name]
            )
            cell_b = ws.cell(row=out_row, column=2, value=summary_text)
            
            # Handle rich text formatting for ***NOT*** marker (bold + red)
            if CellRichText is not None and isinstance(summary_text, str) and "***NOT***" in summary_text:
                parts = summary_text.split("***NOT***")
                if len(parts) == 2:
                    rich_text = CellRichText(
                        parts[0],
                        TextBlock(InlineFont(b=True, color="FF0000"), "NOT"),
                        parts[1]
                    )
                    cell_b.value = rich_text
            
            cell_b.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            # Track longest line in Summary text to size column width dynamically.
            for line in str(summary_text).splitlines():
                max_summary_line_len = max(max_summary_line_len, len(line))

        ws.column_dimensions[get_column_letter(1)].width = max(20, len("Compare_result_99"))
        ws.column_dimensions[get_column_letter(2)].width = min(max_summary_line_len + 2, 255)

        # Conditional formatting on Summary column B: red background for error messages
        if max_responses > 0:
            summary_range = f"B2:B{max_responses + 1}"
            ws.conditional_formatting.add(
                summary_range,
                Rule(
                    type="containsText",
                    operator="containsText",
                    text="Error",
                    formula=[f'NOT(ISERROR(SEARCH("Error",B2)))'],
                    dxf=DifferentialStyle(
                        fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    ),
                ),
            )


def write_responses_to_excel(wb, sheet_name, base_dir, all_payloads_by_source, format_ranges):
    """
    all_payloads_by_source: dict {"customer": [payload,...], "tester": [payload,...]}
    Each payload is a list of hex string lines (output_lines).
    """
    ws = wb[sheet_name]
    Max_col, data_start_row = find_Max_col_and_data_start(ws)
    if Max_col is None or data_start_row is None:
        print(f"  [{sheet_name}] Cannot find Max column or row '00', skipping Excel.")
        return False

    max_col = find_max_col(ws)
    if max_col is None:
        print(f"  [{sheet_name}] Cannot find Max column, Compare_result formulas will be skipped.")

    # Responses are pre-aligned by timestamp (customer/tester can be missing in each slot).
    customer_payloads = all_payloads_by_source.get("customer", [])
    tester_payloads = all_payloads_by_source.get("tester", [])
    num_responses = max(len(customer_payloads), len(tester_payloads))

    # Add Average gap header at Max_col + 1
    avg_gap_col = Max_col + 1
    header_row = data_start_row - 1
    avg_gap_header = ws.cell(row=header_row, column=avg_gap_col, value=AVERAGE_GAP_HEADER)
    avg_gap_header.font = Font(bold=True)
    avg_gap_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(avg_gap_col)].width = len(AVERAGE_GAP_HEADER) + 2

    for idx in range(num_responses):
        # Each response takes 4 columns: customer_response_X, tester_response_X, gap_value_X, Compare_result_X
        # Average_gap_value occupies Max_col+1, so responses start at Max_col+2
        cust_col = Max_col + 2 + idx * 4
        test_col = Max_col + 3 + idx * 4
        gap_col = Max_col + 4 + idx * 4
        compare_col = Max_col + 5 + idx * 4

        header_row = data_start_row - 1  # header row (row 1 in sheet)
        cust_header = ws.cell(row=header_row, column=cust_col, value=f"Customer_response_{idx+1}")
        test_header = ws.cell(row=header_row, column=test_col, value=f"Tester_response_{idx+1}")
        gap_header = ws.cell(row=header_row, column=gap_col, value=f"Gap_value_{idx+1}")
        compare_header = ws.cell(row=header_row, column=compare_col, value=f"Compare_result_{idx+1}")
        
        # Apply font and alignment to all four headers
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for header_cell in [cust_header, test_header, gap_header, compare_header]:
            header_cell.font = Font(bold=True)
            header_cell.alignment = center_align

        cust_col_letter = get_column_letter(cust_col)
        test_col_letter = get_column_letter(test_col)
        gap_col_letter = get_column_letter(gap_col)
        compare_col_letter = get_column_letter(compare_col)
        
        ws.column_dimensions[cust_col_letter].width = max(
            ws.column_dimensions[cust_col_letter].width or 0,
            len(str(cust_header.value)) + 2,
        )
        ws.column_dimensions[test_col_letter].width = max(
            ws.column_dimensions[test_col_letter].width or 0,
            len(str(test_header.value)) + 2,
        )
        ws.column_dimensions[gap_col_letter].width = max(
            ws.column_dimensions[gap_col_letter].width or 0,
            len(str(gap_header.value)) + 2,
        )
        ws.column_dimensions[compare_col_letter].width = len(FORMULA_MSG_ERROR_BOTH_GT_MAX)

        center_mid = Alignment(horizontal="center", vertical="center")

        # Write customer data
        if idx < len(customer_payloads):
            output_lines = customer_payloads[idx]
            if output_lines is not None:
                for line_idx, hex_val in enumerate(output_lines):
                    if hex_val is None:
                        continue  # Missing byte range — leave cell blank
                    c = ws.cell(
                        row=data_start_row + line_idx,
                        column=cust_col,
                        value=f'=HEX2DEC("{hex_val}")',
                    )
                    c.alignment = center_mid

        # Write tester data
        if idx < len(tester_payloads):
            output_lines = tester_payloads[idx]
            if output_lines is not None:
                for line_idx, hex_val in enumerate(output_lines):
                    if hex_val is None:
                        continue  # Missing byte range — leave cell blank
                    c = ws.cell(
                        row=data_start_row + line_idx,
                        column=test_col,
                        value=f'=HEX2DEC("{hex_val}")',
                    )
                    c.alignment = center_mid

        # Write Gap_value formulas; IF guard returns "" when either side is missing
        num_data_rows = len(format_ranges)
        cust_letter = get_column_letter(cust_col)
        test_letter = get_column_letter(test_col)
        for line_idx in range(num_data_rows):
            row = data_start_row + line_idx
            formula = (
                f'=IF(OR(ISBLANK({cust_letter}{row}),ISBLANK({test_letter}{row})),"",'
                f'ABS({cust_letter}{row}-{test_letter}{row}))'
            )
            c = ws.cell(row=row, column=gap_col, value=formula)
            c.alignment = center_mid

        # Write Compare_result formulas
        if max_col is not None:
            max_letter = get_column_letter(max_col)
            for line_idx in range(num_data_rows):
                row = data_start_row + line_idx
                formula = (
                    f'=IF(AND(OR(ISBLANK({cust_letter}{row}),{cust_letter}{row}=""),OR(ISBLANK({test_letter}{row}),{test_letter}{row}="")),'
                    f'"{FORMULA_MSG_BOTH_NO_DATA}",'
                    f'IF(OR(ISBLANK({cust_letter}{row}),{cust_letter}{row}=""),'
                    f'"{FORMULA_MSG_CUST_NO_DATA}",'
                    f'IF(OR(ISBLANK({test_letter}{row}),{test_letter}{row}=""),'
                    f'"{FORMULA_MSG_TEST_NO_DATA}",'
                    f'IF(AND({cust_letter}{row}>{max_letter}{row},{test_letter}{row}>{max_letter}{row}),'
                    f'"{FORMULA_MSG_ERROR_BOTH_GT_MAX}",'
                    f'IF({cust_letter}{row}>{max_letter}{row},'
                    f'"{FORMULA_MSG_ERROR_CUSTOMER_GT_MAX}",'
                    f'IF({test_letter}{row}>{max_letter}{row},'
                    f'"{FORMULA_MSG_ERROR_TESTER_GT_MAX}",'
                    f'IF({cust_letter}{row}={test_letter}{row},'
                    f'"{FORMULA_MSG_OK_EQUAL}",'
                    f'IF({cust_letter}{row}<{test_letter}{row},'
                    f'"{FORMULA_MSG_NOK_CUSTOMER_LT}",'
                    f'"{FORMULA_MSG_NOK_CUSTOMER_GT}"))))))))'
                )
                c = ws.cell(row=row, column=compare_col, value=formula)
                c.alignment = center_mid

            # Apply conditional formatting to the Compare_result column.
            # containsText rule requires an explicit formula to work correctly in Excel.
            compare_range = (
                f"{compare_col_letter}{data_start_row}"
                f":{compare_col_letter}{data_start_row + num_data_rows - 1}"
            )
            # The formula anchor must reference the first cell of the range.
            first_cell = f"{compare_col_letter}{data_start_row}"

            # [OK] → green background
            ws.conditional_formatting.add(
                compare_range,
                Rule(
                    type="containsText",
                    operator="containsText",
                    text=TAG_OK,
                    formula=[f'NOT(ISERROR(SEARCH("{TAG_OK}",{first_cell})))'],
                    dxf=DifferentialStyle(
                        fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    ),
                ),
            )
            # [NOK] → orange background
            ws.conditional_formatting.add(
                compare_range,
                Rule(
                    type="containsText",
                    operator="containsText",
                    text=TAG_NOK,
                    formula=[f'NOT(ISERROR(SEARCH("{TAG_NOK}",{first_cell})))'],
                    dxf=DifferentialStyle(
                        fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                    ),
                ),
            )
            # [Error] and Error: → red background
            ws.conditional_formatting.add(
                compare_range,
                Rule(
                    type="containsText",
                    operator="containsText",
                    text="Error",
                    formula=[f'NOT(ISERROR(SEARCH("Error",{first_cell})))'],
                    dxf=DifferentialStyle(
                        fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    ),
                ),
            )

    # Calculate Average gap_value for each data row
    if num_responses > 0:
        num_data_rows = len(format_ranges)
        
        for row_idx in range(num_data_rows):
            row = data_start_row + row_idx
            # Build AVERAGE formula referencing all gap_value columns for this row
            gap_cols_refs = []
            for idx in range(num_responses):
                gap_col_pos = Max_col + 4 + idx * 4
                gap_cols_refs.append(f"{get_column_letter(gap_col_pos)}{row}")
            
            if gap_cols_refs:
                formula = f"=IFERROR(AVERAGE({','.join(gap_cols_refs)}),\"\")"
                c = ws.cell(row=row, column=avg_gap_col, value=formula)
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Apply AutoFilter to all columns from A to last data column
    if num_responses > 0 and num_data_rows > 0:
        # Calculate the last column with data
        last_response_idx = num_responses - 1
        last_col = Max_col + 5 + last_response_idx * 4  # Compare_result column of last response
        last_col_letter = get_column_letter(last_col)
        
        # Header row to last data row
        header_row = data_start_row - 1
        last_data_row = data_start_row + num_data_rows - 1
        
        # Set AutoFilter range: from column A header to last data column
        filter_range = f"A{header_row}:{last_col_letter}{last_data_row}"
        ws.auto_filter.ref = filter_range

    return True


def has_no_data_issue(excel_data, did_configs):
    """Return True if any DID response triggers one of the 3 no-data errors:
    - Both customer_response & Tester_response have no data
    - Customer_response has no data
    - Tester_response have no data
    """
    for cfg in did_configs:
        did_name = cfg["name"]
        customer_payloads = excel_data.get(did_name, {}).get("customer", [])
        tester_payloads = excel_data.get(did_name, {}).get("tester", [])
        num_responses = max(len(customer_payloads), len(tester_payloads))

        for idx in range(num_responses):
            # Customer absent or partial
            if idx >= len(customer_payloads):
                return True
            if customer_payloads[idx] is None:
                return True
            if any(line is None for line in customer_payloads[idx]):
                return True
            # Tester absent or partial
            if idx >= len(tester_payloads):
                return True
            if tester_payloads[idx] is None:
                return True
            if any(line is None for line in tester_payloads[idx]):
                return True
    return False


def get_pair_output_excel_name(pair, profile):
    tester_suffix = pair.get("tester_suffix")
    customer_suffix = pair["customer_suffix"]
    if tester_suffix is not None:
        return f"{profile['output_excel_prefix']}{tester_suffix}.xlsx"
    return f"{profile['output_excel_prefix']}{customer_suffix}.xlsx"


def get_pair_label(pair):
    tester_suffix = pair.get("tester_suffix")
    customer_suffix = pair["customer_suffix"]
    if tester_suffix is not None:
        return f"{customer_suffix}+{tester_suffix}"
    return customer_suffix


def get_pair_result_status(base_dir: Path, pair, profile):
    output_excel_name = get_pair_output_excel_name(pair, profile)
    results_dir = base_dir / profile["results_dir_name"]
    ok_path = results_dir / output_excel_name
    no_data_path = results_dir / "No-data_issue" / output_excel_name

    if ok_path.exists():
        return "OK (available)"
    if no_data_path.exists():
        return "No-data issue"
    return "Not generated"


def get_pair_output_excel_path(base_dir: Path, pair, profile):
    output_excel_name = get_pair_output_excel_name(pair, profile)
    results_dir = base_dir / profile["results_dir_name"]
    ok_path = results_dir / output_excel_name
    no_data_path = results_dir / "No-data_issue" / output_excel_name

    if ok_path.exists():
        return ok_path
    if no_data_path.exists():
        return no_data_path
    return None


def validate_summary_data_status(excel_path: Path):
    """Return OK/NOK based on workbook summary sheets.

    Sheet selection:
    - MPC workbook: validate A8C6_Summary
    - RADAR workbook: validate both A8C4_Summary and A8C5_Summary

    Rules for each validated summary sheet:
    - Any row containing "Error" => NOK
    - Any parsed "gap value = X" with X > 2 => NOK
    - Otherwise => OK
    """
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    except Exception:
        return "NOK"

    try:
        if "A8C6_Summary" in wb.sheetnames:
            summary_sheet_names = ["A8C6_Summary"]
        else:
            summary_sheet_names = ["A8C4_Summary", "A8C5_Summary"]
            if any(name not in wb.sheetnames for name in summary_sheet_names):
                return "NOK"

        for sheet_name in summary_sheet_names:
            ws = wb[sheet_name]

            for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
                value = row[0]
                if value is None:
                    continue
                text = str(value)

                if "error" in text.lower():
                    return "NOK"

                for match in GAP_VALUE_PATTERN.findall(text):
                    gap_val = int(match)
                    if gap_val > 2:
                        return "NOK"

        return "OK"
    finally:
        wb.close()


def get_pair_data_status(base_dir: Path, pair, profile):
    excel_path = get_pair_output_excel_path(base_dir, pair, profile)
    if excel_path is None:
        return "-"
    return validate_summary_data_status(excel_path)


def get_excel_failure_result(excel_path: Path):
    """Return requested failure reasons from workbook Summary sheets."""
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    except Exception:
        return "Unable to read Excel"

    try:
        summary_sheet_names = [
            sheet_name
            for sheet_name in wb.sheetnames
            if sheet_name.endswith("_Summary")
        ]
        tester_no_data = False
        gap_over_one = False

        for sheet_name in summary_sheet_names:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
                value = row[0]
                if value is None:
                    continue

                text = str(value)
                normalized = text.lower()
                if "tester_response" in normalized and "no data" in normalized:
                    tester_no_data = True

                for line in text.splitlines():
                    if "overtemp" in line.lower():
                        continue
                    for match in GAP_VALUE_PATTERN.findall(line):
                        if int(match) > 1:
                            gap_over_one = True

        failures = []
        if tester_no_data:
            failures.append("Tester response has no data")
        if gap_over_one:
            failures.append("Gap value > 1")
        return "; ".join(failures) if failures else "OK"
    finally:
        wb.close()


def get_pair_failure_result(base_dir: Path, pair, profile):
    excel_path = get_pair_output_excel_path(base_dir, pair, profile)
    if excel_path is None:
        return "-"
    return get_excel_failure_result(excel_path)


def cleanup_pair_response_folder(base_dir: Path, pair, profile):
    folder_suffix = pair.get("tester_suffix") or pair["customer_suffix"]
    response_dir = base_dir / profile["response_dir_name"] / folder_suffix
    if response_dir.exists():
        shutil.rmtree(response_dir)


def cleanup_pair_result_files(base_dir: Path, pair, profile):
    results_dir = base_dir / profile["results_dir_name"]
    output_excel_name = get_pair_output_excel_name(pair, profile)
    output_path = results_dir / output_excel_name
    no_data_path = results_dir / "No-data_issue" / output_excel_name

    if output_path.exists():
        output_path.unlink()
    if no_data_path.exists():
        no_data_path.unlink()


def process_pair(base_dir: Path, pair, results_dir: Path, profile):
    customer_suffix = pair["customer_suffix"]
    tester_suffix = pair["tester_suffix"]
    customer_path = pair["customer"]
    tester_path = pair["tester"]
    folder_suffix = tester_suffix if tester_suffix is not None else customer_suffix

    output_excel_name = get_pair_output_excel_name(pair, profile)
    pair_label = get_pair_label(pair)

    output_excel_path = results_dir / output_excel_name
    summaries = [
        f"[Pair {pair_label}] customer: {customer_path.name}",
    ]
    if tester_path is not None:
        summaries.append(f"[Pair {pair_label}] tester: {tester_path.name}")
    else:
        summaries.append(f"[Pair {pair_label}] tester: (not found)")

    if not clone_sample_excel(base_dir, output_excel_path, profile):
        summaries.append(f"[Pair {pair_label}] [Excel] Failed to clone sample workbook")
        return summaries

    did_configs = profile["did_configs"]
    excel_data = {}
    extracted_entries = {}
    for cfg in did_configs:
        excel_data[cfg["name"]] = {"customer": [], "tester": []}
        extracted_entries[cfg["name"]] = {"customer": [], "tester": []}

    pair_sources = [
        {"name": "customer", "path": customer_path, "output_role": "customer"},
    ]
    if tester_path is not None:
        pair_sources.append({"name": "tester", "path": tester_path, "output_role": "tester"})

    for source in pair_sources:
        in_path = source["path"]

        for cfg in did_configs:
            fmt_path = base_dir / cfg["format_file"]
            format_ranges = parse_format_file(fmt_path)
            payloads = extract_payloads_for_prefix(in_path, cfg["target_prefix"])
            output_dir = base_dir / profile["response_dir_name"] / folder_suffix / f"{cfg['name']}_response"
            file_prefix = f"{source['output_role']}_response_"

            if not payloads:
                summaries.append(f"[Pair {pair_label}] {cfg['name']} {source['name']}: 0 response")
                continue

            write_response_files(output_dir, file_prefix, payloads, format_ranges)
            extracted_entries[cfg["name"]][source["output_role"]].extend(payloads)

            summaries.append(
                f"[Pair {pair_label}] {cfg['name']} {source['name']}: {len(payloads)} file ({output_dir.name}/{file_prefix}1.txt ... {output_dir.name}/{file_prefix}{len(payloads)}.txt)"
            )

    for cfg in did_configs:
        did_name = cfg["name"]
        format_ranges = parse_format_file(base_dir / cfg["format_file"])
        aligned_pairs = align_responses_by_timestamp(
            extracted_entries[did_name]["customer"],
            extracted_entries[did_name]["tester"],
        )

        for pair_entry in aligned_pairs:
            customer_entry = pair_entry["customer"]
            tester_entry = pair_entry["tester"]

            if customer_entry is None:
                excel_data[did_name]["customer"].append(None)
            else:
                excel_data[did_name]["customer"].append(
                    build_output_lines(customer_entry["payload"], format_ranges)
                )

            if tester_entry is None:
                excel_data[did_name]["tester"].append(None)
            else:
                excel_data[did_name]["tester"].append(
                    build_output_lines(tester_entry["payload"], format_ranges)
                )

    excel_path = output_excel_path
    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path)
        for cfg in did_configs:
            sheet_name = cfg["name"]
            if sheet_name in wb.sheetnames:
                fmt_path = base_dir / cfg["format_file"]
                format_ranges = parse_format_file(fmt_path)
                did_written = write_responses_to_excel(
                    wb, sheet_name, base_dir,
                    excel_data[cfg["name"]], format_ranges
                )
                if did_written:
                    summaries.append(f"[Pair {pair_label}] [Excel] {sheet_name}: written to sheet")
                else:
                    summaries.append(
                        f"[Pair {pair_label}] [Excel] {sheet_name}: skipped (missing Max column or row '00')"
                    )
            else:
                summaries.append(f"[Pair {pair_label}] [Excel] Sheet '{sheet_name}' not found")

        write_summary_sheets(wb, excel_data, did_configs)
        summaries.append(f"[Pair {pair_label}] [Excel] Summary sheets updated")

        try:
            wb.save(excel_path)
        except PermissionError:
            summaries.append(
                f"[Pair {pair_label}] [Excel] Error: File {output_excel_name} is currently open. Please close the Excel file before running the script."
            )
            return summaries

        if has_no_data_issue(excel_data, did_configs):
            no_data_dir = results_dir / "No-data_issue"
            no_data_dir.mkdir(parents=True, exist_ok=True)
            dest_path = no_data_dir / output_excel_name
            shutil.move(str(excel_path), str(dest_path))
            summaries.append(f"[Pair {pair_label}] [Excel] Moved to No-data_issue/ (customer no-data detected)")
    else:
        summaries.append(f"[Pair {pair_label}] [Excel] File {output_excel_name} not found")

    return summaries


def run_analysis_for_pairs(base_dir: Path, selected_pairs, profile=None):
    profile = profile or get_analysis_profile(base_dir)
    results_dir = base_dir / profile["results_dir_name"]
    results_dir.mkdir(parents=True, exist_ok=True)
    all_summaries = []

    for pair in selected_pairs:
        cleanup_pair_response_folder(base_dir, pair, profile)
        cleanup_pair_result_files(base_dir, pair, profile)
        all_summaries.extend(process_pair(base_dir, pair, results_dir, profile))

    return all_summaries


class RadarAnalysisGui:
    def __init__(self, root, base_dir: Path):
        self.root = root
        self.base_dir = base_dir
        self.profile = get_analysis_profile(base_dir)
        self.pairs = []
        self.pair_by_row = {}
        initial_base_dir = load_saved_base_dir(base_dir)
        self.selected_base_dir = tk.StringVar(value=str(initial_base_dir))
        self.sample_status_var = tk.StringVar(value="Sample_format: Unknown")
        self.sample_path_var = tk.StringVar(value="Sample_format path: -")
        self.logs_path_var = tk.StringVar(value="Diag logs scan path: -")
        self.results_path_var = tk.StringVar(value="Results path: -")

        self.root.title(f"{self.profile['mode']} DID Analysis")
        self.root.geometry("1150x620")

        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill="x")

        folder_frame = ttk.LabelFrame(top_frame, text="Analysis Folder", padding=(8, 6))
        folder_frame.pack(fill="x", pady=(0, 8))

        folder_entry = ttk.Entry(folder_frame, textvariable=self.selected_base_dir)
        folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        browse_button = ttk.Button(folder_frame, text="Browse...", command=self.browse_folder)
        browse_button.pack(side="right")

        sample_status_label = tk.Label(
            top_frame,
            textvariable=self.sample_status_var,
            anchor="w",
            justify="left",
        )
        sample_status_label.pack(fill="x", pady=(0, 6))
        self.sample_status_label = sample_status_label

        sample_path_label = tk.Label(
            top_frame,
            textvariable=self.sample_path_var,
            anchor="w",
            justify="left",
            fg="#444444",
            font=("Segoe UI", 8),
        )
        sample_path_label.pack(fill="x")
        logs_path_label = tk.Label(
            top_frame,
            textvariable=self.logs_path_var,
            anchor="w",
            justify="left",
            fg="#444444",
            font=("Segoe UI", 8),
        )
        logs_path_label.pack(fill="x")
        results_path_label = tk.Label(
            top_frame,
            textvariable=self.results_path_var,
            anchor="w",
            justify="left",
            fg="#444444",
            font=("Segoe UI", 8),
        )
        results_path_label.pack(fill="x", pady=(0, 6))

        controls_frame = ttk.Frame(top_frame)
        controls_frame.pack(fill="x")

        selection_group = ttk.LabelFrame(controls_frame, text="Selection", padding=(8, 6))
        selection_group.pack(side="left")

        self.scan_button = ttk.Button(selection_group, text="Scan pairs", command=self.refresh_pairs)
        self.scan_button.pack(side="left", padx=(0, 8))

        self.select_all_button = ttk.Button(selection_group, text="Select all", command=self.select_all)
        self.select_all_button.pack(side="left", padx=(0, 8))

        self.clear_button = ttk.Button(selection_group, text="Clear selection", command=self.clear_selection)
        self.clear_button.pack(side="left")

        execution_group = ttk.LabelFrame(controls_frame, text="Execution", padding=(8, 6))
        execution_group.pack(side="right")

        self.execute_button = ttk.Button(execution_group, text="Execute selected", command=self.execute_selected)
        self.execute_button.pack(side="left")

        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(top_frame, textvariable=self.status_var)
        status_label.pack(anchor="e", pady=(6, 0))

        table_frame = ttk.Frame(root, padding=(10, 0, 10, 10), borderwidth=1, relief="solid")
        table_frame.pack(fill="both", expand=True)

        columns = (
            "pair",
            "customer",
            "tester",
            "tester_status",
            "result_status",
            "data_status",
            "failure_result",
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("pair", text="Pair")
        self.tree.heading("customer", text="Customer file")
        self.tree.heading("tester", text="Tester file")
        self.tree.heading("tester_status", text="Tester available")
        self.tree.heading("result_status", text="Result status")
        self.tree.heading("data_status", text="Data status")
        self.tree.heading("failure_result", text="Failure result")

        self.tree.column("pair", width=170, anchor="center")
        self.tree.column("customer", width=290)
        self.tree.column("tester", width=290)
        self.tree.column("tester_status", width=140, anchor="center")
        self.tree.column("result_status", width=180, anchor="center")
        self.tree.column("data_status", width=120, anchor="center")
        self.tree.column("failure_result", width=230, anchor="center")

        # Status color coding for quick visual scan.
        self.tree.tag_configure("status_ok", background="#d9f2d9")
        self.tree.tag_configure("status_issue", background="#ffd9d9")
        self.tree.tag_configure("status_pending", background="#fff4cc")

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.open_excel_for_row)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.refresh_pairs()

    def get_current_base_dir(self) -> Path:
        return Path(self.selected_base_dir.get()).resolve()

    def update_sample_status(self, base_dir: Path):
        self.profile = get_analysis_profile(base_dir)
        self.root.title(f"{self.profile['mode']} DID Analysis")

        sample_dir = base_dir / "Sample_format"
        logs_dir = base_dir / self.profile["logs_dir_name"]
        scan_dir = logs_dir if logs_dir.exists() else base_dir
        results_dir = base_dir / self.profile["results_dir_name"]

        self.sample_path_var.set(f"Sample_format path: {shorten_path_for_display(sample_dir)}")
        self.logs_path_var.set(
            f"{self.profile['logs_dir_name']} scan path: {shorten_path_for_display(scan_dir)}"
        )
        self.results_path_var.set(f"Results path: {shorten_path_for_display(results_dir)}")

        if sample_dir.exists() and sample_dir.is_dir():
            self.sample_status_var.set("Sample_format: Available")
            self.sample_status_label.configure(fg="#0b7a0b")
        else:
            self.sample_status_var.set("Sample_format: Not available")
            self.sample_status_label.configure(fg="#b22222")

    def browse_folder(self):
        selected = filedialog.askdirectory(initialdir=self.selected_base_dir.get())
        if not selected:
            return
        selected_path = Path(selected)
        self.selected_base_dir.set(str(selected_path))
        save_selected_base_dir(self.base_dir, selected_path)
        self.refresh_pairs()

    def refresh_pairs(self):
        current_base_dir = self.get_current_base_dir()
        if current_base_dir.exists() and current_base_dir.is_dir():
            save_selected_base_dir(self.base_dir, current_base_dir)
        self.update_sample_status(current_base_dir)
        self.pairs = discover_trace_pairs(current_base_dir, self.profile)
        self.pairs.sort(
            key=lambda pair: (
                natural_sort_key(pair["customer"].name),
                natural_sort_key(pair["tester"].name if pair["tester"] else ""),
            )
        )
        self.pair_by_row = {}

        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, pair in enumerate(self.pairs):
            pair_label = get_pair_label(pair)
            customer_name = pair["customer"].name
            tester_name = pair["tester"].name if pair["tester"] else "(not found)"
            tester_status = "Yes" if pair["tester"] else "No"
            result_status = get_pair_result_status(current_base_dir, pair, self.profile)
            data_status = get_pair_data_status(current_base_dir, pair, self.profile)
            failure_result = get_pair_failure_result(current_base_dir, pair, self.profile)
            row_tag = self.get_result_status_tag(result_status)

            row_id = f"pair_{idx}"
            self.tree.insert(
                "",
                "end",
                iid=row_id,
                values=(
                    pair_label,
                    customer_name,
                    tester_name,
                    tester_status,
                    result_status,
                    data_status,
                    failure_result,
                ),
                tags=(row_tag,),
            )
            self.pair_by_row[row_id] = pair

        self.status_var.set(
            f"[{self.profile['mode']}] Scanned {len(self.pairs)} pair(s) in {current_base_dir.name}"
        )

    def get_result_status_tag(self, result_status: str) -> str:
        if result_status == "OK (available)":
            return "status_ok"
        if result_status == "No-data issue":
            return "status_issue"
        return "status_pending"

    def select_all(self):
        items = self.tree.get_children()
        self.tree.selection_set(items)

    def clear_selection(self):
        self.tree.selection_remove(self.tree.selection())

    def open_excel_for_row(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        pair = self.pair_by_row.get(row_id)
        if pair is None:
            return

        excel_path = get_pair_output_excel_path(
            self.get_current_base_dir(), pair, self.profile
        )
        if excel_path is None:
            messagebox.showwarning(
                "Excel file not found",
                "No generated Excel file is available for this pair.",
            )
            return

        try:
            os.startfile(str(excel_path))
        except OSError as exc:
            messagebox.showerror(
                "Open Excel failed",
                f"Could not open {excel_path.name}: {exc}",
            )

    def set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.scan_button.configure(state=state)
        self.execute_button.configure(state=state)
        self.select_all_button.configure(state=state)
        self.clear_button.configure(state=state)

    def execute_selected(self):
        selected_ids = list(self.tree.selection())
        if not selected_ids:
            messagebox.showwarning("No selection", "Please select at least one pair to execute.")
            return

        current_base_dir = self.get_current_base_dir()
        selected_pairs = [self.pair_by_row[row_id] for row_id in selected_ids]
        self.set_busy(True)
        self.status_var.set(f"Executing {len(selected_pairs)} selected pair(s)...")

        worker = threading.Thread(
            target=self._run_selected_pairs_worker,
            args=(current_base_dir, selected_pairs),
            daemon=True,
        )
        worker.start()

    def _run_selected_pairs_worker(self, current_base_dir: Path, selected_pairs):
        try:
            summaries = run_analysis_for_pairs(current_base_dir, selected_pairs)
            self.root.after(0, self._on_execute_done, summaries)
        except Exception as exc:
            self.root.after(0, self._on_execute_failed, exc)

    def _on_execute_done(self, summaries):
        self.set_busy(False)
        self.refresh_pairs()
        self.status_var.set("Execution completed")
        messagebox.showinfo("Execution completed", "\n".join(summaries) if summaries else "No output.")

    def _on_execute_failed(self, exc):
        self.set_busy(False)
        self.status_var.set("Execution failed")
        messagebox.showerror("Execution failed", f"An error occurred: {exc}")


def launch_gui():
    root = tk.Tk()
    RadarAnalysisGui(root, resolve_runtime_base_dir())
    root.mainloop()


def main():
    base_dir = resolve_runtime_base_dir()
    profile = get_analysis_profile(base_dir)
    results_dir = base_dir / profile["results_dir_name"]
    all_summaries = []

    # Cleanup old response folders and discover all trace pairs.
    print(f"Performing cleanup for {profile['mode']}...")
    cleanup_response_folders(base_dir, profile)
    cleanup_results_folder(base_dir, profile)
    results_dir.mkdir(parents=True, exist_ok=True)

    pairs = discover_trace_pairs(base_dir, profile)

    if not pairs:
        all_summaries.append("No customer .asc files found.")
        print("\n".join(all_summaries))
        return

    print()

    all_summaries.extend(run_analysis_for_pairs(base_dir, pairs))

    print("\n".join(all_summaries))


if __name__ == "__main__":
    if "--cli" in sys.argv:
        main()
    else:
        launch_gui()