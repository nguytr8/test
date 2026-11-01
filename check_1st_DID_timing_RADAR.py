from pathlib import Path
import re
import openpyxl


BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "Diag_RADAR_logs"
OUTPUT_FILE = BASE_DIR / "result_check_1st_DID_timing_RADAR.txt"
RESULTS_DIR = BASE_DIR / "MaxHold_Re-sim_RADAR_results"
RESULTS_FILE_PREFIX = "MaxHold_Re-sim_RADAR_"

TARGET_ADDRESS_PREFIX = "18DAB0F1"
PAYLOAD_A8C4 = "03 22 A8 C4 00 00 00 00"
PAYLOAD_A8C5 = "03 22 A8 C5 00 00 00 00"
LOW_TIMING_THRESHOLD_SECONDS = 0.7


def is_target_address_line(line: str) -> bool:
    """Check if line contains target tester address 18DAB0F1x / 18DAB0F1n."""
    if TARGET_ADDRESS_PREFIX not in line:
        return False

    pos = line.find(TARGET_ADDRESS_PREFIX)
    if pos < 0:
        return False

    next_index = pos + len(TARGET_ADDRESS_PREFIX)
    if next_index >= len(line):
        return False

    next_char = line[next_index]
    return next_char.lower() in "0123456789abcdefx"


def parse_timestamp(line: str) -> str | None:
    """Get the first token as timestamp if convertible to float."""
    tokens = line.split()
    if not tokens:
        return None

    try:
        float(tokens[0])
    except ValueError:
        return None

    return tokens[0]


def find_first_timings(file_path: Path) -> dict[str, str]:
    first = {}

    with file_path.open("r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            if not is_target_address_line(line):
                continue

            if "A8C4" not in first and PAYLOAD_A8C4 in line:
                ts = parse_timestamp(line)
                if ts is not None:
                    first["A8C4"] = ts

            if "A8C5" not in first and PAYLOAD_A8C5 in line:
                ts = parse_timestamp(line)
                if ts is not None:
                    first["A8C5"] = ts

            if "A8C4" in first and "A8C5" in first:
                break

    return first


def build_output_line(file_stem: str, first: dict[str, str]) -> str:
    parts = []

    if "A8C4" in first:
        parts.append(f"A8C4 - 1st read at {first['A8C4']}")

    if "A8C5" in first:
        parts.append(f"A8C5 - 1st read at {first['A8C5']}")

    return f"{file_stem}: " + " & ".join(parts)


def natural_sort_key(text: str) -> tuple:
    """Sort text by numeric chunks"""
    parts = re.split(r"(\d+)", text)
    return tuple(int(part) if part.isdigit() else part.lower() for part in parts)


def get_min_first_read(first: dict[str, str]) -> tuple[str, str, float] | None:
    """Return (DID, timestamp_text, timestamp_float) with the earliest first read."""
    best = None

    for did in ("A8C4", "A8C5"):
        if did not in first:
            continue
        try:
            ts_float = float(first[did])
        except ValueError:
            continue

        if best is None or ts_float < best[2]:
            best = (did, first[did], ts_float)

    return best


def short_case_name(file_stem: str) -> str:
    prefix = "Diag_customer_RADAR_"
    if file_stem.startswith(prefix):
        return file_stem[len(prefix):]
    return file_stem


def find_column_by_header(ws, target_header: str):
    """Return 1-based column index for the requested header on row 1."""
    target = target_header.strip().lower()
    for cell in ws[1]:
        if cell.value is None:
            continue
        if str(cell.value).strip().lower() == target:
            return cell.column
    return None


def compare_customer_read_2_vs_1(case_name: str, did_sheet: str) -> str:
    """
    Compare customer_response_2 vs customer_response_1 in the DID sheet.
    Returns one of:
    - "2nd Read != 1st Read"
    - "2nd Read = 1st Read"
    - "Only read 1 times"
    - "Result xlsx not found"
    """
    workbook_path = RESULTS_DIR / f"{RESULTS_FILE_PREFIX}{case_name}.xlsx"
    if not workbook_path.exists():
        return "Result xlsx not found"

    wb = openpyxl.load_workbook(workbook_path, data_only=False)
    try:
        if did_sheet not in wb.sheetnames:
            return "Only read 1 times"

        ws = wb[did_sheet]
        col_1 = find_column_by_header(ws, "customer_response_1")
        col_2 = find_column_by_header(ws, "customer_response_2")

        if col_1 is None or col_2 is None:
            return "Only read 1 times"

        for row in range(2, ws.max_row + 1):
            val_1 = ws.cell(row=row, column=col_1).value
            val_2 = ws.cell(row=row, column=col_2).value
            if val_1 != val_2:
                return "2nd Read != 1st Read"

        return "2nd Read = 1st Read"
    finally:
        wb.close()


def main() -> None:
    asc_files = sorted(
        INPUT_DIR.glob("Diag_customer_RADAR*.asc"),
        key=lambda p: natural_sort_key(p.stem),
    )
    output_lines = []
    low_timing_lines = []
    status_cases = {
        "2nd Read != 1st Read": [],
        "2nd Read = 1st Read": [],
        "Only read 1 times": [],
        "Result xlsx not found": [],
    }

    for asc_file in asc_files:
        first = find_first_timings(asc_file)
        if not first:
            continue

        output_lines.append(build_output_line(asc_file.stem, first))

        min_read = get_min_first_read(first)
        if min_read is None:
            continue

        did, ts_text, ts_float = min_read
        if ts_float < LOW_TIMING_THRESHOLD_SECONDS:
            case_name = short_case_name(asc_file.stem)
            read_compare = compare_customer_read_2_vs_1(case_name, did)
            status_cases[read_compare].append(case_name)
            low_timing_lines.append(
                f"{case_name}: {did} - 1st read at {ts_text} | {read_compare}"
            )

    if low_timing_lines:
        output_lines.append("")
        output_lines.append(
            f"->  Below logs have 1st timing of reading DID < {LOW_TIMING_THRESHOLD_SECONDS:g}s:"
        )
        output_lines.extend(low_timing_lines)
        output_lines.append("")
        output_lines.append("Summary:")
        for status, cases in status_cases.items():
            if cases:
                cases_str = ", ".join(cases)
                output_lines.append(f"  {status}: {len(cases)} ({cases_str})")
            else:
                output_lines.append(f"  {status}: 0")

    OUTPUT_FILE.write_text("\n".join(output_lines), encoding="utf-8")
    print(f"Saved {len(output_lines)} lines to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
