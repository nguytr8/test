from pathlib import Path
import re
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Configuration
RESULTS_DIR_NAME = "MaxHold_Re-sim_RADAR_results"
OUTPUT_FILE = "MaxHold_Re-sim_RADAR_results_summary.xlsx"
HEADER_KEYWORD_AVG_GAP = "average gap"
HIGHLIGHT_THRESHOLD = 1.0


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


def parse_number(value):
    """Parse numeric value from plain cells or supported Excel formulas.

    Supported formula patterns for Customer/Tester template cells:
    - =HEX2DEC("FF")
    - =_xlfn.HEX2DEC("FF")
    - =DECIMAL("FF",16)
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None

        if text.startswith("="):
            formula = text[1:].strip()

            # HEX2DEC("xx") or _xlfn.HEX2DEC("xx")
            hex2dec_match = re.search(
                r"(?:_xlfn\.)?hex2dec\(\s*\"([0-9a-fA-F]+)\"\s*\)",
                formula,
                re.IGNORECASE,
            )
            if hex2dec_match:
                return float(int(hex2dec_match.group(1), 16))

            # DECIMAL("xx",16)
            decimal16_match = re.search(
                r"decimal\(\s*\"([0-9a-fA-F]+)\"\s*,\s*16\s*\)",
                formula,
                re.IGNORECASE,
            )
            if decimal16_match:
                return float(int(decimal16_match.group(1), 16))

            return None

        try:
            return float(text)
        except ValueError:
            return None

    return None


def detect_customer_tester_pairs(ws):
    """Detect Customer_response_n / Tester_response_n column pairs from header row."""
    customer_cols = {}
    tester_cols = {}

    for cell in ws[1]:
        if not cell.value:
            continue

        header = str(cell.value).strip().lower().replace("\n", " ")

        customer_match = re.search(r"customer_response_(\d+)", header)
        if customer_match:
            customer_cols[int(customer_match.group(1))] = cell.column
            continue

        tester_match = re.search(r"tester_response_(\d+)", header)
        if tester_match:
            tester_cols[int(tester_match.group(1))] = cell.column

    run_indexes = sorted(set(customer_cols.keys()) & set(tester_cols.keys()))
    return [(customer_cols[idx], tester_cols[idx]) for idx in run_indexes]


def calc_average_gap_for_row(ws, row_num: int, column_pairs):
    """Rebuild Average gap_value for one row from Customer/Tester pairs."""
    gap_values = []
    for customer_col, tester_col in column_pairs:

        customer_val = parse_number(ws.cell(row=row_num, column=customer_col).value)
        tester_val = parse_number(ws.cell(row=row_num, column=tester_col).value)

        # Match Excel logic: gap is blank if either side is blank.
        if customer_val is None or tester_val is None:
            continue

        gap_values.append(abs(customer_val - tester_val))

    if not gap_values:
        return None

    return sum(gap_values) / len(gap_values)


def get_max_average_gap(ws, sheet_name: str):
    """Get max recalculated Average gap_value from a DID sheet."""
    if ws.max_row < 2:
        return None

    column_pairs = detect_customer_tester_pairs(ws)
    if not column_pairs:
        return None

    end_row = ws.max_row
    if sheet_name == "A8C5":
        end_row = max(1, ws.max_row - 2)

    max_val = None
    for row_num in range(2, end_row + 1):
        avg_gap = calc_average_gap_for_row(ws, row_num, column_pairs)
        if avg_gap is None:
            continue
        if max_val is None or avg_gap > max_val:
            max_val = avg_gap

    return max_val


def apply_all_borders(ws, min_row, max_row, min_col, max_col):
    """Apply thin all-borders to every cell in the given range."""
    thin = Side(style="thin")
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = thin_border


def apply_thick_outside_border(ws, min_row, max_row, min_col, max_col):
    """Apply thick outside border around the given rectangle, keeping inner borders thin."""
    thick = Side(style="medium")
    thin = Side(style="thin")

    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        for cell in row:
            is_top = cell.row == min_row
            is_bottom = cell.row == max_row
            is_left = cell.column == min_col
            is_right = cell.column == max_col

            cell.border = Border(
                top=thick if is_top else thin,
                bottom=thick if is_bottom else thin,
                left=thick if is_left else thin,
                right=thick if is_right else thin,
            )


def extract_log_name(file_name: str) -> str:
    """Extract log name from file name by stripping prefix and _run_N suffix.

    Example: MaxHold_Re-sim_RADAR_#2_100k_run_1.xlsx -> #2_100k
    """
    stem = Path(file_name).stem  # remove .xlsx
    # Remove known prefix
    prefix = "MaxHold_Re-sim_RADAR_"
    if stem.startswith(prefix):
        stem = stem[len(prefix):]
    # Remove trailing _run_<digits>
    stem = re.sub(r"_run_\d+$", "", stem)
    return stem


def build_file_sort_key(file_name: str):
    """Build natural numeric sort key so #2 comes before #10."""
    stem = Path(file_name).stem
    hash_pos = stem.find("#")
    sortable_part = stem[hash_pos + 1 :] if hash_pos != -1 else stem

    # Split into numeric/non-numeric chunks for natural ordering.
    chunks = re.split(r"(\d+)", sortable_part)
    key = []
    for chunk in chunks:
        if not chunk:
            continue
        if chunk.isdigit():
            key.append((0, int(chunk)))
        else:
            key.append((1, chunk.lower()))

    # Keep deterministic ordering for edge cases.
    return key, file_name.lower()


def main():
    base_dir = Path(__file__).resolve().parent
    results_dir = base_dir / RESULTS_DIR_NAME

    if not results_dir.exists():
        print(f"Error: {RESULTS_DIR_NAME} folder not found at {base_dir}")
        return

    # Find all Excel files recursively
    excel_files = sorted(
        path for path in results_dir.glob("**/*.xlsx") if not path.name.startswith("~$")
    )
    
    if not excel_files:
        print(f"No Excel files found in {results_dir}")
        return

    print(f"Found {len(excel_files)} Excel files")

    # Collect summary data
    summary_data = []
    for excel_path in excel_files:
        file_name = excel_path.name
        
        try:
            # Load workbook normally (don't need data_only since we compute values manually)
            wb = openpyxl.load_workbook(excel_path)
            
            # Get A8C4 max average gap
            a8c4_max = None
            if "A8C4" in wb.sheetnames:
                a8c4_max = get_max_average_gap(wb["A8C4"], "A8C4")
            
            # Get A8C5 max average gap
            a8c5_max = None
            if "A8C5" in wb.sheetnames:
                a8c5_max = get_max_average_gap(wb["A8C5"], "A8C5")
            
            wb.close()
            
            summary_data.append({
                "file_name": file_name,
                "log_name": extract_log_name(file_name),
                "a8c4_max": a8c4_max,
                "a8c5_max": a8c5_max,
            })
            
            print(f"  {file_name}: A8C4={a8c4_max}, A8C5={a8c5_max}")
        
        except Exception as e:
            print(f"  Error reading {file_name}: {e}")
            continue

    # Sort rows by file name in ascending natural numeric order (#1, #2, ... #10 ...)
    summary_data.sort(key=lambda row: build_file_sort_key(row["file_name"]))

    # Create output Excel file
    output_path = base_dir / OUTPUT_FILE
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Add headers
    headers = [
        "Log name",
        "File name",
        "A8C4 max average gap",
        "A8C5 max average gap (except 2 temperature values)",
    ]
    for col_idx, header_text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Add data rows
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    num_cols = 4

    # Build consecutive groups by log_name
    groups_data = []
    if summary_data:
        current_group = [summary_data[0]]
        for data in summary_data[1:]:
            if data["log_name"] == current_group[-1]["log_name"]:
                current_group.append(data)
            else:
                groups_data.append(current_group)
                current_group = [data]
        groups_data.append(current_group)

    # Write group by group; insert blank separator row between groups
    excel_row = 2
    for g_idx, group in enumerate(groups_data):
        group_start_row = excel_row

        for data in group:
            # Log name (col A)
            cell_log = ws.cell(row=excel_row, column=1, value=data["log_name"])
            cell_log.alignment = Alignment(horizontal="center", vertical="center")

            # File name (col B)
            ws.cell(row=excel_row, column=2, value=data["file_name"]).alignment = Alignment(
                horizontal="left", vertical="center"
            )

            # A8C4 max average gap (col C)
            cell_b = ws.cell(row=excel_row, column=3, value=data["a8c4_max"])
            cell_b.alignment = center_align
            if data["a8c4_max"] is not None:
                if abs(data["a8c4_max"] - HIGHLIGHT_THRESHOLD) < 1e-9:
                    cell_b.fill = yellow_fill
                elif data["a8c4_max"] > HIGHLIGHT_THRESHOLD:
                    cell_b.fill = red_fill

            # A8C5 max average gap (col D)
            cell_c = ws.cell(row=excel_row, column=4, value=data["a8c5_max"])
            cell_c.alignment = center_align
            if data["a8c5_max"] is not None:
                if abs(data["a8c5_max"] - HIGHLIGHT_THRESHOLD) < 1e-9:
                    cell_c.fill = yellow_fill
                elif data["a8c5_max"] > HIGHLIGHT_THRESHOLD:
                    cell_c.fill = red_fill

            excel_row += 1

        group_end_row = excel_row - 1

        # Merge Log name cells for this group
        if group_end_row > group_start_row:
            ws.merge_cells(
                start_row=group_start_row, start_column=1,
                end_row=group_end_row, end_column=1,
            )

        # Thin all-borders then thick outside border for this group
        apply_all_borders(ws, group_start_row, group_end_row, 1, num_cols)
        apply_thick_outside_border(ws, group_start_row, group_end_row, 1, num_cols)

        # Blank separator row between groups (not after the last group)
        if g_idx < len(groups_data) - 1:
            excel_row += 1  # leave this row empty (no fill, no border)

    # Auto-size columns
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 42

    # Save output file
    wb.save(output_path)
    print(f"\nOutput saved to: {output_path}")
    print(f"Total files processed: {len(summary_data)}")


if __name__ == "__main__":
    main()
