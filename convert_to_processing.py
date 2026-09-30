"""
Load RA and FSC input workbooks, convert selected tabs to cleaned dataframes,
save them to processing/, then enrich RA with fuel surcharge values and save to output/.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from excel_formatting import save_formatted_rate_card
from project_paths import (
    INPUT_FSC_DIR,
    INPUT_RA_DIR,
    OUTPUT_DIR,
    PROCESSING_DIR,
    ensure_workspace_dirs,
)

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
RA_SHEET_NAME = "Rate card"
RA_HEADER_MARKER = "lane #"
FSC_HEADER_MARKER = "row number"
TAB_FUEL_RATES_MARKER = "+ Fuel Rates"
THREE_LETTER_CODE_PATTERN = re.compile(r"^[A-Za-z]{3}$")
TWO_LETTER_CODE_PATTERN = re.compile(r"^[A-Za-z]{2}$")
BLANK_CITY_VALUES = {"", "-", "n/a", "#n/a", "na", "none", "null"}
CITY_SYNONYM_GROUPS = (
    ("Findley", "Findlay"),
    ("Bangalore", "Bengaluru"),
    ("United Arab Emirates", "UAE"),
    ("Turkey", "Turkiye"),
    ("Azerbaijan", "Azerbaijaan"),
)


@dataclass(frozen=True)
class ConvertResult:
    ra_file: Path
    fsc_file: Path
    fsc_sheet: str
    ra_processing_path: Path
    fsc_processing_path: Path
    output_path: Path
    matched_lane_count: int
    total_lane_count: int


def select_input_file(files: list[Path], folder: Path, *, auto: bool, label: str) -> Path:
    if not files:
        raise FileNotFoundError(f"No Excel files found in: {folder}")

    if auto:
        selected = files[0]
        if len(files) == 1:
            print(f"\nAuto mode: using {label} file {selected.name}")
        else:
            print(f"\nAuto mode: using first {label} file {selected.name}")
        return selected

    return prompt_file_selection(f"Select {label} file from {folder.name}:", folder, files)


def select_fsc_sheet(file_path: Path, sheet_names: list[str], *, auto: bool) -> str:
    if auto:
        if len(sheet_names) == 1:
            print(f"Auto mode: using only tab in {file_path.name}: {sheet_names[0]}")
            return sheet_names[0]
        selected = sheet_names[0]
        print(f"Auto mode: using first tab in {file_path.name}: {selected}")
        return selected

    return prompt_sheet_selection(file_path, sheet_names)


def list_excel_files(folder: Path) -> list[Path]:
    return [
        path
        for path in sorted(folder.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_fuel_surcharge_value(value: object) -> object:
    """Normalize fuel surcharge values: decimal comma to dot, prefer numeric cells."""
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = _cell_text(value)
    if not text:
        return pd.NA

    normalized_text = text.replace(",", ".")
    try:
        return float(normalized_text)
    except ValueError:
        return normalized_text


def _fuel_surcharge_display_value(value: object) -> object:
    """Format fuel surcharge for output using a dot decimal separator."""
    normalized = _normalize_fuel_surcharge_value(value)
    if pd.isna(normalized):
        return pd.NA

    if isinstance(normalized, (int, float)) and not isinstance(normalized, bool):
        text = format(float(normalized), "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text or "0"

    return _cell_text(normalized).replace(",", ".")


def _is_meaningful_city_value(value: object) -> bool:
    text = _cell_text(value).casefold()
    return bool(text) and text not in BLANK_CITY_VALUES


def _first_existing_column(columns: list[str], names: tuple[str, ...]) -> str | None:
    for name in names:
        column = _optional_column(columns, name)
        if column is not None:
            return column
    return None


def _normalize_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = _cell_text(header)
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        count = seen.get(base, 0)
        if count:
            value = f"{base}_{count + 1}"
        seen[base] = count + 1
        cleaned.append(value)

    return cleaned


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    keep_columns = [
        column
        for column in df.columns
        if any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def _find_header_row_index(df_raw: pd.DataFrame, marker: str, max_scan_rows: int = 40) -> int | None:
    marker_lower = marker.lower()
    scan_limit = min(len(df_raw), max_scan_rows)

    for row_idx in range(scan_limit):
        row = df_raw.iloc[row_idx]
        for value in row:
            if _cell_text(value).lower() == marker_lower:
                return row_idx
    return None


def _raw_sheet_to_df(df_raw: pd.DataFrame, header_marker: str) -> pd.DataFrame:
    header_row_idx = _find_header_row_index(df_raw, header_marker)
    if header_row_idx is None:
        raise ValueError(f"Could not find header row containing '{header_marker}'")

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    return df.reset_index(drop=True)


def ra_rate_card_to_df(file_path: Path) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=RA_SHEET_NAME, header=None)
    return _raw_sheet_to_df(raw, RA_HEADER_MARKER)


def fsc_tab_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    return _raw_sheet_to_df(raw, FSC_HEADER_MARKER)


def prompt_file_selection(title: str, folder: Path, files: list[Path]) -> Path:
    if not files:
        raise FileNotFoundError(f"No Excel files found in: {folder}")

    print(f"\n{title}")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")

    while True:
        raw = input("Enter file number: ").strip()
        try:
            choice = int(raw) - 1
            if 0 <= choice < len(files):
                return files[choice]
        except ValueError:
            pass
        print(f"Please enter a number between 1 and {len(files)}.")


def prompt_sheet_selection(file_path: Path, sheet_names: list[str]) -> str:
    if len(sheet_names) == 1:
        print(f"\nUsing only tab in {file_path.name}: {sheet_names[0]}")
        return sheet_names[0]

    print(f"\nAvailable tabs in {file_path.name}:")
    for index, name in enumerate(sheet_names, start=1):
        print(f"  {index}. {name}")

    while True:
        raw = input("Enter tab number: ").strip()
        try:
            choice = int(raw) - 1
            if 0 <= choice < len(sheet_names):
                return sheet_names[choice]
        except ValueError:
            pass
        print(f"Please enter a number between 1 and {len(sheet_names)}.")


def sanitize_output_stem(text: str) -> str:
    return re.sub(r"[^\w\-]+", "_", text).strip("_") or "output"


def save_dataframe(df: pd.DataFrame, output_path: Path, sheet_name: str = "Sheet1") -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output_path


def _find_column(columns: list[str], name: str) -> str:
    target = name.strip().lower()
    for column in columns:
        if _cell_text(column).lower() == target:
            return column
    raise ValueError(f"Could not find column '{name}'")


def _find_destination_city_column(columns: list[str]) -> str:
    try:
        return _find_column(columns, "Destination City")
    except ValueError:
        return _find_column(columns, "Destination")


def _optional_column(columns: list[str], name: str) -> str | None:
    try:
        return _find_column(columns, name)
    except ValueError:
        return None


def _columns_by_normalized_names(columns: list[str], ordered_names: tuple[str, ...]) -> list[str]:
    """Return columns whose names match any of the given labels (case-insensitive), in sheet order."""
    pattern_set = {_cell_text(name).casefold() for name in ordered_names}
    return [
        column
        for column in columns
        if _cell_text(column).casefold() in pattern_set
    ]


def _origin_lookup_columns(columns: list[str]) -> list[str]:
    lookup_columns = _columns_by_normalized_names(
        columns,
        ("Origin City", "Orig City", "Origin City_2"),
    )
    if not lookup_columns:
        raise ValueError("Could not find origin city columns in RA dataframe")
    return lookup_columns


def _destination_lookup_columns(columns: list[str]) -> list[str]:
    lookup_columns: list[str] = []
    for names in (
        ("Destination",),
        ("Dest City", "Destination City"),
        ("Destination_2",),
        ("Destination City_2",),
    ):
        for column in _columns_by_normalized_names(columns, names):
            if column not in lookup_columns:
                lookup_columns.append(column)

    if not lookup_columns:
        raise ValueError("Could not find destination columns in RA dataframe")
    return lookup_columns


def _normalize_city_key(value: object) -> str:
    text = _cell_text(value)
    return re.sub(r"\s+", "", text).casefold()


def _build_synonym_expansion() -> dict[str, set[str]]:
    expansion: dict[str, set[str]] = {}
    for group in CITY_SYNONYM_GROUPS:
        normalized_group = {key for name in group if (key := _normalize_city_key(name))}
        for key in normalized_group:
            expansion[key] = normalized_group
    return expansion


_SYNONYM_EXPANSION = _build_synonym_expansion()


def _expand_synonym_keys(keys: set[str]) -> set[str]:
    expanded = set(keys)
    for key in keys:
        expanded |= _SYNONYM_EXPANSION.get(key, set())
    return expanded


def _city_alias_keys(value: object) -> set[str]:
    """Return normalized city keys, expanding slash-separated and synonym aliases."""
    text = _cell_text(value)
    if not text:
        return set()

    parts = [part.strip() for part in text.split("/") if part.strip()]
    if not parts:
        return set()

    keys = {key for part in parts if (key := _normalize_city_key(part))}
    if len(parts) > 1:
        combined_key = _normalize_city_key(text)
        if combined_key:
            keys.add(combined_key)
    return _expand_synonym_keys(keys)


def _is_three_letter_code(value: object) -> bool:
    return bool(THREE_LETTER_CODE_PATTERN.fullmatch(_cell_text(value)))


def _is_two_letter_code(value: object) -> bool:
    return bool(TWO_LETTER_CODE_PATTERN.fullmatch(_cell_text(value)))


def _origin_keys_for_column(row: pd.Series, column_name: str) -> set[str]:
    value = row[column_name]
    if _is_three_letter_code(value):
        return set()
    if not _is_meaningful_city_value(value):
        return set()
    return _city_alias_keys(value)


def _destination_keys_for_column(row: pd.Series, column_name: str) -> set[str]:
    value = row[column_name]
    if _is_two_letter_code(value) or _is_three_letter_code(value):
        return set()
    if not _is_meaningful_city_value(value):
        return set()
    return _city_alias_keys(value)


def _lookup_column_pairs(
    origin_columns: list[str],
    destination_columns: list[str],
) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()

    def add_pair(origin_column: str, destination_column: str) -> None:
        key = (origin_column, destination_column)
        if key not in seen:
            seen.add(key)
            pairs.append(key)

    if origin_columns and destination_columns:
        add_pair(origin_columns[0], destination_columns[0])

    for origin_column in origin_columns:
        for destination_column in destination_columns:
            add_pair(origin_column, destination_column)

    return pairs


def _lookup_fsc_value(
    row: pd.Series,
    fsc_lookup: dict[tuple[str, str], object],
    origin_columns: list[str],
    destination_columns: list[str],
) -> object:
    for origin_column, destination_column in _lookup_column_pairs(origin_columns, destination_columns):
        origin_keys = _origin_keys_for_column(row, origin_column)
        destination_keys = _destination_keys_for_column(row, destination_column)
        if not origin_keys or not destination_keys:
            continue

        for destination_key in sorted(destination_keys):
            for origin_key in sorted(origin_keys):
                value = fsc_lookup.get((origin_key, destination_key), pd.NA)
                if pd.notna(value):
                    return _fuel_surcharge_display_value(value)
    return pd.NA


def _build_fsc_lookup(
    fsc_df: pd.DataFrame,
    origin_column: str,
    destination_column: str,
    value_column: str,
) -> dict[tuple[str, str], object]:
    lookup: dict[tuple[str, str], object] = {}
    for _, row in fsc_df.iterrows():
        value = row[value_column]
        normalized_value = _normalize_fuel_surcharge_value(value)
        for origin_key in _city_alias_keys(row[origin_column]):
            for destination_key in _city_alias_keys(row[destination_column]):
                lookup[(origin_key, destination_key)] = normalized_value
    return lookup


def trim_ra_after_first_currency(df: pd.DataFrame) -> pd.DataFrame:
    """Keep columns up to and including the first Currency column."""
    currency_idx = next(
        (idx for idx, column in enumerate(df.columns) if _cell_text(column).lower() == "currency"),
        None,
    )
    if currency_idx is None:
        raise ValueError("Could not find Currency column in RA dataframe")
    return df.iloc[:, : currency_idx + 1].copy()


def fsc_value_column_before_difference_mom(fsc_df: pd.DataFrame) -> str:
    columns = list(fsc_df.columns)
    try:
        diff_idx = next(
            idx for idx, column in enumerate(columns) if _cell_text(column) == "Difference MoM"
        )
    except StopIteration as exc:
        raise ValueError("Could not find 'Difference MoM' column in FSC dataframe") from exc

    if diff_idx == 0:
        raise ValueError("No value column found before 'Difference MoM' in FSC dataframe")
    return columns[diff_idx - 1]


def fuel_surcharge_column_name(fsc_value_column: str) -> str:
    return f"Fuel Surcharge ({fsc_value_column})"


def fsc_period_label_from_value_column(value_column: str) -> str:
    text = re.sub(r"\s*Value\s*\(USD\)\s*$", "", _cell_text(value_column), flags=re.IGNORECASE)
    # Column headers sometimes include a year marker (e.g. Oct'26) that Calc. Rules omits.
    text = re.sub(r"'\d{2}", "", text)
    return text.strip()


def _normalize_period_key(value: object) -> str:
    text = re.sub(r"'\d{2}", "", _cell_text(value))
    return re.sub(r"\s+", "", text).casefold()


def _format_valid_to_date(value: object) -> str:
    if pd.isna(value):
        raise ValueError("Missing valid-to date value in FSC file")

    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")

    text = _cell_text(value)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue

    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"Could not parse valid-to date: {value}")
    return parsed.strftime("%d.%m.%Y")


def find_fsc_period_valid_to_date(fsc_file: Path, sheet_name: str, period_label: str) -> str:
    raw = pd.read_excel(fsc_file, sheet_name=sheet_name, header=None)
    calc_rules_col: int | None = None
    scan_header_rows = min(10, len(raw))

    for row_idx in range(scan_header_rows):
        for col_idx in range(raw.shape[1]):
            if _cell_text(raw.iloc[row_idx, col_idx]).lower() == "calc. rules":
                calc_rules_col = col_idx
                break
        if calc_rules_col is not None:
            break

    if calc_rules_col is None:
        raise ValueError("Could not find 'Calc. Rules' column in FSC file")

    to_col = calc_rules_col + 2
    if to_col >= raw.shape[1]:
        raise ValueError("Could not find 'To (yy/mm/dd)' column in FSC file")

    target_period = _normalize_period_key(period_label)
    for row_idx in range(len(raw)):
        period_value = _normalize_period_key(raw.iloc[row_idx, calc_rules_col])
        if period_value == target_period:
            return _format_valid_to_date(raw.iloc[row_idx, to_col])

    raise ValueError(f"Could not find FSC validity period '{period_label}' in Calc. Rules")


def _apply_valid_to_updates(result: pd.DataFrame, surcharge_column: str, valid_to_date: str) -> pd.DataFrame:
    valid_to_column = _find_column(list(result.columns), "Valid to")
    updated = result.copy()
    fuel_lane_mask = updated[surcharge_column].notna()
    updated.loc[fuel_lane_mask, valid_to_column] = valid_to_date
    return updated


def _lane_eligible_for_fuel_surcharge(tab_value: object) -> bool:
    return TAB_FUEL_RATES_MARKER in _cell_text(tab_value)


def apply_fuel_surcharge_from_fsc(
    ra_df: pd.DataFrame,
    fsc_df: pd.DataFrame,
    *,
    fsc_file: Path | None = None,
    fsc_sheet: str | None = None,
) -> pd.DataFrame:
    """Trim RA columns, append fuel surcharge values, and update Valid to from FSC."""
    fsc_value_column = fsc_value_column_before_difference_mom(fsc_df)
    surcharge_column = fuel_surcharge_column_name(fsc_value_column)

    trimmed_ra = trim_ra_after_first_currency(ra_df)
    currency_column = _find_column(list(trimmed_ra.columns), "Currency")
    tab_column = _find_column(list(trimmed_ra.columns), "Tab")
    ra_columns = list(trimmed_ra.columns)
    origin_columns = _origin_lookup_columns(ra_columns)
    destination_columns = _destination_lookup_columns(ra_columns)

    fsc_origin_column = _find_column(list(fsc_df.columns), "Origin City")
    fsc_destination_column = _find_column(list(fsc_df.columns), "Destination City")

    fsc_lookup = _build_fsc_lookup(
        fsc_df,
        fsc_origin_column,
        fsc_destination_column,
        fsc_value_column,
    )

    result = trimmed_ra.copy()
    surcharge_values: list[object] = []
    for _, row in result.iterrows():
        if not _lane_eligible_for_fuel_surcharge(row[tab_column]):
            surcharge_values.append(pd.NA)
            continue

        surcharge_values.append(_lookup_fsc_value(row, fsc_lookup, origin_columns, destination_columns))

    currency_idx = list(result.columns).index(currency_column)
    result.insert(currency_idx + 1, surcharge_column, surcharge_values)
    result.loc[result[surcharge_column].isna(), currency_column] = pd.NA

    if result[surcharge_column].notna().any():
        if fsc_file is None or fsc_sheet is None:
            raise ValueError("fsc_file and fsc_sheet are required to update Valid to dates")
        period_label = fsc_period_label_from_value_column(fsc_value_column)
        valid_to_date = find_fsc_period_valid_to_date(fsc_file, fsc_sheet, period_label)
        result = _apply_valid_to_updates(result, surcharge_column, valid_to_date)

    return result


def run_convert(
    *,
    auto: bool = False,
    ra_file: Path | None = None,
    fsc_file: Path | None = None,
    fsc_sheet: str | None = None,
    output_path: Path | None = None,
) -> ConvertResult:
    ensure_workspace_dirs()

    ra_files = list_excel_files(INPUT_RA_DIR)
    fsc_files = list_excel_files(INPUT_FSC_DIR)

    selected_ra_file = ra_file or select_input_file(ra_files, INPUT_RA_DIR, auto=auto, label="RA")
    selected_fsc_file = fsc_file or select_input_file(fsc_files, INPUT_FSC_DIR, auto=auto, label="FSC")

    if not selected_ra_file.exists():
        raise FileNotFoundError(f"RA file not found: {selected_ra_file}")
    if not selected_fsc_file.exists():
        raise FileNotFoundError(f"FSC file not found: {selected_fsc_file}")

    print(f"\nLoading RA tab '{RA_SHEET_NAME}' from {selected_ra_file.name}...")
    ra_df = ra_rate_card_to_df(selected_ra_file)
    print(f"  RA dataframe: {len(ra_df)} rows, {len(ra_df.columns)} columns")

    workbook = pd.ExcelFile(selected_fsc_file)
    selected_fsc_sheet = fsc_sheet or select_fsc_sheet(
        selected_fsc_file,
        workbook.sheet_names,
        auto=auto,
    )
    if selected_fsc_sheet not in workbook.sheet_names:
        raise ValueError(f"Sheet not found in {selected_fsc_file.name}: {selected_fsc_sheet}")

    print(f"\nLoading FSC tab '{selected_fsc_sheet}' from {selected_fsc_file.name}...")
    fsc_df = fsc_tab_to_df(selected_fsc_file, selected_fsc_sheet)
    print(f"  FSC dataframe: {len(fsc_df)} rows, {len(fsc_df.columns)} columns")

    ra_output = PROCESSING_DIR / f"{sanitize_output_stem(selected_ra_file.stem)}_rate_card.xlsx"
    fsc_output = (
        PROCESSING_DIR
        / f"{sanitize_output_stem(selected_fsc_file.stem)}_{sanitize_output_stem(selected_fsc_sheet)}.xlsx"
    )

    save_dataframe(ra_df, ra_output, sheet_name=RA_SHEET_NAME)
    save_dataframe(fsc_df, fsc_output, sheet_name=selected_fsc_sheet)

    print(f"\nSaved RA dataframe to:  {ra_output}")
    print(f"Saved FSC dataframe to: {fsc_output}")

    print("\nApplying fuel surcharge lookup and trimming RA columns...")
    ra_with_fsc = apply_fuel_surcharge_from_fsc(
        ra_df,
        fsc_df,
        fsc_file=selected_fsc_file,
        fsc_sheet=selected_fsc_sheet,
    )
    fsc_value_column = fsc_value_column_before_difference_mom(fsc_df)
    surcharge_column = fuel_surcharge_column_name(fsc_value_column)
    period_label = fsc_period_label_from_value_column(fsc_value_column)
    eligible_count = int(ra_with_fsc["Tab"].map(_lane_eligible_for_fuel_surcharge).sum())
    matched_count = int(ra_with_fsc[surcharge_column].notna().sum())
    valid_to_date = find_fsc_period_valid_to_date(selected_fsc_file, selected_fsc_sheet, period_label)
    print(
        f"  Result: {len(ra_with_fsc)} rows, {len(ra_with_fsc.columns)} columns "
        f"({matched_count} of {eligible_count} '+ Fuel Rates' lanes matched in FSC)"
    )
    print(f"  Valid to updated to {valid_to_date} for {matched_count} lanes with fuel surcharge ({period_label})")

    final_output = output_path or (
        OUTPUT_DIR / f"{sanitize_output_stem(selected_ra_file.stem)}_with_fuel_surcharge.xlsx"
    )
    save_formatted_rate_card(
        ra_with_fsc,
        final_output,
        source_ra_file=selected_ra_file,
        sheet_name=RA_SHEET_NAME,
    )
    print(f"\nSaved final RA output to: {final_output}")

    return ConvertResult(
        ra_file=selected_ra_file,
        fsc_file=selected_fsc_file,
        fsc_sheet=selected_fsc_sheet,
        ra_processing_path=ra_output,
        fsc_processing_path=fsc_output,
        output_path=final_output,
        matched_lane_count=matched_count,
        total_lane_count=len(ra_with_fsc),
    )


def main() -> int:
    try:
        run_convert()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
