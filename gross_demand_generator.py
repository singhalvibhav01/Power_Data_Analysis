#!/usr/bin/env python3
"""
Create ONE consolidated hourly load CSV per label (country/zone) for 2016-2050,
matching the AlbaniaLoad.csv template: YEAR, MONTH, DAY, 1..24.

UPDATED (annual-level normalisation):
- CF profile is normalised across the *entire year* so that for each YEAR:
    sum(over all days, hours 1..24) == 1
- Annual TWh is preserved by scaling that annual-normalised profile:
    hourly_mwh = annual_fraction * annual_mwh
    (output columns are named as hourly MW in the template, but numerically this is MWh-per-hour.)

Input Excel:
  Z:\\EGP\\EP\\Model Runs\\2025-12 Planning Case\\Inputs\\Demand\\Demandbaseyears - 202606.xlsx
  Sheet: Inputs-Planning
  Names: E8:E65
  Annual TWh: V8:BD65 (assumed 2016..2050 in order)

CF profile dir:
  Z:\\EGP\\EP\\Model Runs\\2025-12 Planning Case\\Inputs\\Demand\\CSV Files

Output dir (one file per label, NOT per year):
  Z:\\EGP\\EP\\Model Runs\\2025-12 Planning Case\\Inputs\\Demand\\Gross Demand

Row rules:
- For {Denmark, Italy, Norway, Sweden, Serbia and Montenegro, United Kingdom}:
  * skip Excel rows 16, 24, 31, 35, 39, 42
  * ONLY generate for Excel rows 43..65

Validation reports:
- cf_validation_detailed.csv: days with zero values in CF profiles (hours 1..24)
- demand_validation_detailed.csv: days with negative values in gross demand outputs (hours 1..24)

Notes / assumptions:
- If CF profile does not contain the target year, the script uses YEAR=2016 from the CF profile
  as a shape and then rewrites the YEAR column to the target year.
- Annual-level normalisation preserves seasonal/day-to-day variations present in the CF profile
  (unlike day-level normalisation which forced equal energy per day).
"""

from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path
from typing import Dict, List

import pandas as pd

# -------------------------
# Excel layout
# -------------------------
LABEL_COL = "E"
ROW_START = 8
ROW_END = 66

DATA_COL_START = "V"
DATA_COL_END = "BD"
FIRST_YEAR = 2016
LAST_YEAR = 2051  # inclusive (matches your existing code / range check)

# -------------------------
# Hourly/profile constants
# -------------------------
DATE_COLS = ["YEAR", "MONTH", "DAY"]
HOURS = [str(h) for h in range(1, 25)]
MWH_PER_TWH = 1_000_000.0

# -------------------------
# Row rules
# -------------------------
RESTRICTED_COUNTRIES = {
    "DENMARK",
    "ITALY",
    "NORWAY",
    "SWEDEN",
    "SERBIA AND MONTENEGRO",
    "UNITED KINGDOM",
}
EXCLUDED_ROWS_FOR_RESTRICTED = {16, 25, 32, 36, 40, 43}
RESTRICTED_ALLOWED_RANGE = range(44, 67)  # 43..65 inclusive


# -------------------------
# CLI / Paths
# -------------------------
def _env(name: str, default: str) -> str:
    """Retrieve environment variable or return default."""
    return os.environ.get(name, default)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments with environment variable fallbacks."""
    script_dir = Path(__file__).resolve().parent

    p = argparse.ArgumentParser(
        description="Generate consolidated gross demand and validation reports for CF profiles and demand outputs."
    )

    p.add_argument(
        "--base-dir",
        default=_env("DEMAND_BASE_DIR", str(script_dir)),
        help="Base directory for resolving relative paths (default: script directory, or env DEMAND_BASE_DIR).",
    )

    p.add_argument(
        "--excel-path",
        default=_env(
            "DEMAND_EXCEL_PATH",
            r"\\Info.corp\emea\Markit\UK\Shared\EGPData\EGP\EP\Model Runs\2025-12 Planning Case\Inputs\Demand\Demandbaseyears - 202606.xlsx",
        ),
        help="Path to the input Excel file (env DEMAND_EXCEL_PATH).",
    )
    p.add_argument(
        "--sheet-name",
        default=_env("DEMAND_SHEET_NAME", "Inputs-Planning"),
        help="Excel sheet name (env DEMAND_SHEET_NAME).",
    )

    p.add_argument(
        "--cf-dir",
        default=_env(
            "DEMAND_CF_DIR",
            r"\\Info.corp\emea\Markit\UK\Shared\EGPData\EGP\EP\Model Runs\2025-12 Planning Case\Inputs\Demand\CSV Files",
        ),
        help="Directory containing CF load profiles (env DEMAND_CF_DIR).",
    )
    p.add_argument(
        "--output-dir",
        default=_env(
            "DEMAND_OUTPUT_DIR",
            r"\\Info.corp\emea\Markit\UK\Shared\EGPData\EGP\EP\Model Runs\2025-12 Planning Case\Inputs\Demand\Gross Demand",
        ),
        help="Output directory for consolidated load CSVs and summaries (env DEMAND_OUTPUT_DIR).",
    )

    p.add_argument(
        "--resolve-relative-to-base",
        action="store_true",
        default=(_env("DEMAND_RESOLVE_RELATIVE_TO_BASE", "0") == "1"),
        help="If set (or env DEMAND_RESOLVE_RELATIVE_TO_BASE=1), resolve non-absolute paths under --base-dir.",
    )

    return p.parse_args()


def resolve_path(pth: str | Path, base_dir: Path, resolve_relative: bool) -> Path:
    """Resolve a path, optionally relative to base_dir if not absolute."""
    pth = Path(pth)
    if resolve_relative and not pth.is_absolute():
        return (base_dir / pth).resolve()
    return pth


# -------------------------
# Helpers
# -------------------------
def setup_logger() -> logging.Logger:
    """Configure and return a logger instance."""
    logger = logging.getLogger("gross_demand_consolidated")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        h = logging.StreamHandler()
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s - %(message)s"))
        logger.addHandler(h)
    return logger


def excel_col_to_index(col: str) -> int:
    """Convert Excel column letter(s) to 0-based index (e.g., 'A' -> 0, 'Z' -> 25, 'AA' -> 26)."""
    col = col.strip().upper()
    n = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Bad Excel column: {col}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def norm_date_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize date column names to uppercase (YEAR, MONTH, DAY)."""
    df = df.copy()
    df.columns = [
        c.upper() if str(c).lower() in ["year", "month", "day"] else c
        for c in df.columns
    ]
    return df


def require_columns(df: pd.DataFrame, required: List[str], label: str) -> None:
    """Raise ValueError if required columns are missing from DataFrame."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{label}: missing columns {missing}. Available: {list(df.columns)}")


def base_country_from_label(label: str) -> str:
    """Extract base country name from label (e.g., 'DENMARK_WEST' -> 'DENMARK')."""
    u = label.strip().upper()
    for c in RESTRICTED_COUNTRIES:
        if u.startswith(c):
            return c
    return u


def should_process_excel_row(excel_row: int, label: str) -> bool:
    """Determine if an Excel row should be processed based on row rules."""
    base = base_country_from_label(label)
    if base in RESTRICTED_COUNTRIES:
        if excel_row in EXCLUDED_ROWS_FOR_RESTRICTED:
            return False
        return excel_row in RESTRICTED_ALLOWED_RANGE
    return True


def read_row_label(sheet_df: pd.DataFrame, excel_row: int) -> str:
    """Read the label (country/zone name) from a specific Excel row."""
    r = excel_row - 1  # pandas is 0-based, excel rows are 1-based
    c = excel_col_to_index(LABEL_COL)
    v = sheet_df.iat[r, c]
    if pd.isna(v):
        return ""
    return str(v).strip()


def read_row_annual_twh(
    sheet_df: pd.DataFrame, excel_row: int, logger: logging.Logger
) -> Dict[int, float]:
    """Read annual TWh values for all years (2016-2050) from a specific Excel row."""
    r = excel_row - 1
    c0 = excel_col_to_index(DATA_COL_START)
    c1 = excel_col_to_index(DATA_COL_END)

    n_years = c1 - c0 + 1
    years = list(range(FIRST_YEAR, FIRST_YEAR + n_years))

    if years[-1] != LAST_YEAR:
        logger.warning(
            f"Column range {DATA_COL_START}:{DATA_COL_END} implies years {years[0]}..{years[-1]} "
            f"(expected {FIRST_YEAR}..{LAST_YEAR}). Proceeding with implied years."
        )

    out: Dict[int, float] = {}
    for j, y in enumerate(years):
        val = sheet_df.iat[r, c0 + j]
        if pd.isna(val):
            continue
        try:
            out[int(y)] = float(val)
        except Exception:
            continue
    return out


def find_cf_profile(cf_dir: Path, label: str) -> Path:
    """Locate the CF profile CSV file for a given label."""
    candidates = [
        cf_dir / f"{label}Load2016.csv",
        cf_dir / f"{label}Load.csv",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"Missing CF profile for '{label}'. Tried:\n  - "
        + "\n  - ".join(str(x) for x in candidates)
    )


# -------------------------
# Annual-level normalisation + shaping
# -------------------------
def normalize_annual_hourly_weights(profile_df: pd.DataFrame, year: int) -> pd.DataFrame:
    """
    Return a dataframe for the given YEAR where the sum of ALL hourly values across
    the entire year equals 1.0:

        sum_{days} sum_{h=1..24} weight(day,h) == 1

    This matches: hourly_profile = hourly_mwh / annual_mwh.
    
    Args:
        profile_df: DataFrame with YEAR, MONTH, DAY, 1..24 columns
        year: Target year to normalize
    
    Returns:
        DataFrame for the target year with normalized hourly weights
    
    Raises:
        ValueError: If year not found or annual sum is non-positive
    """
    profile_df = norm_date_cols(profile_df)
    require_columns(profile_df, DATE_COLS + HOURS, "CF profile")

    yr = profile_df[profile_df["YEAR"] == year].copy()
    if yr.empty:
        raise ValueError(f"CF profile missing YEAR={year}")

    annual_sum = float(yr[HOURS].to_numpy(dtype="float64").sum())
    if annual_sum <= 0.0:
        raise ValueError(f"CF profile YEAR={year} has non-positive annual sum: {annual_sum}")

    yr[HOURS] = yr[HOURS] / annual_sum
    return yr


def shape_twh_to_hourly_mw_annual(weights_df: pd.DataFrame, twh: float) -> pd.DataFrame:
    """
    Apply annual-normalised weights directly to annual MWh.

    If weights sum to 1 over the year, then:
        hourly = weight * annual_mwh
    
    Args:
        weights_df: DataFrame with normalized weights (sum to 1 over year)
        twh: Annual demand in TWh
    
    Returns:
        DataFrame with hourly demand values in MWh
    """
    require_columns(weights_df, DATE_COLS + HOURS, "weights")

    out = weights_df[DATE_COLS].copy()
    annual_mwh = float(twh) * MWH_PER_TWH
    out[HOURS] = weights_df[HOURS].to_numpy(dtype="float64") * annual_mwh
    return out


def ensure_template_order(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure output DataFrame has correct column order: YEAR, MONTH, DAY, 1..24."""
    df = norm_date_cols(df)
    require_columns(df, DATE_COLS + HOURS, "output dataframe")
    return df[DATE_COLS + HOURS].copy()


def build_consolidated_load_for_label(
    *,
    label: str,
    annual_twh: Dict[int, float],
    cf_df: pd.DataFrame,
    logger: logging.Logger,
) -> tuple[pd.DataFrame, List[dict]]:
    """
    For each year with an annual TWh value, build a shaped hourly load.
    Uses annual-level normalisation (sum all hours in shape year = 1).
    
    Args:
        label: Country/zone label
        annual_twh: Dictionary mapping year -> TWh demand
        cf_df: DataFrame with CF profile (YEAR, MONTH, DAY, 1..24)
        logger: Logger instance
    
    Returns:
        Tuple of (consolidated_df, summary_rows):
        - consolidated_df: Complete hourly load for all years
        - summary_rows: List of summary statistics per year
    
    Raises:
        ValueError: If no yearly data can be produced
    """
    cf_df = norm_date_cols(cf_df)
    require_columns(cf_df, DATE_COLS + HOURS, f"CF profile for {label}")

    cf_years = set(int(x) for x in cf_df["YEAR"].dropna().unique().tolist())
    parts: List[pd.DataFrame] = []
    summary_rows: List[dict] = []

    for y in range(FIRST_YEAR, LAST_YEAR + 1):
        if y not in annual_twh:
            continue

        twh = float(annual_twh[y])
        shape_year = y if y in cf_years else 2016

        weights = normalize_annual_hourly_weights(cf_df, year=shape_year)
        hourly = shape_twh_to_hourly_mw_annual(weights, twh)

        # If we used 2016 shape, rewrite the YEAR column to the target year
        if shape_year != y:
            hourly["YEAR"] = y

        hourly = ensure_template_order(hourly)
        parts.append(hourly)

        achieved_twh = hourly[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH
        summary_rows.append(
            {
                "Label": label,
                "Year": y,
                "AnnualDemand_TWh_input": twh,
                "AnnualDemand_TWh_achieved": achieved_twh,
                "ShapeYearUsed": shape_year,
                "CFHasTargetYear": (y in cf_years),
                "RowsInShapeYear": int(len(weights)),  # days (rows) in that shape year
            }
        )

    if not parts:
        raise ValueError(f"No yearly data produced for label={label}")

    out = pd.concat(parts, ignore_index=True)
    out = out.sort_values(DATE_COLS).reset_index(drop=True)
    out = ensure_template_order(out)
    return out, summary_rows


# -------------------------
# CF validation (zeros)
# -------------------------
def validate_cf_profile_for_zeros(
    cf_df: pd.DataFrame,
    label: str,
    profile_path: Path,
    logger: logging.Logger,
) -> List[dict]:
    """
    Scan a CF profile and collect detailed records for any rows (days) that have
    zero values in any hourly columns 1..24.
    
    Args:
        cf_df: DataFrame with CF profile (YEAR, MONTH, DAY, 1..24)
        label: Country/zone label for identification
        profile_path: Path to the CF profile file
        logger: Logger instance for warnings
    
    Returns:
        List of dictionaries, one per day with zero values, containing:
        - Label, ProfileFile, Year, Month, Day
        - ZeroHours (comma-separated list of affected hours)
        - TotalZeros (count of zero values in that day)
        - AllHoursZero (boolean: True if all 24 hours are zero)
    """
    cf_df = norm_date_cols(cf_df)
    require_columns(cf_df, DATE_COLS + HOURS, f"CF profile validation for {label}")

    # Ensure we can compare numeric zeros robustly
    hourly_numeric = cf_df[HOURS].apply(pd.to_numeric, errors="coerce")

    issues: List[dict] = []
    # Iterate only over rows that have at least one exact zero (ignoring NaNs).
    zero_any = (hourly_numeric == 0.0).any(axis=1)
    if not bool(zero_any.any()):
        return issues

    problem_df = cf_df.loc[zero_any, DATE_COLS].copy()
    problem_hourly = hourly_numeric.loc[zero_any, HOURS]

    for i in range(len(problem_df)):
        row_date = problem_df.iloc[i]
        vals = problem_hourly.iloc[i].to_numpy(dtype="float64")
        zero_mask = vals == 0.0
        zero_hours = [HOURS[j] for j, is_zero in enumerate(zero_mask) if is_zero]
        zero_count = int(len(zero_hours))

        issues.append(
            {
                "Label": label,
                "ProfileFile": profile_path.name,
                "Year": int(row_date["YEAR"]) if not pd.isna(row_date["YEAR"]) else pd.NA,
                "Month": int(row_date["MONTH"]) if not pd.isna(row_date["MONTH"]) else pd.NA,
                "Day": int(row_date["DAY"]) if not pd.isna(row_date["DAY"]) else pd.NA,
                "ZeroHours": ",".join(zero_hours),
                "TotalZeros": zero_count,
                "AllHoursZero": (zero_count == 24),
            }
        )

    logger.warning(
        f"CF profile '{label}' ({profile_path.name}) contains {len(issues)} day(s) with one or more 0 values."
    )
    return issues


# -------------------------
# Gross demand validation (negatives)
# -------------------------
def validate_gross_demand_for_negatives(
    demand_df: pd.DataFrame,
    label: str,
    output_path: Path,
    logger: logging.Logger,
) -> List[dict]:
    """
    Scan a gross demand output file and collect detailed records for any rows (days)
    that have negative values in any hourly columns 1..24.
    
    This validation ensures that the generated demand profiles do not contain
    physically impossible negative demand values, which could indicate data quality
    issues or calculation errors in the normalization/shaping process.
    
    Args:
        demand_df: DataFrame containing the gross demand data with YEAR, MONTH, DAY, 1..24 columns
        label: Country/zone label for identification in the report
        output_path: Path to the output file being validated
        logger: Logger instance for warnings
    
    Returns:
        List of dictionaries, one per day with negative values, containing:
        - Label, OutputFile, Year, Month, Day
        - NegativeHours (comma-separated list of affected hours)
        - TotalNegatives (count of negative values in that day)
        - MinValue (most negative value found in that day)
    """
    demand_df = norm_date_cols(demand_df)
    require_columns(demand_df, DATE_COLS + HOURS, f"Gross demand validation for {label}")

    # Convert hourly columns to numeric, coercing errors to NaN
    hourly_numeric = demand_df[HOURS].apply(pd.to_numeric, errors="coerce")

    issues: List[dict] = []
    
    # Identify rows with at least one negative value (excluding NaNs)
    negative_any = (hourly_numeric < 0.0).any(axis=1)
    if not bool(negative_any.any()):
        return issues

    problem_df = demand_df.loc[negative_any, DATE_COLS].copy()
    problem_hourly = hourly_numeric.loc[negative_any, HOURS]

    for i in range(len(problem_df)):
        row_date = problem_df.iloc[i]
        vals = problem_hourly.iloc[i].to_numpy(dtype="float64")
        negative_mask = vals < 0.0
        negative_hours = [HOURS[j] for j, is_neg in enumerate(negative_mask) if is_neg]
        negative_count = int(len(negative_hours))
        min_value = float(vals[negative_mask].min()) if negative_count > 0 else 0.0

        issues.append(
            {
                "Label": label,
                "OutputFile": output_path.name,
                "Year": int(row_date["YEAR"]) if not pd.isna(row_date["YEAR"]) else pd.NA,
                "Month": int(row_date["MONTH"]) if not pd.isna(row_date["MONTH"]) else pd.NA,
                "Day": int(row_date["DAY"]) if not pd.isna(row_date["DAY"]) else pd.NA,
                "NegativeHours": ",".join(negative_hours),
                "TotalNegatives": negative_count,
                "MinValue": min_value,
            }
        )

    logger.warning(
        f"Gross demand '{label}' ({output_path.name}) contains {len(issues)} day(s) with negative values."
    )
    return issues


# -------------------------
# Validation report generation
# -------------------------
def generate_validation_report(
    cf_validation_issues: List[dict],
    demand_validation_issues: List[dict],
    output_dir: Path,
    logger: logging.Logger,
) -> None:
    """
    Write validation reports for both CF profiles and gross demand outputs:
    - cf_validation_detailed.csv: days with zero values in CF profiles
    - demand_validation_detailed.csv: days with negative values in gross demand outputs
    
    Args:
        cf_validation_issues: List of CF profile zero-value issues
        demand_validation_issues: List of gross demand negative-value issues
        output_dir: Directory where validation reports will be saved
        logger: Logger instance for status messages
    """
    # CF Profile Zero-Value Report
    if not cf_validation_issues:
        logger.info("Validation: no zero values found in CF profiles (hours 1..24).")
    else:
        cf_detailed_df = pd.DataFrame(cf_validation_issues)
        cf_detailed_path = output_dir / "cf_validation_detailed.csv"
        cf_detailed_df.to_csv(cf_detailed_path, index=False)
        logger.warning(f"[VALIDATION] Wrote CF zero report: {cf_detailed_path}")
    
    # Gross Demand Negative-Value Report
    if not demand_validation_issues:
        logger.info("Validation: no negative values found in gross demand outputs (hours 1..24).")
    else:
        demand_detailed_df = pd.DataFrame(demand_validation_issues)
        demand_detailed_path = output_dir / "demand_validation_detailed.csv"
        demand_detailed_df.to_csv(demand_detailed_path, index=False)
        logger.warning(f"[VALIDATION] Wrote gross demand negative value report: {demand_detailed_path}")


# -------------------------
# Main
# -------------------------
def main() -> int:
    """
    Main processing pipeline:
    1. Read Excel file with labels and annual TWh values
    2. For each label, load CF profile and build consolidated hourly load
    3. Validate CF profiles for zero values
    4. Validate gross demand outputs for negative values
    5. Generate summary and validation reports
    
    Returns:
        Exit code (0 for success)
    """
    args = parse_args()
    logger = setup_logger()

    base_dir = Path(args.base_dir).expanduser().resolve()

    excel_path = resolve_path(args.excel_path, base_dir, args.resolve_relative_to_base)
    cf_dir = resolve_path(args.cf_dir, base_dir, args.resolve_relative_to_base)
    output_dir = resolve_path(args.output_dir, base_dir, args.resolve_relative_to_base)

    output_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"BASE_DIR    = {base_dir}")
    logger.info(f"EXCEL_PATH  = {excel_path}")
    logger.info(f"CF_DIR      = {cf_dir}")
    logger.info(f"OUTPUT_DIR  = {output_dir}")
    logger.info(f"SHEET_NAME  = {args.sheet_name}")

    # Read without headers to match your cell-coordinate extraction
    sheet_df = pd.read_excel(excel_path, sheet_name=args.sheet_name, header=None)

    yearly_summary: List[dict] = []
    validation_issues: List[dict] = []
    demand_validation_issues: List[dict] = []

    for excel_row in range(ROW_START, ROW_END + 1):
        label = read_row_label(sheet_df, excel_row)
        if not label:
            continue

        if not should_process_excel_row(excel_row, label):
            logger.info(f"Skipping row {excel_row}: '{label}' (by rule)")
            continue

        annual_twh = read_row_annual_twh(sheet_df, excel_row, logger)
        if not annual_twh:
            logger.warning(f"Row {excel_row}: '{label}' has no annual TWh. Skipping.")
            continue

        base_country = base_country_from_label(label)

        try:
            profile_path = find_cf_profile(cf_dir, label)
            cf_df = pd.read_csv(profile_path)
        except Exception as e:
            logger.warning(f"Row {excel_row}: '{label}' cannot load CF profile ({e}). Skipping.")
            continue

        # Validation: detect zeros in the CF profile (does not block outputs)
        try:
            validation_issues.extend(
                validate_cf_profile_for_zeros(
                    cf_df=cf_df,
                    label=label,
                    profile_path=profile_path,
                    logger=logger,
                )
            )
        except Exception as e:
            logger.error(f"Row {excel_row}: '{label}' CF validation failed ({e}). Continuing.")

        try:
            consolidated_df, summary_rows = build_consolidated_load_for_label(
                label=label,
                annual_twh=annual_twh,
                cf_df=cf_df,
                logger=logger,
            )
        except Exception as e:
            logger.warning(
                f"Row {excel_row}: '{label}' failed to build consolidated load ({e}). Skipping."
            )
            continue

        out_path = output_dir / f"{label}Load.csv"
        consolidated_df.to_csv(out_path, index=False)
        logger.info(f"[OK] Saved consolidated load: {out_path}")

        # Validation: detect negative values in the generated gross demand file
        try:
            demand_validation_issues.extend(
                validate_gross_demand_for_negatives(
                    demand_df=consolidated_df,
                    label=label,
                    output_path=out_path,
                    logger=logger,
                )
            )
        except Exception as e:
            logger.error(f"Row {excel_row}: '{label}' demand validation failed ({e}). Continuing.")

        for r in summary_rows:
            r.update(
                {
                    "ExcelRow": excel_row,
                    "BaseCountry": base_country,
                    "CFProfileFile": profile_path.name,
                    "OutputFile": str(out_path),
                }
            )
            yearly_summary.append(r)

    if yearly_summary:
        summary_df = pd.DataFrame(yearly_summary).sort_values(["BaseCountry", "Label", "Year"])
        summary_path = output_dir / "yearly_summary.csv"
        summary_df.to_csv(summary_path, index=False)
        logger.info(f"[OK] Saved yearly summary: {summary_path}")
    else:
        logger.warning("No outputs generated; yearly_summary.csv not created.")

    # Validation reports at the end
    generate_validation_report(validation_issues, demand_validation_issues, output_dir, logger)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())