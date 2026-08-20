"""
ONE SCRIPT that does BOTH:

(A) Hourly-shape comparison between two folders for years 2025 and 2030
    (per common file):
      - full-year hourly metrics: MAE, RMSE, max abs diff, mean diff, energy sum diff
      - average 24-hour shape (mean by hour-of-day) + difference
      - full-year time-series plot
      - ONE combined monthly figure (12 subplots) comparing each month within the year
      - optional representative N-day window plots
      - outputs:
          * comparison_summary_by_file_year.csv
          * optional aligned hourly CSVs
          * PNG plots

(B) Annual TWh by country for ALL YEARS 2016–2050 from both folders
    written to ONE Excel file with ONE sheet having columns:
      Country, Year, Gross_Demand_CSVs_TWh, Gross_demand_with_peak_load_TWh

Fixed: Replaced deprecated applymap() with map() for pandas 2.2.0+ compatibility.
"""

from __future__ import annotations

from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# -----------------------------
# USER SETTINGS (EDIT AS NEEDED)
# -----------------------------
FOLDER_A = Path(r"Z:\EGP\EP\Model Runs\2025-12 Planning Case\Inputs\Demand\Gross Demand")
FOLDER_B = Path(r"Z:\EGP\EP\Model Runs\2025-12 Planning Case\Inputs\Demand\Netted Demand")

# These become column names in the annual TWh Excel output
ANNUAL_COL_A = "Gross_Demand_CSVs_TWh"
ANNUAL_COL_B = "Gross_demand_with_peak_load_TWh"

TARGET_FILENAME = None  # e.g. "AlbaniaLoad.csv" to run just one file; None = all common files

# Hourly comparison years
YEARS_TO_ANALYZE = [2026, 2050]

# Annual TWh years
ANNUAL_YEARS = list(range(2016, 2051))  # inclusive 2016..2050

# Representative period inside each year for time-series plots (optional)
MAKE_WINDOW_PLOTS = True
WINDOW_START_BY_YEAR = {
    # If omitted for a year -> script chooses earliest overlapping timestamp in that year
    # 2025: "2025-01-01",
    # 2030: "2030-01-01",
}
WINDOW_N_DAYS = 7  # e.g. 7 or 30

# Output folder (created if doesn't exist)
OUTPUT_DIR = Path(r".\demand_compare_output")

# Annual TWh Excel output
ANNUAL_TWH_XLSX = OUTPUT_DIR / "annual_demand_twh_by_country.xlsx"
ANNUAL_TWH_SHEET = "Annual_TWh_2016_2050"

# Parsing options
CSV_HAS_HEADER = False  # many model-run CSVs are headerless
CSV_DELIMITER = ","
TIMEZONE = None  # e.g. "UTC" to localize; otherwise None

# Units:
# If your demand values are MW, leave as 1.0
# If kW, set 0.001 (kW -> MW)
# If GW, set 1000.0 (GW -> MW)
UNIT_SCALE_TO_MW = 1.0

# Output controls (can create large files)
WRITE_ALIGNED_FULL_YEAR_CSV = False  # full-year aligned hourly can be large
WRITE_ALIGNED_WINDOW_CSV = True

# Plot styling
FIG_DPI = 140
SHOW_PLOTS = False  # set True to display interactively


# -----------------------------
# Helpers: IO + parsing
# -----------------------------
def _is_number_like(x: str) -> bool:
    try:
        float(x)
        return True
    except Exception:
        return False


def _hour_col_to_int(colname: str) -> int | None:
    """
    Accept: 1..24, H1..H24, Hour1..Hour24, hour_1..hour_24, etc.
    Return hour integer 1..24 or None.
    """
    s = str(colname).strip().lower()

    if s.isdigit():
        h = int(s)
        return h if 1 <= h <= 24 else None

    m = re.match(r"^(h|hour)[\s_]*(\d{1,2})$", s)
    if m:
        h = int(m.group(2))
        return h if 1 <= h <= 24 else None

    return None


def read_tabular(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        if CSV_HAS_HEADER:
            return pd.read_csv(path, sep=CSV_DELIMITER)
        return pd.read_csv(path, sep=CSV_DELIMITER, header=None)

    if suffix in {".xlsx", ".xls"}:
        # .xlsx requires openpyxl installed
        return pd.read_excel(path)

    raise ValueError(f"Unsupported file type: {path.name}")


def parse_to_hourly_series(df: pd.DataFrame, series_name: str) -> pd.Series:
    """
    Convert a dataframe into an hourly time series.
    Supports:
      A) Already-hourly: has a datetime column + one numeric demand column
      B) Daily-wide: (Year, Month, Day) + 24 hourly columns
      C) Headerless daily-wide numeric: first 3 cols are Year,Month,Day then 24 values
    """
    df = df.copy()

    # If headerless (0..N columns), rename to col0..colN
    if all(isinstance(c, int) for c in df.columns):
        df.columns = [f"col{c}" for c in df.columns]

    df.columns = [str(c).strip() for c in df.columns]
    lower = {c.lower(): c for c in df.columns}

    # Case A: datetime column exists
    dt_col = None
    for key in ("date and time", "datetime", "timestamp", "time", "date_time", "date"):
        if key in lower:
            dt_col = lower[key]
            break

    if dt_col is not None:
        tmp = df.copy()
        tmp[dt_col] = pd.to_datetime(tmp[dt_col], errors="coerce")
        tmp = tmp.dropna(subset=[dt_col]).sort_values(dt_col)

        # demand column: first numeric not dt_col
        numeric_cols = [c for c in tmp.columns if c != dt_col and pd.api.types.is_numeric_dtype(tmp[c])]
        if not numeric_cols:
            for c in tmp.columns:
                if c == dt_col:
                    continue
                tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
            numeric_cols = [c for c in tmp.columns if c != dt_col and pd.api.types.is_numeric_dtype(tmp[c])]

        if not numeric_cols:
            raise ValueError("Could not find a numeric demand column in datetime-shaped file.")

        dcol = numeric_cols[0]
        s = tmp.set_index(dt_col)[dcol].astype(float)
        s.name = series_name
        s = s[~s.index.duplicated(keep="first")].sort_index()
        if TIMEZONE:
            s.index = s.index.tz_localize(TIMEZONE, nonexistent="shift_forward", ambiguous="NaT")
        return s

    # Case B/C: Year/Month/Day
    ycol = lower.get("year")
    mcol = lower.get("month")
    dcol = lower.get("day")

    if ycol and mcol and dcol:
        base_cols = [ycol, mcol, dcol]
        rest = [c for c in df.columns if c not in base_cols]
    else:
        cols = list(df.columns)
        if len(cols) < 27:
            raise ValueError("Not enough columns to be Year/Month/Day + 24 hours.")

        c0, c1, c2 = cols[0], cols[1], cols[2]
        probe = df[[c0, c1, c2]].head(50).astype(str)
        # FIX: Replace applymap() with map()
        ok = probe.map(_is_number_like).mean().mean() > 0.9
        if not ok:
            raise ValueError("Could not detect Year/Month/Day columns or a datetime column.")

        ycol, mcol, dcol = c0, c1, c2
        base_cols = [ycol, mcol, dcol]
        rest = cols[3:]

    # Interpret rest as 24 hour columns or first 24 values after day
    hour_map = {}
    hour_cols = []
    for c in rest:
        h = _hour_col_to_int(c)
        if h is not None:
            hour_map[c] = h
            hour_cols.append(c)

    if len(hour_cols) >= 24:
        hour_cols = sorted(hour_cols, key=lambda c: hour_map[c])[:24]
    else:
        hour_cols = rest[:24]
        hour_map = {c: i + 1 for i, c in enumerate(hour_cols)}

    tmp = df[base_cols + hour_cols].copy()

    tmp[ycol] = pd.to_numeric(tmp[ycol], errors="coerce").astype("Int64")
    tmp[mcol] = pd.to_numeric(tmp[mcol], errors="coerce").astype("Int64")
    tmp[dcol] = pd.to_numeric(tmp[dcol], errors="coerce").astype("Int64")
    tmp = tmp.dropna(subset=[ycol, mcol, dcol])

    tmp["date"] = pd.to_datetime(
        dict(year=tmp[ycol].astype(int), month=tmp[mcol].astype(int), day=tmp[dcol].astype(int)),
        errors="coerce",
    )
    tmp = tmp.dropna(subset=["date"])

    for c in hour_cols:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce")

    long = tmp.melt(id_vars=["date"], value_vars=hour_cols, var_name="hour_col", value_name=series_name)
    long["hour"] = long["hour_col"].map(hour_map).astype(int)
    long["timestamp"] = long["date"] + pd.to_timedelta(long["hour"] - 1, unit="h")

    s = long.set_index("timestamp")[series_name].astype(float).sort_index()
    s = s[~s.index.duplicated(keep="first")]
    if TIMEZONE:
        s.index = s.index.tz_localize(TIMEZONE, nonexistent="shift_forward", ambiguous="NaT")
    return s


# -----------------------------
# Hourly comparison helpers
# -----------------------------
def filter_year(s: pd.Series, year: int) -> pd.Series:
    if s.empty:
        return s
    return s.loc[s.index.year == year]


def align_inner(s_a: pd.Series, s_b: pd.Series, name_a: str, name_b: str) -> pd.DataFrame:
    return pd.concat([s_a.rename(name_a), s_b.rename(name_b)], axis=1, join="inner").dropna().sort_index()


def compute_metrics(aligned: pd.DataFrame, col_a: str, col_b: str) -> dict:
    diff = aligned[col_a] - aligned[col_b]
    mae = float(diff.abs().mean())
    rmse = float(np.sqrt((diff**2).mean()))
    max_abs = float(diff.abs().max())
    mean_diff = float(diff.mean())

    sum_a = float(aligned[col_a].sum())
    sum_b = float(aligned[col_b].sum())
    sum_diff = float(diff.sum())

    return {
        "n_points": int(len(aligned)),
        "mae": mae,
        "rmse": rmse,
        "max_abs_diff": max_abs,
        "mean_diff": mean_diff,
        "sum_a": sum_a,
        "sum_b": sum_b,
        "sum_diff_a_minus_b": sum_diff,
        "pct_sum_diff_vs_b": (sum_diff / sum_b * 100.0) if sum_b != 0 else np.nan,
    }


def choose_window_in_year(
    aligned_year: pd.DataFrame, year: int, n_days: int
) -> tuple[pd.DataFrame, pd.Timestamp, pd.Timestamp]:
    if aligned_year.empty:
        return aligned_year, pd.NaT, pd.NaT

    start = WINDOW_START_BY_YEAR.get(year)
    if start is None:
        start_ts = aligned_year.index.min()
    else:
        start_ts = pd.to_datetime(start)
        if start_ts < aligned_year.index.min():
            start_ts = aligned_year.index.min()
        if start_ts > aligned_year.index.max():
            start_ts = aligned_year.index.min()

    end_ts = start_ts + pd.Timedelta(days=n_days)
    if end_ts > aligned_year.index.max():
        end_ts = aligned_year.index.max()

    win = aligned_year.loc[start_ts:end_ts].copy()
    return win, start_ts, end_ts


# -----------------------------
# Plotting
# -----------------------------
def plot_time_series(aligned: pd.DataFrame, col_a: str, col_b: str, title: str, out_png: Path):
    fig = plt.figure(figsize=(14, 6), dpi=FIG_DPI)

    ax1 = plt.subplot(2, 1, 1)
    ax1.plot(aligned.index, aligned[col_a], label=col_a, linewidth=1.3)
    ax1.plot(aligned.index, aligned[col_b], label=col_b, linewidth=1.3, alpha=0.85)
    ax1.set_title(title)
    ax1.set_ylabel("Demand")
    ax1.grid(True, alpha=0.25)
    ax1.legend()

    ax2 = plt.subplot(2, 1, 2, sharex=ax1)
    diff = aligned[col_a] - aligned[col_b]
    ax2.plot(aligned.index, diff, color="black", linewidth=1.1, label=f"{col_a} - {col_b}")
    ax2.axhline(0, color="red", linewidth=1, alpha=0.7)
    ax2.set_title("Hourly Difference")
    ax2.set_xlabel("Timestamp")
    ax2.set_ylabel("Difference")
    ax2.grid(True, alpha=0.25)
    ax2.legend()

    plt.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    if SHOW_PLOTS:
        plt.show()
    plt.close(fig)


def plot_avg_hourly_shape(aligned_year: pd.DataFrame, col_a: str, col_b: str, title: str, out_png: Path):
    df = aligned_year[[col_a, col_b]].copy()
    df["hour"] = df.index.hour

    shape = df.groupby("hour")[[col_a, col_b]].mean()
    shape["diff"] = shape[col_a] - shape[col_b]

    fig = plt.figure(figsize=(12, 5), dpi=FIG_DPI)

    ax1 = plt.subplot(1, 2, 1)
    ax1.plot(shape.index, shape[col_a], marker="o", label=col_a)
    ax1.plot(shape.index, shape[col_b], marker="o", label=col_b)
    ax1.set_title(title + "\nAverage Hour-of-Day Shape")
    ax1.set_xlabel("Hour of day (0-23)")
    ax1.set_ylabel("Average demand")
    ax1.grid(True, alpha=0.25)
    ax1.legend()

    ax2 = plt.subplot(1, 2, 2)
    ax2.bar(shape.index, shape["diff"], color="gray")
    ax2.axhline(0, color="red", linewidth=1, alpha=0.7)
    ax2.set_title("Average Shape Difference")
    ax2.set_xlabel("Hour of day (0-23)")
    ax2.set_ylabel(f"{col_a} - {col_b}")
    ax2.grid(True, axis="y", alpha=0.25)

    plt.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    if SHOW_PLOTS:
        plt.show()
    plt.close(fig)


def plot_full_year_time_series(aligned_year: pd.DataFrame, col_a: str, col_b: str, title: str, out_png: Path):
    fig = plt.figure(figsize=(16, 6), dpi=FIG_DPI)

    ax1 = plt.subplot(2, 1, 1)
    ax1.plot(aligned_year.index, aligned_year[col_a], label=col_a, linewidth=1.0)
    ax1.plot(aligned_year.index, aligned_year[col_b], label=col_b, linewidth=1.0, alpha=0.85)
    ax1.set_title(title)
    ax1.set_ylabel("Demand")
    ax1.grid(True, alpha=0.25)
    ax1.legend()

    ax2 = plt.subplot(2, 1, 2, sharex=ax1)
    diff = aligned_year[col_a] - aligned_year[col_b]
    ax2.plot(aligned_year.index, diff, color="black", linewidth=0.9, label=f"{col_a} - {col_b}")
    ax2.axhline(0, color="red", linewidth=1, alpha=0.7)
    ax2.set_title("Hourly Difference")
    ax2.set_xlabel("Timestamp")
    ax2.set_ylabel("Difference")
    ax2.grid(True, alpha=0.25)
    ax2.legend()

    plt.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    if SHOW_PLOTS:
        plt.show()
    plt.close(fig)


def plot_monthly_segments_combined(
    aligned_year: pd.DataFrame,
    col_a: str,
    col_b: str,
    year: int,
    fname: str,
    output_dir: Path,
):
    """
    Single PNG with 12 subplots (one per month). Each subplot compares the hourly demand
    time series for that month (A vs B).
    """
    stem = Path(fname).stem
    out_combined = output_dir / f"{stem}_timeseries_{year}_MONTHLY_COMBINED.png"

    fig, axes = plt.subplots(6, 2, figsize=(18, 18), dpi=FIG_DPI, sharey=True)
    axes = axes.flatten()

    for month in range(1, 13):
        month_data = aligned_year[aligned_year.index.month == month]
        ax = axes[month - 1]
        month_name = pd.to_datetime(f"{year}-{month:02d}-01").strftime("%B")

        if month_data.empty:
            ax.set_title(f"{month_name} (No Data)")
            ax.axis("off")
            continue

        ax.plot(month_data.index, month_data[col_a], label=col_a, linewidth=1.0)
        ax.plot(month_data.index, month_data[col_b], label=col_b, linewidth=1.0, alpha=0.85)
        ax.set_title(month_name)
        ax.grid(True, alpha=0.25)

        # Keep legend small; only show on first subplot to reduce clutter
        if month == 1:
            ax.legend(fontsize=8, loc="upper right")

    fig.suptitle(f"{fname} ({year}) - Hourly Demand Comparison by Month", fontsize=16)
    plt.tight_layout(rect=[0, 0.03, 1, 0.97])
    fig.savefig(out_combined, bbox_inches="tight")
    if SHOW_PLOTS:
        plt.show()
    plt.close(fig)


# -----------------------------
# Annual TWh helpers (2016–2050)
# -----------------------------
def country_from_filename(filename: str) -> str:
    """
    Example: AlbaniaLoad.csv -> Albania
    Adjust if your naming differs.
    """
    stem = Path(filename).stem
    stem = re.sub(r"load$", "", stem, flags=re.IGNORECASE)
    return stem


def annual_twh_by_year_from_hourly_mw(s_mw: pd.Series) -> pd.Series:
    """
    Return Series indexed by year with TWh values.
    Assumes hourly MW: sum(MW) across hours => MWh, /1e6 => TWh.
    """
    if s_mw.empty:
        return pd.Series(dtype=float)
    mwh_by_year = s_mw.groupby(s_mw.index.year).sum()
    twh_by_year = mwh_by_year / 1e6
    twh_by_year.index.name = "Year"
    twh_by_year.name = "TWh"
    return twh_by_year


def build_annual_twh_excel(common_files: list[str], files_a: dict[str, Path], files_b: dict[str, Path]) -> Path:
    """
    For common country files, compute annual TWh 2016..2050 for both folders
    and write ONE Excel with ONE sheet containing:
      Country, Year, Gross_Demand_CSVs_TWh, Gross_demand_with_peak_load_TWh
    """
    rows: list[dict] = []

    for fname in common_files:
        country = country_from_filename(fname)

        try:
            s_a = parse_to_hourly_series(read_tabular(files_a[fname]), series_name="A") * UNIT_SCALE_TO_MW
            s_b = parse_to_hourly_series(read_tabular(files_b[fname]), series_name="B") * UNIT_SCALE_TO_MW
        except Exception as e:
            rows.append(
                {
                    "Country": country,
                    "Year": np.nan,
                    ANNUAL_COL_A: np.nan,
                    ANNUAL_COL_B: np.nan,
                    "Error": str(e),
                    "File": fname,
                }
            )
            continue

        twh_a = annual_twh_by_year_from_hourly_mw(s_a)
        twh_b = annual_twh_by_year_from_hourly_mw(s_b)

        for yr in ANNUAL_YEARS:
            rows.append(
                {
                    "Country": country,
                    "Year": int(yr),
                    ANNUAL_COL_A: float(twh_a.get(yr, np.nan)),
                    ANNUAL_COL_B: float(twh_b.get(yr, np.nan)),
                    "File": fname,
                }
            )

    out = pd.DataFrame(rows)

    ordered_cols = ["Country", "Year", ANNUAL_COL_A, ANNUAL_COL_B, "File"]
    extra = [c for c in out.columns if c not in ordered_cols]
    out = out[ordered_cols + extra]
    out = out.sort_values(["Country", "Year"]).reset_index(drop=True)

    ANNUAL_TWH_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ANNUAL_TWH_XLSX, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name=ANNUAL_TWH_SHEET, index=False)

    return ANNUAL_TWH_XLSX


# -----------------------------
# Folder comparison driver
# -----------------------------
def list_files(folder: Path) -> dict[str, Path]:
    allowed = {".csv", ".xlsx", ".xls"}
    return {p.name: p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in allowed}


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    files_a = list_files(FOLDER_A)
    files_b = list_files(FOLDER_B)

    if TARGET_FILENAME is not None:
        common = [TARGET_FILENAME] if (TARGET_FILENAME in files_a and TARGET_FILENAME in files_b) else []
        if not common:
            raise FileNotFoundError(f"{TARGET_FILENAME} not found in both folders.")
    else:
        common = sorted(set(files_a).intersection(set(files_b)))

    if not common:
        raise ValueError("No common CSV/XLSX/XLS filenames found between the two folders.")

    # (B) Annual TWh Excel output for 2016–2050 (for all common files)
    annual_path = build_annual_twh_excel(common_files=common, files_a=files_a, files_b=files_b)
    print(f"Annual TWh Excel written: {annual_path.resolve()}")

    # (A) Hourly comparison for 2025 and 2030
    name_a = f"A:{FOLDER_A.name}"
    name_b = f"B:{FOLDER_B.name}"

    results = []
    print(f"\nComparing {len(common)} common files for years: {YEARS_TO_ANALYZE}")

    for fname in common:
        path_a = files_a[fname]
        path_b = files_b[fname]
        stem = Path(fname).stem

        print(f"\n=== {fname} ===")
        try:
            s_a = parse_to_hourly_series(read_tabular(path_a), series_name=name_a)
            s_b = parse_to_hourly_series(read_tabular(path_b), series_name=name_b)

            for yr in YEARS_TO_ANALYZE:
                s_a_y = filter_year(s_a, yr)
                s_b_y = filter_year(s_b, yr)
                aligned_y = align_inner(s_a_y, s_b_y, name_a, name_b)

                if aligned_y.empty:
                    print(f"  {yr}: no overlapping hourly data.")
                    results.append({"file": fname, "year": yr, "error": "no overlapping hourly data in year"})
                    continue

                metrics = compute_metrics(aligned_y, name_a, name_b)
                metrics.update({"file": fname, "year": yr})
                results.append(metrics)

                print(
                    f"  {yr}: n={metrics['n_points']:,}  MAE={metrics['mae']:.3f}  RMSE={metrics['rmse']:.3f}  "
                    f"Max|diff|={metrics['max_abs_diff']:.3f}  SumDiff%={metrics['pct_sum_diff_vs_b']:.3f}%"
                )

                # Optional: write aligned full-year (can be big)
                if WRITE_ALIGNED_FULL_YEAR_CSV:
                    out_full = OUTPUT_DIR / f"{stem}_aligned_{yr}_FULLYEAR.csv"
                    aligned_y.to_csv(out_full, index=True)

                # Average hourly shape plot
                out_shape = OUTPUT_DIR / f"{stem}_avg_hourly_shape_{yr}.png"
                plot_avg_hourly_shape(
                    aligned_year=aligned_y,
                    col_a=name_a,
                    col_b=name_b,
                    title=f"{fname} ({yr})",
                    out_png=out_shape,
                )

                # Full-year time-series plot (entire year in one plot)
                out_year_ts = OUTPUT_DIR / f"{stem}_timeseries_{yr}_FULLYEAR.png"
                plot_full_year_time_series(
                    aligned_year=aligned_y,
                    col_a=name_a,
                    col_b=name_b,
                    title=f"{fname} ({yr}) - Full Year Hourly Demand Comparison",
                    out_png=out_year_ts,
                )

                # ONE combined monthly figure (12 subplots)
                plot_monthly_segments_combined(
                    aligned_year=aligned_y,
                    col_a=name_a,
                    col_b=name_b,
                    year=yr,
                    fname=fname,
                    output_dir=OUTPUT_DIR,
                )

                # Optional: representative N-day window plots
                if MAKE_WINDOW_PLOTS:
                    win, wstart, wend = choose_window_in_year(aligned_y, yr, WINDOW_N_DAYS)
                    if not win.empty:
                        out_ts = OUTPUT_DIR / f"{stem}_timeseries_{yr}_{wstart.date()}_{wend.date()}.png"
                        plot_time_series(
                            aligned=win,
                            col_a=name_a,
                            col_b=name_b,
                            title=f"{fname} ({yr}) Hourly Demand Comparison\n{wstart} to {wend}",
                            out_png=out_ts  
                        )

                        if WRITE_ALIGNED_WINDOW_CSV:
                            out_win = OUTPUT_DIR / f"{stem}_aligned_window_{yr}_{wstart.date()}_{wend.date()}.csv"
                            win.to_csv(out_win, index=True)

        except Exception as e:
            print(f"ERROR processing {fname}: {e}")
            for yr in YEARS_TO_ANALYZE:
                results.append({"file": fname, "year": yr, "error": str(e)})

    summary = pd.DataFrame(results)
    summary_out = OUTPUT_DIR / "comparison_summary_by_file_year.csv"
    summary.to_csv(summary_out, index=False)

    print(f"\nDone. Outputs written to: {OUTPUT_DIR.resolve()}")
    print(f"Summary CSV: {summary_out.resolve()}")


if __name__ == "__main__":
    main()