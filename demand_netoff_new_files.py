#!/usr/bin/env python3
from __future__ import annotations
print(
    "=== YEARLY_NETOFF_GENERATOR (MODEL B) LOADING (v2026-02-11, annual-normalized-Diego, cf_key_map, embedded_pct_only, indent_fix, cap_rows_440_447_italy_exception, net_validation_zero_and_negative) ===",
    flush=True,
)
"""
Demand Netoff Generator (MODEL B): Net shaped with Diego profile; Embedded computed as residual.

This version includes:
- Annual-level normalization (entire year sums to 1 across all days and 24 hours)
- Annual energy preserved by direct scaling of annual-normalized profile
- Date alignment (inner join on YEAR/MONTH/DAY)
- Fix B (cap NetMW ≤ GrossMW and redistribute to meet annual target)
- Zero rule (gross-net rounds to 0.00 => embedded forced to 0 and Net=Gross)
- All-countries default (ALL)
- Embedded post-redistribution to Diego annual profile (annual TWh preserved)
- CF-key mapping written to output/cf_key_mapping.json
- Spain country code forced to 'SP' for CF key construction
- Embedded target reading: ONLY reads DEMAND row "Embedded" with unit "%" and uses YEAR header row = 1.
  Embedded target TWh by year = COUNTRY gross TWh by year * (embedded_percent/100).

Fix in this version:
- Correct indentation in process_zone(): the main net-shaping + FixB block now runs for ALL years,
  not only for the final else-branch of the target scaling logic.

Additional change in this version:
- Installed capacity (embedded) reading changed to fixed row ranges:
        - SolarPV: rows 486-493 (Excel)
        - Wind: rows 486-489 onshore + 490-493 offshore (Excel)
  Values assumed to be in GW and converted to MW by *1000.
- Yearly summary now includes Wind Capacity (MW) and Solar Capacity (MW).

NORMALIZATION CHANGE (this version):
- Diego profile normalization changed from intraday-only to annual-level:
    * sum(all days, hours 1..24) == 1 for each year
    * Preserves seasonal and day-to-day variations in the Diego profile
    * hourly_mw = annual_fraction * annual_mwh

VALIDATION ENHANCEMENT (this version):
- Net load validation: detects BOTH zero values and negative values in hourly columns 1-24
- Generates net_validation_detailed.csv report
- Validation runs after each net load file is written
- Does NOT validate embedded files (only net load files)
"""

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any

import pandas as pd

# ----------------------------------------------------------------------
# Global constants
# ----------------------------------------------------------------------
DATE_COLS = ["YEAR", "MONTH", "DAY"]
HOURS = [str(h) for h in range(1, 25)]
MWH_PER_TWH = 1_000_000.0


# ----------------------------------------------------------------------
# Utility helpers
# ----------------------------------------------------------------------
def norm_date_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Upper-case YEAR/MONTH/DAY column names (if present)."""
    df = df.copy()
    df.columns = [
        c.upper() if str(c).lower() in ["year", "month", "day"] else c
        for c in df.columns
    ]
    return df


def safe_country_dirname(name: str) -> str:
    """Make a string safe for a folder name (Windows-safe)."""
    bad = '<>:"/\\|?*'
    out = "".join("_" if c in bad else c for c in str(name).strip())
    return " ".join(out.split()) or "UNKNOWN"


def require_columns(df: pd.DataFrame, required: List[str], label: str) -> None:
    """Raise a clear error if required columns are missing."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"{label}: missing required columns: {missing}. Available: {list(df.columns)}"
        )


def annual_twh_from_hourly_mw(df: pd.DataFrame) -> Dict[int, float]:
    """Annual TWh = sum(hourly MW in year) / 1,000,000."""
    df = norm_date_cols(df)
    require_columns(df, DATE_COLS + HOURS, "gross/net dataframe")
    out: Dict[int, float] = {}
    for y, g in df.groupby("YEAR"):
        out[int(y)] = g[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH
    return out


# ----------------------------------------------------------------------
# Diego normalization / shaping helpers (ANNUAL-LEVEL)
# ----------------------------------------------------------------------
def diego_annual_normalized_weights(profile_df: pd.DataFrame, year: int) -> pd.DataFrame:
    """
    Return a dataframe for the requested year where the sum of ALL hourly values
    across the entire year equals 1.0:

        sum_{days} sum_{h=1..24} weight(day,h) == 1

    This preserves seasonal and day-to-day variations in the Diego profile.
    """
    profile_df = norm_date_cols(profile_df)
    require_columns(profile_df, DATE_COLS + HOURS, "Diego profile dataframe")

    yr = profile_df[profile_df["YEAR"] == year].copy()
    if yr.empty:
        raise ValueError(f"Diego profile missing YEAR={year}")

    annual_sum = float(yr[HOURS].to_numpy(dtype="float64").sum())
    if annual_sum <= 0.0:
        raise ValueError(f"Diego profile YEAR={year} has non-positive annual sum: {annual_sum}")

    yr[HOURS] = yr[HOURS] / annual_sum
    return yr


def shape_twh_to_hourly_mw_annual(weights_df: pd.DataFrame, twh: float) -> pd.DataFrame:
    """
    Shape annual TWh into hourly MW using annual-normalized weights.

    Since the entire year's weights sum to 1, we directly scale:
        hourly_mw = weight * annual_mwh
    """
    weights_df = norm_date_cols(weights_df)
    require_columns(weights_df, DATE_COLS + HOURS, "weights dataframe")

    annual_mwh = float(twh) * MWH_PER_TWH

    out = weights_df[DATE_COLS].copy()
    out[HOURS] = weights_df[HOURS].to_numpy(dtype="float64") * annual_mwh
    return out


def zero_rule_applies(gross_twh: float, net_twh: float) -> bool:
    """If embedded (=gross-net) rounds to 0.00 at 2 decimals => force embedded=0."""
    return round(gross_twh - net_twh, 2) == 0.00


def align_year_on_dates(
    gross_y_df: pd.DataFrame,
    net_y_df: pd.DataFrame,
    logger: Optional[logging.Logger] = None,
    label: str = "",
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Inner-join gross and net on (YEAR,MONTH,DAY) to guarantee same shape.
    Prevents broadcasting errors when one side has missing/extra days.
    """
    g = gross_y_df[DATE_COLS + HOURS].copy()
    n = net_y_df[DATE_COLS + HOURS].copy()

    g["__key__"] = list(zip(g["YEAR"], g["MONTH"], g["DAY"]))
    n["__key__"] = list(zip(n["YEAR"], n["MONTH"], n["DAY"]))

    gk = set(g["__key__"])
    nk = set(n["__key__"])
    only_g = gk - nk
    only_n = nk - gk

    if logger and (only_g or only_n):
        logger.warning(
            f"{label} date mismatch: gross_only={len(only_g)} net_only={len(only_n)} "
            f"(using intersection={len(gk & nk)})"
        )

    merged = g.merge(n, on="__key__", how="inner", suffixes=("_gross", "_net"))

    g2 = merged[[f"{c}_gross" for c in DATE_COLS] + [f"{h}_gross" for h in HOURS]].copy()
    g2.columns = DATE_COLS + HOURS

    n2 = merged[[f"{c}_net" for c in DATE_COLS] + [f"{h}_net" for h in HOURS]].copy()
    n2.columns = DATE_COLS + HOURS

    return g2, n2


def cap_and_redistribute_net_to_gross(
    gross_df_y: pd.DataFrame,
    net_df_y: pd.DataFrame,
    target_net_twh: float,
    max_iters: int = 30,
    tol_mwh: float = 0.001,
) -> pd.DataFrame:
    """
    Fix B:
    1. Cap NetMW ≤ GrossMW hour-by-hour.
    2. Redistribute energy to meet annual net target where possible.
    """
    require_columns(gross_df_y, DATE_COLS + HOURS, "gross_df_y")
    require_columns(net_df_y, DATE_COLS + HOURS, "net_df_y")

    gross = gross_df_y[HOURS].to_numpy(dtype="float64")
    net = net_df_y[HOURS].to_numpy(dtype="float64")

    net = net.clip(min=0.0)
    over = net > gross
    if over.any():
        net[over] = gross[over]

    target_mwh = target_net_twh * MWH_PER_TWH
    current_mwh = net.sum()
    remaining = target_mwh - current_mwh

    it = 0
    while abs(remaining) > tol_mwh and it < max_iters:
        it += 1
        if remaining > 0:
            headroom = (gross - net).clip(min=0.0)
            total_headroom = headroom.sum()
            if total_headroom <= 0:
                break
            add = headroom * (remaining / total_headroom)
            add = add.clip(min=0.0)
            too_big = add > headroom
            if too_big.any():
                add[too_big] = headroom[too_big]
            net = net + add
        else:
            removable = net.clip(min=0.0)
            total_removable = removable.sum()
            if total_removable <= 0:
                break
            remove = removable * ((-remaining) / total_removable)
            remove = remove.clip(min=0.0)
            too_big = remove > net
            if too_big.any():
                remove[too_big] = net[too_big]
            net = net - remove

        current_mwh = net.sum()
        remaining = target_mwh - current_mwh

    out = net_df_y[DATE_COLS].copy()
    out[HOURS] = net
    return out


# ----------------------------------------------------------------------
# Net load validation (zeros + negatives) - NET FILES ONLY
# ----------------------------------------------------------------------
def validate_net_load_for_zeros_and_negatives(
    net_df: pd.DataFrame,
    country_display: str,
    zone: Optional[str],
    output_path: Path,
    logger: logging.Logger,
) -> List[dict]:
    """
    Scan a NET load dataframe and collect detailed records for any rows (days) that have:
      - ZERO values in any hourly columns 1..24
      - NEGATIVE values in any hourly columns 1..24

    Produces one issue row per day per issue type (ZERO or NEGATIVE).
    """
    net_df = norm_date_cols(net_df)
    require_columns(net_df, DATE_COLS + HOURS, f"Net load validation for {country_display}")

    hourly_numeric = net_df[HOURS].apply(pd.to_numeric, errors="coerce")

    issues: List[dict] = []

    # ---- ZERO detection ----
    zero_any = (hourly_numeric == 0.0).any(axis=1)
    if bool(zero_any.any()):
        dates = net_df.loc[zero_any, DATE_COLS].copy()
        mask_df = (hourly_numeric == 0.0).loc[zero_any, HOURS]
        vals_df = hourly_numeric.loc[zero_any, HOURS]

        for i in range(len(dates)):
            row_date = dates.iloc[i]
            mask = mask_df.iloc[i].to_numpy(dtype=bool)
            hours_list = [HOURS[j] for j, v in enumerate(mask) if v]
            cnt = int(len(hours_list))

            issues.append(
                {
                    "IssueType": "ZERO",
                    "Country": country_display,
                    "Zone": zone or "",
                    "OutputFile": output_path.name,
                    "Year": int(row_date["YEAR"]) if not pd.isna(row_date["YEAR"]) else pd.NA,
                    "Month": int(row_date["MONTH"]) if not pd.isna(row_date["MONTH"]) else pd.NA,
                    "Day": int(row_date["DAY"]) if not pd.isna(row_date["DAY"]) else pd.NA,
                    "Hours": ",".join(hours_list),
                    "Count": cnt,
                    "All24": (cnt == 24),
                    "MinValue": float(vals_df.iloc[i].min(skipna=True)),
                    "MaxValue": float(vals_df.iloc[i].max(skipna=True)),
                }
            )

        logger.warning(
            f"Net validation: '{country_display} {zone or ''}' has {len(dates)} day(s) with ZERO values."
        )

    # ---- NEGATIVE detection ----
    neg_any = (hourly_numeric < 0.0).any(axis=1)
    if bool(neg_any.any()):
        dates = net_df.loc[neg_any, DATE_COLS].copy()
        mask_df = (hourly_numeric < 0.0).loc[neg_any, HOURS]
        vals_df = hourly_numeric.loc[neg_any, HOURS]

        for i in range(len(dates)):
            row_date = dates.iloc[i]
            mask = mask_df.iloc[i].to_numpy(dtype=bool)
            hours_list = [HOURS[j] for j, v in enumerate(mask) if v]
            cnt = int(len(hours_list))

            issues.append(
                {
                    "IssueType": "NEGATIVE",
                    "Country": country_display,
                    "Zone": zone or "",
                    "OutputFile": output_path.name,
                    "Year": int(row_date["YEAR"]) if not pd.isna(row_date["YEAR"]) else pd.NA,
                    "Month": int(row_date["MONTH"]) if not pd.isna(row_date["MONTH"]) else pd.NA,
                    "Day": int(row_date["DAY"]) if not pd.isna(row_date["DAY"]) else pd.NA,
                    "Hours": ",".join(hours_list),
                    "Count": cnt,
                    "All24": (cnt == 24),
                    "MinValue": float(vals_df.iloc[i].min(skipna=True)),
                    "MaxValue": float(vals_df.iloc[i].max(skipna=True)),
                }
            )

        logger.warning(
            f"Net validation: '{country_display} {zone or ''}' has {len(dates)} day(s) with NEGATIVE values."
        )

    return issues


def write_net_validation_report(
    issues: List[dict],
    output_dir: Path,
    logger: logging.Logger,
) -> None:
    """Write net_validation_detailed.csv if any issues exist."""
    if not issues:
        logger.info("Validation: no zero or negative values found in net load outputs (hours 1..24).")
        return

    out_path = output_dir / "net_validation_detailed.csv"
    pd.DataFrame(issues).to_csv(out_path, index=False)
    logger.warning(f"[VALIDATION] Wrote net load validation report: {out_path}")


# ----------------------------------------------------------------------
# Main generator class
# ----------------------------------------------------------------------
class YearlyNetoffGeneratorModelB_FixB:
    def __init__(
        self,
        cf_dir: str,
        load_dir: str,
        supply_dir: str,
        weights_file: str = "embedded_zones_weights.json",
        scaled_profiles_dir: Optional[str] = None,
        country_profile_dir: str = r"Z:\EGP\EP\Model Runs\Automated Process\CSV Files\hourly profile files",
        use_country_dir: bool = True,
    ):
        self.cf_dir = Path(cf_dir)
        self.load_dir = Path(load_dir)
        self.supply_dir = Path(supply_dir)
        self.scaled_profiles_dir = Path(scaled_profiles_dir) if scaled_profiles_dir else None

        self.country_profile_dir = Path(country_profile_dir)
        self.use_country_dir = bool(use_country_dir)

        self.italy_zone_map = {
            "CSUD": "csu",
            "NORD": "nor",
            "CALA": "cal",
            "SUD": "sud",
            "CNOR": "cno",
            "SICI": "sic",
            "SARD": "sar",
        }

        self.country_display_names = {
            "ALBANIA": "Albania",
            "AUSTRIA": "Austria",
            "BELGIUM": "Belgium",
            "BOSNIA AND HERZEGOVINA": "Bosnia Herzegovina",
            "BOSNIA HERZEGOVINA": "Bosnia Herzegovina",
            "BULGARIA": "Bulgaria",
            "CROATIA": "Croatia",
            "CYPRUS": "Cyprus",
            "CZECH REPUBLIC": "Czechia",
            "CZECHIA": "Czechia",
            "DENMARK": "Denmark",
            "ESTONIA": "Estonia",
            "FINLAND": "Finland",
            "FRANCE": "France",
            "GERMANY": "Germany",
            "GREECE": "Greece",
            "HUNGARY": "Hungary",
            "ICELAND": "Iceland",
            "IRELAND": "Ireland",
            "ITALY": "Italy",
            "LATVIA": "Latvia",
            "LITHUANIA": "Lithuania",
            "LUXEMBOURG": "Luxembourg",
            "NORTH MACEDONIA": "Macedonia",
            "MACEDONIA": "Macedonia",
            "MALTA": "Malta",
            "NETHERLANDS": "Netherlands",
            "NORWAY": "Norway",
            "POLAND": "Poland",
            "PORTUGAL": "Portugal",
            "ROMANIA": "Romania",
            "SERBIA AND MONTENEGRO": "Serbia and Montenegro",
            "SLOVAKIA": "Slovakia",
            "SLOVENIA": "Slovenia",
            "SPAIN": "Spain",
            "SWEDEN": "Sweden",
            "SWITZERLAND": "Switzerland",
            "TURKEY": "Turkey",
            "UNITED KINGDOM": "United Kingdom",
        }

        self.setup_logging()

        weights_path = Path(__file__).parent / weights_file
        if not weights_path.exists():
            raise FileNotFoundError(f"Weights file not found: {weights_path}")
        self.weights_config = json.loads(weights_path.read_text(encoding="utf-8"))

        self.multi_zone_countries = {
            k: v for k, v in self.weights_config.items() if not k.startswith("_")
        }
        self.single_zone_countries = list(self.weights_config.get("_single_zone_countries", []))
        self.country_codes = dict(self.weights_config.get("_country_codes", {}))

        # Force Spain code override as requested
        self.country_codes["SPAIN"] = "SP"

        # Net validation issues (zeros + negatives) - NET FILES ONLY
        self.net_validation_issues: List[dict] = []

    # ----------------------------------------------------------------------
    # Logging
    # ----------------------------------------------------------------------
    def setup_logging(self) -> None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = (
            Path(__file__).parent / "output" / f"yearly_netoff_modelB_fixB_{timestamp}.log"
        )
        log_file.parent.mkdir(exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler()],
        )
        self.logger = logging.getLogger(__name__)

    # ----------------------------------------------------------------------
    # CF key mapping
    # ----------------------------------------------------------------------
    def _country_code(self, country: str) -> str:
        c = str(country).upper()
        return self.country_codes.get(c, c[:2])

    def cf_key_for_zone(self, country: str, country_code: str, file_zone: Optional[str]) -> str:
        if country.upper() == "ITALY" and file_zone:
            return f"{country_code}{self.italy_zone_map.get(file_zone, file_zone.lower())}"
        if file_zone:
            return file_zone
        return country_code

    def generate_cf_key_mapping(self) -> Dict[str, Any]:
        mapping: Dict[str, Any] = {}

        for country, config in self.multi_zone_countries.items():
            country_u = str(country).upper()
            code = str(config.get("_file_prefix", self._country_code(country_u)))

            zones = [k for k in config.keys() if not str(k).startswith("_")]
            zone_display_names = dict(config.get("_zone_display_names", {}))

            zone_map: Dict[str, str] = {}
            for zone_key in zones:
                if zone_key in zone_display_names:
                    file_zone = zone_display_names[zone_key]
                elif str(zone_key).isdigit():
                    file_zone = f"{code}{zone_key}"
                else:
                    file_zone = str(zone_key)

                zone_map[str(zone_key)] = self.cf_key_for_zone(country_u, code, file_zone)

            mapping[country_u] = {"country_code": code, "zones": zone_map}

        for country in self.single_zone_countries:
            country_u = str(country).upper()
            code = self._country_code(country_u)
            mapping[country_u] = {
                "country_code": code,
                "zones": {"": self.cf_key_for_zone(country_u, code, None)},
            }

        return mapping

    # ---------------- Diego profile lookup ----------------
    def ensure_country_dir(self, country_display: str) -> Path:
        self.country_profile_dir.mkdir(parents=True, exist_ok=True)
        if not self.use_country_dir:
            return self.country_profile_dir
        d = self.country_profile_dir / safe_country_dirname(country_display)
        d.mkdir(parents=True, exist_ok=True)
        return d

    def find_country_profile(self, country_display: str, zone: Optional[str]) -> Path:
        base = self.country_profile_dir
        country_dir = self.ensure_country_dir(country_display)

        if zone:
            candidates = [
                country_dir / f"{country_display} {zone}Load2016.csv",
                country_dir / f"{country_display} {zone}Load.csv",
                base / f"{country_display} {zone}Load2016.csv",
                base / f"{country_display} {zone}Load.csv",
            ]
            for p in candidates:
                if p.exists():
                    self.logger.info(f"  Diego shape (zonal): {p}")
                    return p
            raise FileNotFoundError(
                f"Missing REQUIRED zonal Diego profile for '{country_display}' zone='{zone}'. "
                f"Tried:\n  - " + "\n  - ".join(str(x) for x in candidates)
            )

        candidates = [
            country_dir / f"{country_display}Load2016.csv",
            country_dir / f"{country_display}Load.csv",
            base / f"{country_display}Load2016.csv",
            base / f"{country_display}Load.csv",
        ]
        for p in candidates:
            if p.exists():
                self.logger.info(f"  Diego shape (single-zone): {p}")
                return p
        raise FileNotFoundError(
            f"Missing Diego profile for single-zone '{country_display}'. "
            f"Tried:\n  - " + "\n  - ".join(str(x) for x in candidates)
        )

    # ---------------- Generic file finders ----------------
    def find_supply_file(self, country: str) -> Optional[Path]:
        name_variations = {
            "BOSNIA AND HERZEGOVINA": ["BOSNIA HERZEGOVINA", "BOSNIA AND HERZEGOVINA"],
            "CZECH REPUBLIC": ["CZECHIA", "CZECH REPUBLIC"],
            "NORTH MACEDONIA": ["MACEDONIA", "NORTH MACEDONIA"],
            "SERBIA AND MONTENEGRO": ["SERBIA AND MONTENEGRO", "SERBIA"],
        }
        
        # 1. Get base variation forms (mapped aliases or case variations)
        bases = name_variations.get(country.upper(), [])
        bases += [country.upper(), country, country.title()]
        
        # 2. De-duplicate bases while preserving sequence order
        seen_bases = set()
        unique_bases = [b for b in bases if not (b in seen_bases or seen_bases.add(b))]
        
        # 3. Dynamically apply version variations ('- v2', '- V2', or raw) to ALL resolved base names
        names_to_try = []
        for b in unique_bases:
            names_to_try.append(f"{b} - v2")
            names_to_try.append(f"{b} - V2")
            names_to_try.append(f"{b} - v4")
            names_to_try.append(f"{b} v4")
            names_to_try.append(b)
            
        # 4. De-duplicate final list
        seen_final = set()
        final_search_list = [n for n in names_to_try if not (n in seen_final or seen_final.add(n))]
        
        # 5. Search for files
        for name in final_search_list:
            for ext in [".xlsx", ".xls"]:
                p = self.supply_dir / f"{name}{ext}"
                if p.exists():
                    self.logger.info(f"  Supply file found: {p.name}")
                    return p
        return None

    def find_load_file_gross(self, country_display: str, zone: Optional[str]) -> Optional[Path]:
        candidates: List[Path] = []
        if zone:
            candidates += [
                self.load_dir / f"{country_display} {zone}Load2016.csv",
                self.load_dir / f"{country_display} {zone}Load.csv",
            ]
        candidates += [
            self.load_dir / f"{country_display}Load2016.csv",
            self.load_dir / f"{country_display}Load.csv",
        ]
        for p in candidates:
            if p.exists():
                self.logger.info(f"  Gross demand file: {p.name}")
                return p
        return None

    def _search_cf_dirs(self) -> List[Path]:
        dirs: List[Path] = []
        if self.scaled_profiles_dir and self.scaled_profiles_dir.exists():
            dirs.append(self.scaled_profiles_dir)
        dirs.append(self.cf_dir)
        return dirs

    # ---------------- CF file finders ----------------
    def find_wind_cf_files(self, key: str) -> Tuple[Optional[Path], Optional[Path]]:
        onshore = None
        offshore = None

        for d in self._search_cf_dirs():
            p = d / f"{key}WindON.csv"
            if onshore is None and p.exists():
                onshore = p
            p = d / f"{key}WindOFF.csv"
            if offshore is None and p.exists():
                offshore = p

        if onshore is None:
            old = self.cf_dir / f"{key}WindCF2016.csv"
            if old.exists():
                onshore = old

        if onshore is None:
            for d in self._search_cf_dirs():
                fb = d / "WindON.csv"
                if fb.exists():
                    onshore = fb
                    self.logger.info(f"  Using fallback WindON.csv for key={key}")
                    break

        if offshore is None:
            for d in self._search_cf_dirs():
                fb = d / "WindOFF.csv"
                if fb.exists():
                    offshore = fb
                    self.logger.info(f"  Using fallback WindOFF.csv for key={key}")
                    break

        return onshore, offshore

    def find_solar_cf_file(self, key: str) -> Optional[Path]:
        for d in self._search_cf_dirs():
            p = d / f"{key}Solar.csv"
            if p.exists():
                return p

        old = self.cf_dir / f"{key}SolarCF2016.csv"
        if old.exists():
            return old

        for d in self._search_cf_dirs():
            fb = d / "Solar.csv"
            if fb.exists():
                self.logger.info(f"  Using fallback Solar.csv for key={key}")
                return fb
        return None

    # ---------------- Supply readers ----------------
    def read_supply_file_capacity(self, country: str) -> Dict[int, Dict[str, float]]:
        """
        Read embedded installed capacity from fixed row ranges (Excel row numbers).

        Non-Italy:
          - SolarPV: rows 440-447
          - Wind: onshore rows 440-443, offshore rows 444-447 (wind = onshore + offshore)

        Italy exception:
          - SolarPV: rows 486-493
          - Wind: onshore rows 486-489, offshore rows 490-493 (wind = onshore + offshore)

        Notes:
          - Values are assumed in GW and converted to MW by *1000.
          - Year header row/col convention preserved from previous implementation.
        """
        supply_file = self.find_supply_file(country)
        if not supply_file:
            return {}

        wind_df = pd.read_excel(supply_file, sheet_name="Wind", header=None)
        solar_df = pd.read_excel(supply_file, sheet_name="SolarPV", header=None)

        year_row = 5
        year_start_col = 5

        country_u = str(country).upper()

        # UNIFIED ROWS FOR ALL COUNTRIES (use Italy rows everywhere)
        # Excel rows:
        #   SolarPV: 486-493
        #   Wind onshore: 486-489
        #   Wind offshore: 490-493
        # pandas iloc is 0-based -> subtract 1
        solar_rows   = range(485, 493)  # 486-493 inclusive
        wind_on_rows = range(485, 489)  # 486-489 inclusive
        wind_off_rows= range(489, 493)  # 490-493 inclusive

        def sum_rows_as_mw(df: pd.DataFrame, rows: range, col: int) -> float:
            total_gw = 0.0
            for r in rows:
                if r < 0 or r >= len(df):
                    continue
                v = df.iloc[r, col]
                if pd.notna(v) and isinstance(v, (int, float)):
                    total_gw += float(v)
            return total_gw * 1000.0  # GW -> MW

        cap: Dict[int, Dict[str, float]] = {}

        for c in range(year_start_col, min(200, len(wind_df.columns))):
            yr = wind_df.iloc[year_row, c]
            if pd.notna(yr) and isinstance(yr, (int, float)):
                yr = int(yr)

                wind_on_mw = sum_rows_as_mw(wind_df, wind_on_rows, c)
                wind_off_mw = sum_rows_as_mw(wind_df, wind_off_rows, c)
                wind_mw = wind_on_mw + wind_off_mw

                solar_mw = sum_rows_as_mw(solar_df, solar_rows, c)

                cap[yr] = {"wind": wind_mw, "solar": solar_mw}

        self.logger.info(f"  Capacity years read: {len(cap)}")
        return cap

    def read_embedded_target_twh(self, country: str) -> Dict[int, float]:
        supply_file = self.find_supply_file(country)
        if not supply_file:
            return {}
            
        df = pd.read_excel(supply_file, sheet_name="Demand", header=None)
        target_row = None
        for i in range(len(df)):
            a = df.iloc[i, 0]
            b = df.iloc[i, 1] if len(df.columns) > 1 else None
            if pd.notna(a) and str(a).strip() == "Embedded":
                if pd.notna(b) and "%" in str(b):
                    target_row = i
                    break
                    
        if target_row is None:
            self.logger.warning(f"{country}: 'Embedded %' row not found.")
            return {}
            
        year_row = 1
        year_start_col = 2
        embedded_pct_by_year: Dict[int, float] = {}
        
        for c in range(year_start_col, min(200, len(df.columns))):
            yr = df.iloc[year_row, c]
            try:
                if pd.notna(yr):
                    yr = int(float(yr))
                else:
                    continue
            except (ValueError, TypeError):
                continue
                
            v = df.iloc[target_row, c]
            # Safely catch and assign 0.0 value if column cell is empty or NaN
            if pd.isna(v) or str(v).strip() == "":
                val = 0.0
            else:
                try:
                    val = float(v)
                except (ValueError, TypeError):
                    val = 0.0
            embedded_pct_by_year[yr] = val
                    
        if not embedded_pct_by_year:
            return {}

        country_display = self.country_display_names.get(country.upper(), country.title())

        # Compute COUNTRY gross by year (sum zones if multi-zone)
        if country.upper() in self.multi_zone_countries:
            config = self.multi_zone_countries[country.upper()]
            zones = [k for k in config.keys() if not str(k).startswith("_")]
            country_code = config.get("_file_prefix", self._country_code(country))
            zone_display_names = config.get("_zone_display_names", {})

            gross_twh_total: Dict[int, float] = {}
            for zone_key in zones:
                file_zone = self.resolve_zone_for_files(country_code, zone_key, zone_display_names)
                gross_path = self.find_load_file_gross(country_display, file_zone)
                if not gross_path:
                    self.logger.error(
                        f"{country}: missing gross demand file for zone {file_zone} "
                        f"(needed for embedded % -> TWh)"
                    )
                    continue

                gdf = pd.read_csv(gross_path)
                gdf = norm_date_cols(gdf)
                require_columns(gdf, DATE_COLS + HOURS, f"Gross file {gross_path.name}")
                zone_gross = annual_twh_from_hourly_mw(gdf)

                for y, twh in zone_gross.items():
                    gross_twh_total[y] = gross_twh_total.get(y, 0.0) + float(twh)
        else:
            gross_path = self.find_load_file_gross(country_display, None)
            if not gross_path:
                self.logger.error(
                    f"{country}: missing country gross demand file (needed for embedded % -> TWh)"
                )
                return {}

            gdf = pd.read_csv(gross_path)
            gdf = norm_date_cols(gdf)
            require_columns(gdf, DATE_COLS + HOURS, f"Gross file {gross_path.name}")
            gross_twh_total = annual_twh_from_hourly_mw(gdf)

        out: Dict[int, float] = {}
        for y, pct in embedded_pct_by_year.items():
            g = gross_twh_total.get(y)
            if g is None:
                continue
            out[y] = float(g) * float(pct)

        self.logger.info(
            f"  Target years read (Embedded % → TWh using gross): {len(out)} (year_row=1)"
        )
        return out

    def read_market_zone_weights_for_year(self, country: str, year: int) -> Optional[Dict[str, float]]:
        if country.upper() not in self.multi_zone_countries:
            return None
        supply_file = self.find_supply_file(country)
        if not supply_file:
            return None
            
        demand_df = pd.read_excel(supply_file, sheet_name="Demand", header=None)
        config = self.multi_zone_countries[country.upper()]
        zones = [k for k in config.keys() if not k.startswith("_")]
        
        zone_start_row = 40
        start_col = 8
        end_col = 43
        year_col = None
        
        for r in range(0, min(60, len(demand_df))):
            for c in range(start_col, min(end_col, len(demand_df.columns))):
                v = demand_df.iloc[r, c]
                if pd.notna(v) and isinstance(v, (int, float)) and int(v) == int(year):
                    year_col = c
                    break
            if year_col is not None:
                break
                
        if year_col is None:
            raise ValueError(f"{country}: zone weight column missing for Year {year}")
            
        raw: Dict[str, float] = {}
        for i, z in enumerate(zones):
            row_idx = zone_start_row + i
            if row_idx >= len(demand_df):
                raw[z] = 0.0
                continue
                
            v = demand_df.iloc[row_idx, year_col]
            
            # If the cell is NaN or contains an empty/whitespace string, treat it as 0.0
            if pd.isna(v) or str(v).strip() == "":
                raw[z] = 0.0
            else:
                try:
                    raw[z] = float(v)
                except (ValueError, TypeError):
                    raw[z] = 0.0  # Fallback for any invalid text cells
            
        total = sum(raw.values())
        if total <= 0:
            n = len(zones)
            return {z: 1.0 / n for z in zones}
        return {z: raw[z] / total for z in zones}

    # ---------------- CF processing ----------------
    def read_combined_wind_cf(
        self, onshore: Optional[Path], offshore: Optional[Path]
    ) -> Optional[pd.DataFrame]:
        def read(p: Path) -> pd.DataFrame:
            df = pd.read_csv(p)
            df = norm_date_cols(df)
            require_columns(df, DATE_COLS + HOURS, f"CF file {p.name}")
            return df

        if onshore and offshore:
            on = read(onshore)
            off = read(offshore)
            comb = on.copy()
            for h in HOURS:
                comb[h] = (on[h] + off[h]) / 2.0
            return comb
        if onshore:
            return read(onshore)
        if offshore:
            return read(offshore)
        return None

    def annual_twh_from_cf_and_capacity(
        self, cf_df: pd.DataFrame, cap_mw_by_year: Dict[int, float]
    ) -> Dict[int, float]:
        out: Dict[int, float] = {}
        cf_df = norm_date_cols(cf_df)
        require_columns(cf_df, DATE_COLS + HOURS, "CF dataframe")
        for y, g in cf_df.groupby("YEAR"):
            y = int(y)
            cap = cap_mw_by_year.get(y)
            if cap is None:
                continue
            mwh = (g[HOURS].to_numpy(dtype="float64") / 100.0 * cap).sum()
            out[y] = mwh / MWH_PER_TWH
        return out

    def resolve_zone_for_files(
        self, country_code: str, zone_key: str, zone_display_names: dict
    ) -> str:
        if zone_key in zone_display_names:
            return zone_display_names[zone_key]
        if str(zone_key).isdigit():
            return f"{country_code}{zone_key}"
        return str(zone_key)

    # ---------------- Core processing ----------------
    def process_zone(
        self,
        country: str,
        country_display: str,
        country_code: str,
        zone_key: Optional[str],
        zone: Optional[str],
        cap_by_year: Dict[int, Dict[str, float]],
        target_twh_by_year: Dict[int, float],
        output_dir: Path,
        summary_rows: List[dict],
    ) -> Optional[pd.DataFrame]:
        zone_label = f"{country_display} {zone}".strip()
        self.logger.info(f"Processing zone (FixB): {zone_label}")

        gross_path = self.find_load_file_gross(country_display, zone)
        if not gross_path:
            self.logger.error(f"  Missing gross demand file for {zone_label}")
            return None

        gross_df = pd.read_csv(gross_path)
        gross_df = norm_date_cols(gross_df)
        require_columns(gross_df, DATE_COLS + HOURS, f"Gross file {gross_path.name}")

        gross_twh = annual_twh_from_hourly_mw(gross_df)
        years = sorted(set(gross_twh.keys()) & set(target_twh_by_year.keys()))
        if not years:
            self.logger.warning(f"  No overlapping years for {zone_label}")
            return None

        cf_key = self.cf_key_for_zone(country, country_code, zone if zone else None)

        wind_on, wind_off = self.find_wind_cf_files(cf_key)
        solar_cf_path = self.find_solar_cf_file(cf_key)
        wind_cf = self.read_combined_wind_cf(wind_on, wind_off)
        if wind_cf is None or solar_cf_path is None:
            self.logger.error(
                f"  Missing CF files (even after fallback) for key={cf_key} ({zone_label})"
            )
            return None

        solar_cf = pd.read_csv(solar_cf_path)
        solar_cf = norm_date_cols(solar_cf)
        require_columns(solar_cf, DATE_COLS + HOURS, f"Solar CF file {solar_cf_path.name}")

        profile_path = self.find_country_profile(country_display, zone)
        diego_df = pd.read_csv(profile_path)
        diego_df = norm_date_cols(diego_df)
        require_columns(diego_df, DATE_COLS + HOURS, f"Diego profile file {profile_path.name}")

        net_parts: List[pd.DataFrame] = []
        emb_parts: List[pd.DataFrame] = []

        for y in years:
            # zone weight
            if zone_key is None:
                zone_weight_y = 1.0
            else:
                zone_weights_y = self.read_market_zone_weights_for_year(country, y)
                zone_weight_y = zone_weights_y.get(zone_key) if zone_weights_y else None
                if zone_weight_y is None:
                    raise KeyError(
                        f"Missing zone weight for {country} zone_key={zone_key} year={y}"
                    )

            # annual gross and zonal embedded target
            gross_y_twh = gross_twh[y]
            target_zone = target_twh_by_year[y] * zone_weight_y

            # allocate embedded capacities to zone
            wind_cap_zone = {yy: cap_by_year[yy]["wind"] * zone_weight_y for yy in cap_by_year}
            solar_cap_zone = {yy: cap_by_year[yy]["solar"] * zone_weight_y for yy in cap_by_year}

            # compute wind and solar generation (TWh)
            wind_twh_y = self.annual_twh_from_cf_and_capacity(wind_cf, wind_cap_zone).get(y, 0.0)
            solar_twh_y = self.annual_twh_from_cf_and_capacity(solar_cf, solar_cap_zone).get(
                y, 0.0
            )
            ws = wind_twh_y + solar_twh_y

            # embedded-target rules
            if target_zone == 0:
                target_eff = ws
                scale = 1.0
            elif target_zone < ws and ws > 0:
                scale = target_zone / ws
                wind_twh_y *= scale
                solar_twh_y *= scale
                target_eff = target_zone
            elif target_zone > ws:
                target_eff = target_zone
                scale = 1.0
            else:
                target_eff = target_zone
                scale = 1.0

            # Net-load target (annual TWh)
            net_twh_target = gross_y_twh - target_eff

            # Shape net load using annual-normalized Diego (aligned to gross days)
            weights_y = diego_annual_normalized_weights(diego_df, year=y)

            gross_y_df = gross_df[gross_df["YEAR"] == y][DATE_COLS + HOURS].copy()

            # Align gross and Diego day-rows (leap/missing-day safety)
            gross_y_df_aligned, weights_y_aligned = align_year_on_dates(
                gross_y_df,
                weights_y,
                logger=self.logger,
                label=f"{zone_label} Year {y} (gross vs diego):",
            )

            # Convert the net target (TWh) into hourly MW using aligned weights
            net_h_raw_aligned = shape_twh_to_hourly_mw_annual(weights_y_aligned, net_twh_target)

            # Fix B: cap NetMW ≤ GrossMW and redistribute to hit annual target if possible
            net_h = cap_and_redistribute_net_to_gross(
                gross_df_y=gross_y_df_aligned,
                net_df_y=net_h_raw_aligned,
                target_net_twh=net_twh_target,
                max_iters=30,
                tol_mwh=0.001,
            )

            # Annual checks (on aligned data)
            gross_check = gross_y_df_aligned[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH
            net_check = net_h[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH

            # Zero rule: if embedded rounds to 0.00 TWh => Embedded=0 and Net=Gross
            zero_applied = False
            if zero_rule_applies(gross_check, net_check):
                zero_applied = True
                self.logger.info(f"  Year {y} zero-rule triggered: forcing Embedded=0 and Net=Gross.")
                net_h = gross_y_df_aligned[DATE_COLS + HOURS].copy()
                net_check = gross_check

            # Embedded residual = Gross - Net
            merged = gross_y_df_aligned.merge(
                net_h, on=DATE_COLS, how="inner", suffixes=("_gross", "_net")
            )
            emb_h = merged[DATE_COLS].copy()
            for h in HOURS:
                emb_h[h] = merged[f"{h}_gross"] - merged[f"{h}_net"]

            emb_check = emb_h[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH

            # Diagnostics
            raw_net = net_h_raw_aligned[HOURS].to_numpy(dtype="float64")
            gross_np = gross_y_df_aligned[HOURS].to_numpy(dtype="float64")
            capped_hours_raw = int((raw_net > gross_np).sum())

            self.logger.info(
                f"  Year {y} check (TWh): Gross={gross_check:.3f} Net={net_check:.3f} "
                f"Embedded={emb_check:.3f} CappedHoursRaw={capped_hours_raw} ZeroRule={zero_applied}"
            )

            summary_rows.append(
                {
                    "Country": country_display,
                    "Zone": zone or "",
                    "Year": y,
                    "Gross Demand (TWh)": gross_y_twh,
                    "Target Embedded effective (TWh)": target_eff,
                    "Net Load target (TWh)": net_twh_target,
                    "Net Load achieved after FixB (TWh)": net_check,
                    "Embedded (residual) (TWh)": emb_check,
                    "Zone weight (year-specific)": zone_weight_y,
                    "Wind (TWh) zonal (post-scale)": wind_twh_y,
                    "Solar (TWh) zonal (post-scale)": solar_twh_y,
                    "Wind+Solar scaling (if target<WS)": scale,
                    "Wind Capacity (MW)": float(wind_cap_zone.get(y, 0.0)),
                    "Solar Capacity (MW)": float(solar_cap_zone.get(y, 0.0)),
                    "CF key": cf_key,
                    "Country code": country_code,
                    "Diego shape file": str(profile_path),
                    "Gross file": str(gross_path),
                    "Capped hours (raw)": capped_hours_raw,
                    "Zero rule applied": zero_applied,
                    "Aligned days used (gross∩diego)": int(len(gross_y_df_aligned)),
                }
            )

            net_parts.append(net_h[DATE_COLS + HOURS])
            emb_parts.append(emb_h[DATE_COLS + HOURS])

        # ----- concat all years for this zone -----
        net_hourly = pd.concat(net_parts, ignore_index=True) if net_parts else None
        emb_hourly = pd.concat(emb_parts, ignore_index=True) if emb_parts else None

        # ----- write net file for this zone (or country for single-zone) -----
        if net_hourly is not None:
            out_net = output_dir / (
                f"{country_display} {zone}Load.csv" if zone else f"{country_display}Load.csv"
            )
            net_hourly.to_csv(out_net, index=False)
            self.logger.info(f"  [OK] Saved net load: {out_net.name}")

            # VALIDATION: net load only (NOT embedded) - check for ZEROS and NEGATIVES
            try:
                self.net_validation_issues.extend(
                    validate_net_load_for_zeros_and_negatives(
                        net_hourly,
                        country_display=country_display,
                        zone=zone,
                        output_path=out_net,
                        logger=self.logger,
                    )
                )
            except Exception as e:
                self.logger.error(f"  Net validation failed for {out_net.name}: {e}")

        return emb_hourly

    def aggregate_embedded(self, embedded_dfs: List[pd.DataFrame]) -> Optional[pd.DataFrame]:
        if not embedded_dfs:
            return None
        if len(embedded_dfs) == 1:
            return embedded_dfs[0][DATE_COLS + HOURS].copy()

        def prep(df: pd.DataFrame) -> pd.DataFrame:
            df = df.copy()
            df["__key__"] = list(zip(df["YEAR"], df["MONTH"], df["DAY"]))
            return df.set_index("__key__")

        agg = prep(embedded_dfs[0])
        for df in embedded_dfs[1:]:
            other = prep(df)
            for h in HOURS:
                agg[h] = agg[h].add(other[h], fill_value=0.0)

        out = agg.reset_index()
        out[["YEAR", "MONTH", "DAY"]] = pd.DataFrame(out["__key__"].tolist(), index=out.index)
        out = out.drop(columns=["__key__"])
        return out[DATE_COLS + HOURS]

    def redistribute_embedded_load_to_diego_profile(
        self,
        *,
        country_display: str,
        zone: Optional[str],
        output_dir: Path,
    ) -> Path:
        """
        Read "{CountryDisplay} EmbeddedLoad.csv", compute annual embedded TWh, and
        redistribute hourly MW to match Diego annual-normalized profile (zonal if zone provided),
        while preserving annual embedded energy.
        Writes "{CountryDisplay} EmbeddedLoad.csv".
        """
        emb_path = output_dir / f"{country_display} EmbeddedLoad.csv"
        if not emb_path.exists():
            raise FileNotFoundError(f"Embedded load CSV not found: {emb_path}")

        emb_df = pd.read_csv(emb_path)
        emb_df = norm_date_cols(emb_df)
        require_columns(emb_df, DATE_COLS + HOURS, f"Embedded file {emb_path.name}")

        profile_path = self.find_country_profile(country_display, zone)
        diego_df = pd.read_csv(profile_path)
        diego_df = norm_date_cols(diego_df)
        require_columns(diego_df, DATE_COLS + HOURS, f"Diego profile file {profile_path.name}")

        years = sorted(int(y) for y in emb_df["YEAR"].dropna().unique().tolist())
        out_parts: List[pd.DataFrame] = []

        for y in years:
            emb_y = emb_df[emb_df["YEAR"] == y][DATE_COLS + HOURS].copy()
            if emb_y.empty:
                continue

            emb_twh_y = emb_y[HOURS].to_numpy(dtype="float64").sum() / MWH_PER_TWH
            weights_y = diego_annual_normalized_weights(diego_df, year=y)

            emb_y_aligned, weights_y_aligned = align_year_on_dates(
                emb_y,
                weights_y,
                logger=self.logger,
                label=f"{country_display} embedded redistribution Year {y} (emb vs diego):",
            )

            if len(weights_y_aligned) == 0:
                self.logger.warning(
                    f"  Embedded redistribution {country_display} Year {y}: "
                    f"no overlapping days (skipping year)"
                )
                continue

            redist_y = shape_twh_to_hourly_mw_annual(weights_y_aligned, emb_twh_y)
            out_parts.append(redist_y[DATE_COLS + HOURS])

            self.logger.info(
                f"  Embedded redistribution {country_display} Year {y}: "
                f"OriginalEmbedded={emb_twh_y:.3f} TWh using Diego='{profile_path.name}' "
                f"(aligned_days={len(weights_y_aligned)})"
            )

        if not out_parts:
            raise ValueError(f"No yearly data found to redistribute for {country_display}")

        out_df = pd.concat(out_parts, ignore_index=True)
        out_path = output_dir / f"{country_display} EmbeddedLoad.csv"
        out_df.to_csv(out_path, index=False)
        self.logger.info(f"  [OK] Saved embedded load redistributed: {out_path.name}")
        return out_path

    def generate(self, output_dir="output", country_filter=None):
        out_dir = Path(output_dir)
        out_dir.mkdir(exist_ok=True)

        # ---- CF-key mapping diagnostics ----
        try:
            cf_key_map = self.generate_cf_key_mapping()
            (out_dir / "cf_key_mapping.json").write_text(
                json.dumps(cf_key_map, indent=2), encoding="utf-8"
            )
            self.logger.info(f"[OK] Saved cf_key_mapping.json to: {out_dir}")
        except Exception as e:
            self.logger.warning(f"Could not write cf_key_mapping.json ({e})")

        # Default to ALL if not provided (or explicitly ALL)
        if (
            not country_filter
            or str(country_filter).strip() == ""
            or str(country_filter).upper() == "ALL"
        ):
            countries = list(self.multi_zone_countries.keys()) + list(self.single_zone_countries)
        else:
            countries = [str(country_filter).upper()]

        summary_rows: List[dict] = []
        countries_with_embedded_written: List[str] = []

        for country in countries:
            country_display = self.country_display_names.get(country.upper(), country.title())
            self.logger.info("=" * 80)
            self.logger.info(f"PROCESSING: {country} ({country_display})")
            self.logger.info("=" * 80)

            cap_by_year = self.read_supply_file_capacity(country)
            target_twh_by_year = self.read_embedded_target_twh(country)
            if not cap_by_year or not target_twh_by_year:
                self.logger.error(f"Skipping {country}: missing capacities or embedded targets")
                continue

            if country.upper() in self.multi_zone_countries:
                embedded_zone_dfs: List[pd.DataFrame] = []

                config = self.multi_zone_countries[country.upper()]
                zones = [k for k in config.keys() if not k.startswith("_")]
                country_code = config.get("_file_prefix", self._country_code(country))
                zone_display_names = config.get("_zone_display_names", {})

                for zone_key in zones:
                    file_zone = self.resolve_zone_for_files(country_code, zone_key, zone_display_names)
                    emb_df = self.process_zone(
                        country=country,
                        country_display=country_display,
                        country_code=country_code,
                        zone_key=zone_key,
                        zone=file_zone,
                        cap_by_year=cap_by_year,
                        target_twh_by_year=target_twh_by_year,
                        output_dir=out_dir,
                        summary_rows=summary_rows,
                    )
                    if emb_df is not None:
                        embedded_zone_dfs.append(emb_df)

                agg_emb = self.aggregate_embedded(embedded_zone_dfs)
                if agg_emb is not None:
                    out_emb = out_dir / f"{country_display} EmbeddedLoad.csv"
                    agg_emb.to_csv(out_emb, index=False)
                    self.logger.info(f"  [OK] Saved embedded load (aggregated): {out_emb.name}")
                    countries_with_embedded_written.append(country_display)

            else:
                country_code = self._country_code(country)
                emb_df = self.process_zone(
                    country=country,
                    country_display=country_display,
                    country_code=country_code,
                    zone_key=None,
                    zone=None,
                    cap_by_year=cap_by_year,
                    target_twh_by_year=target_twh_by_year,
                    output_dir=out_dir,
                    summary_rows=summary_rows,
                )
                if emb_df is not None:
                    out_emb = out_dir / f"{country_display} EmbeddedLoad.csv"
                    emb_df.to_csv(out_emb, index=False)
                    self.logger.info(f"  [OK] Saved embedded load: {out_emb.name}")
                    countries_with_embedded_written.append(country_display)

        if summary_rows:
            pd.DataFrame(summary_rows).to_csv(out_dir / "yearly_summary.csv", index=False)
            self.logger.info("[OK] Saved yearly_summary.csv")

        # Post-processing: redistribute embedded to Diego annual-normalized profile (annual TWh preserved)
        multi_zone_embedded_shape_override = {
            "Denmark": "DK1",
            "Italy": "SARD",
            "Sweden": "SE1",
            "Norway": "NO4",
            "United Kingdom": "Great Britain",
            "Serbia and Montenegro": "Serbia",
        }

        for country_display in sorted(set(countries_with_embedded_written)):
            try:
                zone_for_profile = multi_zone_embedded_shape_override.get(country_display, None)
                self.redistribute_embedded_load_to_diego_profile(
                    country_display=country_display,
                    zone=zone_for_profile,
                    output_dir=out_dir,
                )
            except Exception as e:
                self.logger.error(f"Embedded redistribution failed for {country_display}: {e}")

        # ---- WRITE NET VALIDATION REPORT (zeros + negatives) ----
        write_net_validation_report(
            self.net_validation_issues,
            output_dir=out_dir,
            logger=self.logger,
        )


def main() -> int:
    p = argparse.ArgumentParser(
        description="Yearly netoff generator - annual-normalized Diego + CF key map + net validation (zeros + negatives)"
    )
    p.add_argument("--cf-dir", required=True, help="Directory containing CF CSVs")
    p.add_argument("--load-dir", required=True, help="Gross demand directory (365x24 hourly MW)")
    p.add_argument("--supply-dir", required=True, help="Supply Excel directory")
    p.add_argument("--output", default="output", help="Output directory")
    p.add_argument("--country", help="Country filter (e.g., SWEDEN) or ALL (default: ALL)")
    p.add_argument("--weights-file", default="embedded_zones_weights.json")
    p.add_argument("--scaled-profiles-dir", help="Optional directory containing scaled CF profiles")
    p.add_argument(
        "--country-profile-dir",
        default=r"Z:\EGP\EP\Model Runs\Automated Process\CSV Files\hourly profile files",
        help="Directory containing Diego hourly profiles",
    )
    p.add_argument("--no-country-dir", action="store_true", help="Disable per-country subfolder in Diego dir")
    args = p.parse_args()

    try:
        gen = YearlyNetoffGeneratorModelB_FixB(
            cf_dir=args.cf_dir,
            load_dir=args.load_dir,
            supply_dir=args.supply_dir,
            weights_file=args.weights_file,
            scaled_profiles_dir=args.scaled_profiles_dir,
            country_profile_dir=args.country_profile_dir,
            use_country_dir=(not args.no_country_dir),
        )
        gen.generate(output_dir=args.output, country_filter=args.country)
        return 0
    except Exception as e:
        print(f"FATAL ERROR: {e}")
        import traceback

        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
