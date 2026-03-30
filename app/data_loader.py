"""
Data loading layer.

Supports two modes:
  1. PQ mode  – reads from PQ_* sheets already populated inside the Excel template.
  2. Direct mode – reads planner files from a folder + master data files directly.

The optimizer always works from the in-memory objects; the mode only affects
where pandas DataFrames come from.
"""
import os
import glob
import re
from datetime import datetime
from typing import List, Tuple

import pandas as pd

from app.models import Config, LoadRecord, CapacityRecord, RoutingRecord


PLANNER_FILE_RE = re.compile(r"^planner([1-6])_load\.(xlsx|xls|csv)$", re.IGNORECASE)


# ────────────────────────────────────────────────────────────────────────────
# Config reader
# ────────────────────────────────────────────────────────────────────────────

def load_config(template_path: str) -> Config:
    """Read the Config sheet (key→value table) from the Excel template."""
    template_dir = os.path.dirname(os.path.abspath(template_path))
    with pd.ExcelFile(template_path) as workbook:
        sheet_name = "Control_Panel" if "Control_Panel" in workbook.sheet_names else "Config"
        df = workbook.parse(sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    # Expect columns: Parameter, Value
    mapping = dict(zip(df["Parameter"].str.strip(), df["Value"].astype(str).str.strip()))

    def _get(key, default=None):
        v = mapping.get(key, default)
        if v in (None, "nan", "None", ""):
            return default
        return v

    start_year = _get("Start_Year")
    start_month_num = _get("Start_Month_Num")
    if start_year not in (None, "") and start_month_num not in (None, ""):
        start_month = f"{int(float(start_year))}-{int(float(start_month_num)):02d}"
    else:
        start_month = _get("Start_Month", _default_start_month())

    default_project_root = os.path.abspath(os.path.join(template_dir, ".."))
    project_root_folder = _resolve_config_folder(
        _get("Project_Root_Folder"),
        base_dir=template_dir,
        default=default_project_root,
    )

    return Config(
        project_root_folder=project_root_folder,
        input_load_folder=_resolve_config_folder(
            _get("Input_Load_Folder", "Data_Input"),
            base_dir=project_root_folder,
            default=os.path.join(project_root_folder, "Data_Input"),
        ),
        input_master_folder=_resolve_config_folder(
            _get("Input_Master_Folder", "Data_Input"),
            base_dir=project_root_folder,
            default=os.path.join(project_root_folder, "Data_Input"),
        ),
        output_folder=_resolve_config_folder(
            _get("Output_Folder", "output"),
            base_dir=project_root_folder,
            default=os.path.join(project_root_folder, "output"),
        ),
        output_file_name=_get("Output_FileName", "capacity_optimizer_result.xlsx"),
        scenario_name=_get("Scenario_Name", "Base"),
        start_month=start_month,
        horizon_months=int(float(_get("Horizon_Months", 60))),
        run_timestamp=_get("Run_Timestamp"),
        notes=_get("Notes"),
        run_mode=_normalize_run_mode(_get("Run_Mode", "ModeB")),
        direct_mode=_parse_bool(_get("Direct_Mode", "Yes"), True),
        verbose=_parse_bool(_get("Verbose", "No"), False),
        skip_validation_errors=_parse_bool(_get("Skip_Validation_Errors", "No"), False),
    )


def _resolve_config_folder(value, base_dir: str, default: str) -> str:
    if value in (None, "", "nan", "None"):
        return os.path.normpath(default)
    text = str(value).strip()
    if not text:
        return os.path.normpath(default)
    if os.path.isabs(text):
        return os.path.normpath(text)
    return os.path.normpath(os.path.join(base_dir, text))


def _default_start_month() -> str:
    now = datetime.now()
    return f"{now.year}-{now.month:02d}"


def _parse_bool(value, default: bool) -> bool:
    if value in (None, ""):
        return default
    text = str(value).strip().lower()
    if text in {"y", "yes", "true", "1", "on"}:
        return True
    if text in {"n", "no", "false", "0", "off"}:
        return False
    return default


def _normalize_run_mode(value) -> str:
    text = str(value or "").strip().lower()
    if text in {"modea", "mode-a", "a"}:
        return "ModeA"
    if text in {"both", "modea+modeb", "all"}:
        return "Both"
    return "ModeB"


# ────────────────────────────────────────────────────────────────────────────
# Mode 1: read from PQ output sheets inside the template
# ────────────────────────────────────────────────────────────────────────────

def load_from_template_pq(
    template_path: str,
    include_routing: bool = True,
    selected_scenario: str | None = None,
) -> Tuple[
    List[LoadRecord], List[CapacityRecord], List[RoutingRecord]
]:
    """Read consolidated data from PQ_* sheets already in the template."""
    xl = pd.ExcelFile(template_path)
    load_df = _read_sheet(xl, "PQ_Load_Consolidated")
    cap_df  = _read_sheet(xl, "PQ_Capacity")

    loads     = _apply_scenario_filter(_parse_load_df(load_df), selected_scenario)
    caps      = _parse_capacity_df(cap_df)
    if include_routing:
        rout_df = _read_sheet(xl, "PQ_Routing")
        routings = _parse_routing_df(rout_df)
    else:
        routings = []
    return loads, caps, routings


def _read_sheet(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    if sheet not in xl.sheet_names:
        raise ValueError(f"Sheet '{sheet}' not found in template. "
                         "Run Power Query or use --direct-mode.")
    df = xl.parse(sheet, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ────────────────────────────────────────────────────────────────────────────
# Mode 2: read directly from folder + master files (bypasses Power Query)
# ────────────────────────────────────────────────────────────────────────────

def _read_tabular(filepath: str) -> pd.DataFrame:
    """
    Read a tabular data file into a DataFrame.
    Supports .xlsx / .xls (Excel) and .csv (auto-detects separator).
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(filepath, header=0)
    elif ext == ".csv":
        # Try comma first; fall back to semicolon and tab
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(filepath, sep=sep, header=0, encoding="utf-8-sig")
                if len(df.columns) > 1:          # more than 1 col → correct sep
                    return df
            except Exception:
                continue
        # Last resort: let pandas sniff
        return pd.read_csv(filepath, sep=None, engine="python",
                           header=0, encoding="utf-8-sig")
    else:
        raise ValueError(f"Unsupported file type '{ext}' for file: {filepath}")


def _find_master_file(folder: str, stem: str) -> str:
    """
    Find master data file by stem name (e.g. 'master_capacity').
    Tries .xlsx then .csv.  Raises FileNotFoundError if neither exists.
    """
    for ext in (".xlsx", ".xls", ".csv"):
        path = os.path.join(folder, stem + ext)
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        f"Master data file '{stem}' not found in {folder}. "
        f"Expected '{stem}.xlsx', '{stem}.xls', or '{stem}.csv'."
    )


def _planner_file_rank(path: str) -> int:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return 0
    if ext == ".csv":
        return 1
    return 9


def _find_planner_files(load_folder: str) -> List[str]:
    files = (
        glob.glob(os.path.join(load_folder, "*.xlsx")) +
        glob.glob(os.path.join(load_folder, "*.xls")) +
        glob.glob(os.path.join(load_folder, "*.csv"))
    )
    if not files:
        raise FileNotFoundError(
            f"No planner files (.xlsx or .csv) found in: {load_folder}"
        )

    recognised: dict[int, str] = {}
    for fp in files:
        match = PLANNER_FILE_RE.match(os.path.basename(fp))
        if not match:
            continue
        idx = int(match.group(1))
        existing = recognised.get(idx)
        if existing is None or _planner_file_rank(fp) < _planner_file_rank(existing):
            recognised[idx] = fp

    if recognised:
        return [recognised[idx] for idx in sorted(recognised)]

    raise FileNotFoundError(
        f"No valid planner files found in: {load_folder}. "
        "Expected planner1_load to planner6_load in .xlsx/.xls/.csv format."
    )


def load_direct_mode_a(
    load_folder: str,
    master_folder: str,
    selected_scenario: str | None = None,
) -> Tuple[List[LoadRecord], List[CapacityRecord], List[RoutingRecord]]:
    """
    Mode A: load planner files + master_capacity only.
    No routing table.  Returns empty routing list.
    """
    loads, caps, _ = load_direct(
        load_folder=load_folder,
        master_folder=master_folder,
        routing_filename=None,          # skip routing
        selected_scenario=selected_scenario,
    )
    return loads, caps, []


def load_direct_mode_b(
    load_folder: str,
    master_folder: str,
    selected_scenario: str | None = None,
) -> Tuple[List[LoadRecord], List[CapacityRecord], List[RoutingRecord]]:
    """
    Mode B: load planner files + master_capacity + alternative_routing.
    Tries 'alternative_routing' first, then 'master_routing' as fallback.
    """
    return load_direct(
        load_folder=load_folder,
        master_folder=master_folder,
        routing_filename="alternative_routing",
        routing_required=True,
        selected_scenario=selected_scenario,
    )


def load_direct(
    load_folder: str,
    master_folder: str,
    capacity_filename: str = "master_capacity",
    routing_filename: str | None = "alternative_routing",
    routing_required: bool = False,
    selected_scenario: str | None = None,
) -> Tuple[List[LoadRecord], List[CapacityRecord], List[RoutingRecord]]:
    """
    Read all planner files (*.xlsx or *.csv) from load_folder, append them,
    then read master data files from master_folder.
    Supports both Excel and CSV for all input files.
    """
    # --- planner files (.xlsx and .csv) ---
    files = _find_planner_files(load_folder)

    # Parse each planner file separately so row numbers reflect the original file
    all_load_records: List[LoadRecord] = []
    found_any = False

    for fp in sorted(files):
        bn = os.path.basename(fp).lower()
        if bn.startswith("master_") or bn.startswith("alternative_routing"):
            print(f"  [SKIP] {bn} (master data file, not a planner file)")
            continue
        try:
            df = _read_tabular(fp)
            df.columns = [str(c).strip() for c in df.columns]
            df = _normalize_load_columns(df)
            required_cols = {"Month", "Product", "Forecast_Tons"}
            cols_present = set(df.columns)
            if not required_cols.issubset(cols_present):
                missing = required_cols - cols_present
                print(f"  [SKIP] {bn} – missing required columns: {missing}")
                continue
            # Drop fully-empty key rows before parsing (preserves original index for row_num)
            df = df.dropna(subset=["Month", "Product", "Forecast_Tons"], how="all")
            df = df[
                df["Month"].astype(str).str.strip().ne("") &
                df["Product"].astype(str).str.strip().ne("")
            ]
            records = _parse_load_df(df, source_file=os.path.basename(fp))
            all_load_records.extend(records)
            found_any = True
            print(f"  [READ] {bn}  ({len(records)} records)")
        except Exception as e:
            print(f"  [WARN] Could not read planner file {fp}: {e}")

    if not found_any:
        raise FileNotFoundError(
            f"No valid planner files found in: {load_folder}. "
            "Ensure files are named without 'master_' prefix and "
            "contain Month / Product / Forecast_Tons columns."
        )

    # --- master data files (.xlsx or .csv) ---
    def _resolve_master(folder: str, name_or_stem: str) -> str:
        if "." in os.path.basename(name_or_stem):
            path = os.path.join(folder, name_or_stem)
            if os.path.exists(path):
                return path
            raise FileNotFoundError(f"Master data file not found: {path}")
        return _find_master_file(folder, name_or_stem)

    cap_path = _resolve_master(master_folder, capacity_filename)
    print(f"  [READ] {os.path.basename(cap_path)}  (capacity master)")
    cap_df = _read_tabular(cap_path)
    cap_df.columns = [str(c).strip() for c in cap_df.columns]

    loads = _apply_scenario_filter(_aggregate_load_records(all_load_records), selected_scenario)
    caps  = _parse_capacity_df(cap_df)

    # Routing is optional (None = skip)
    if routing_filename is None:
        return loads, caps, []

    # Try the requested name, then fallback to master_routing for backward compat
    rout_path = None
    stems_to_try = [routing_filename]
    if routing_filename != "master_routing":
        stems_to_try.append("master_routing")

    for stem in stems_to_try:
        try:
            rout_path = _resolve_master(master_folder, stem)
            break
        except FileNotFoundError:
            continue

    if rout_path is None:
        if routing_required:
            raise FileNotFoundError(
                f"No routing file found in {master_folder}. "
                f"Tried: {stems_to_try}."
            )
        print(f"  [WARN] No routing file found (tried: {stems_to_try}). Running without routing.")
        return loads, caps, []

    print(f"  [READ] {os.path.basename(rout_path)}  (routing master)")
    rout_df = _read_tabular(rout_path)
    rout_df.columns = [str(c).strip() for c in rout_df.columns]
    routings = _parse_routing_df(rout_df)
    return loads, caps, routings


def discover_planner_scenarios(load_folder: str) -> List[str]:
    """Return distinct scenario values found in the planner files' Scenario-like columns."""
    files = _find_planner_files(load_folder)
    discovered: dict[str, str] = {}

    for fp in sorted(files):
        bn = os.path.basename(fp).lower()
        if bn.startswith("master_") or bn.startswith("alternative_routing"):
            continue
        try:
            df = _read_tabular(fp)
            df.columns = [str(c).strip() for c in df.columns]
            df = _normalize_load_columns(df)
            scenario_col = None
            if "Scenario" in df.columns:
                scenario_col = "Scenario"
            elif "ScenarioVersion" in df.columns:
                scenario_col = "ScenarioVersion"
            if scenario_col is None:
                continue
            for raw_value in df[scenario_col].tolist():
                scenario = _norm_optional_text(raw_value)
                if scenario:
                    discovered.setdefault(scenario.casefold(), scenario)
        except Exception:
            continue

    return sorted(discovered.values(), key=str.casefold)


# ────────────────────────────────────────────────────────────────────────────
# Parsers
# ────────────────────────────────────────────────────────────────────────────

def _header_key(value: str) -> str:
    """Normalise a column header for alias matching."""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _merge_text_values(*values: str | None) -> str:
    """Join distinct non-empty text values into a deterministic display string."""
    unique: dict[str, str] = {}
    for value in values:
        text = str(value or "").strip()
        if not text or text.lower() in {"nan", "none"}:
            continue
        unique.setdefault(text.casefold(), text)
    merged = [unique[key] for key in sorted(unique)]
    return " | ".join(merged)


def _col(df: pd.DataFrame, *candidates) -> pd.Series:
    """Return the first matching column (case-insensitive)."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower()
        if key in cols_lower:
            return df[cols_lower[key]]
    raise KeyError(f"None of {candidates} found in columns: {list(df.columns)}")


def _opt_col(df: pd.DataFrame, *candidates) -> pd.Series:
    try:
        return _col(df, *candidates)
    except KeyError:
        return pd.Series([None] * len(df))


def _normalize_load_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename common column name variants to the canonical names the parser expects.
    Handles differences like spaces, capitalisation, and alternate names.
    """
    rename_map = {
        "month":                 "Month",
        "planner":               "PlannerName",
        "plannername":           "PlannerName",
        "product":               "Product",
        "productfamily":         "ProductFamily",
        "plant":                 "Plant",
        "forecasttons":          "Forecast_Tons",
        # ResourceGroupOwner aliases
        "resource":              "ResourceGroupOwner",
        "resourcegroupowner":    "ResourceGroupOwner",
        "resourcegroup":         "ResourceGroupOwner",
        # Scenario aliases
        "scenario":              "Scenario",
        "scenarioname":          "Scenario",
        "case":                  "Scenario",
        # ScenarioVersion aliases
        "scenarioversion":       "ScenarioVersion",
        "version":               "ScenarioVersion",
        "comment":               "Comment",
    }
    new_cols = {}
    for col in df.columns:
        key = _header_key(col)
        if key in rename_map:
            new_cols[col] = rename_map[key]
    return df.rename(columns=new_cols)


def _parse_load_df(df: pd.DataFrame, source_file: str = "") -> List[LoadRecord]:
    df = _normalize_load_columns(df)
    records = []
    for df_idx, row in df.iterrows():
        # df_idx is 0-based; add 2 to get Excel/CSV row number (1 header + 1-based)
        row_num = int(df_idx) + 2
        try:
            ft = float(row.get("Forecast_Tons", 0) or 0)
            records.append(LoadRecord(
                month=_norm_month(str(row.get("Month", ""))),
                planner_name=str(row.get("PlannerName", "")).strip(),
                product=_norm_product(str(row.get("Product", "")).strip()),
                product_family=str(row.get("ProductFamily", "")).strip(),
                plant=str(row.get("Plant", "")).strip(),
                forecast_tons=ft,
                scenario=_norm_optional_text(row.get("Scenario", "")) or _norm_optional_text(row.get("ScenarioVersion", "")),
                resource_group_owner=str(row.get("ResourceGroupOwner", "") or "").strip() or None,
                scenario_version=str(row.get("ScenarioVersion", "") or "").strip() or None,
                comment=str(row.get("Comment", "") or "").strip() or None,
                source_file=source_file,
                row_num=row_num,
            ))
        except Exception as e:
            print(f"  [WARN] Skipping load row (行 {row_num}) due to: {e}  →  {dict(row)}")
    return records


def _aggregate_load_records(records: List[LoadRecord]) -> List[LoadRecord]:
    """
    Merge duplicate non-negative rows by key and sum Forecast_Tons.
    Negative rows stay separate so validation can still surface them.
    Metadata fields that differ are merged into a combined display value.
    """
    grouped: dict = {}
    for r in records:
        key = (
            r.month,
            r.product,
            r.planner_name,
            _scenario_key(r.scenario),
            r.forecast_tons < 0,
        )
        if key not in grouped:
            grouped[key] = r
        else:
            # Accumulate tons; keep first row's metadata
            existing = grouped[key]
            grouped[key] = LoadRecord(
                month=existing.month,
                planner_name=existing.planner_name,
                product=existing.product,
                product_family=_merge_text_values(existing.product_family, r.product_family),
                plant=_merge_text_values(existing.plant, r.plant),
                forecast_tons=existing.forecast_tons + r.forecast_tons,
                scenario=existing.scenario,
                resource_group_owner=existing.resource_group_owner,
                scenario_version=existing.scenario_version,
                comment=existing.comment,
                source_file=existing.source_file,
                row_num=existing.row_num,   # 保留首次出现的行号
            )

    aggregated = list(grouped.values())
    original_count = len(records)
    merged_count = original_count - len(aggregated)
    if merged_count > 0:
        print(f"  [INFO] 合并了 {merged_count} 条重复行（吨位已累加），"
              f"剩余 {len(aggregated)} 条有效记录")
    return aggregated


def _norm_optional_text(value) -> str | None:
    text = str(value or "").strip()
    if not text or text.lower() in {"nan", "none"}:
        return None
    return text


def _scenario_key(value: str | None) -> str | None:
    scenario = _norm_optional_text(value)
    return scenario.casefold() if scenario else None


def _apply_scenario_filter(
    records: List[LoadRecord],
    selected_scenario: str | None,
) -> List[LoadRecord]:
    scenario_key = _scenario_key(selected_scenario)
    if not scenario_key:
        return records

    if not any(_scenario_key(record.scenario) for record in records):
        return records

    filtered = [
        record
        for record in records
        if _scenario_key(record.scenario) == scenario_key
    ]
    if not filtered:
        raise ValueError(f"No planner rows found for scenario '{selected_scenario}'.")
    return filtered


def _normalize_capacity_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map common column name variants to the canonical names for capacity master."""
    rename_map = {
        "annual capacity tons":  "Annual_Capacity_Tons",
        "annual_capacity_tons":  "Annual_Capacity_Tons",
        "annualcapacitytons":    "Annual_Capacity_Tons",
        "capacity tons":         "Annual_Capacity_Tons",
        "utilization target":    "Utilization_Target",
        "utilization_target":    "Utilization_Target",
        "utilizationtarget":     "Utilization_Target",
        "util target":           "Utilization_Target",
        "resource":              "WorkCenter",
        "work center":           "WorkCenter",
        "workcenter":            "WorkCenter",
        "line":                  "WorkCenter",
        "product family":        "ProductFamily",
        "product_family":        "ProductFamily",
    }
    new_cols = {col: rename_map[col.strip().lower()]
                for col in df.columns if col.strip().lower() in rename_map}
    return df.rename(columns=new_cols)


def _normalize_routing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map common column name variants to the canonical names for routing master."""
    rename_map = {
        "resource":              "WorkCenter",
        "work center":           "WorkCenter",
        "workcenter":            "WorkCenter",
        "line":                  "WorkCenter",
        "product family":        "ProductFamily",
        "product_family":        "ProductFamily",
        "router type":           "RouteType",
        "route type":            "RouteType",
        "routetype":             "RouteType",
        "routing type":          "RouteType",
        "eligiblefalg":          "EligibleFlag",   # common typo in source data
        "eligible flag":         "EligibleFlag",
        "eligible":              "EligibleFlag",
        "penaltyweight":         "PenaltyWeight",
        "penalty weight":        "PenaltyWeight",
        "penalty":               "PenaltyWeight",
        "capacity ton":          "_CapacityTon",   # store for priority derivation
        "capacity tons":         "_CapacityTon",
    }
    new_cols = {col: rename_map[col.strip().lower()]
                for col in df.columns if col.strip().lower() in rename_map}
    return df.rename(columns=new_cols)


# Priority derived from RouteType when no explicit Priority column exists
_ROUTETYPE_PRIORITY = {
    "primary":     1,
    "alternative": 2,
    "toller":      3,
}


def _parse_capacity_df(df: pd.DataFrame) -> List[CapacityRecord]:
    df = _normalize_capacity_columns(df)
    records = []
    for _, row in df.iterrows():
        try:
            ut = float(row.get("Utilization_Target", 0.88) or 0.88)
            if ut > 1.0:
                ut = ut / 100.0     # tolerate % entry (e.g. 88 → 0.88)
            records.append(CapacityRecord(
                product=str(row.get("Product", "")).strip(),
                work_center=str(row.get("WorkCenter", "")).strip(),
                annual_capacity_tons=float(row.get("Annual_Capacity_Tons", 0) or 0),
                utilization_target=ut,
                effective_from=str(row.get("Effective_From", "") or "").strip() or None,
                effective_to=str(row.get("Effective_To", "") or "").strip() or None,
            ))
        except Exception as e:
            print(f"  [WARN] Skipping capacity row due to: {e}  →  {dict(row)}")
    return records


def _parse_routing_df(df: pd.DataFrame) -> List[RoutingRecord]:
    df = _normalize_routing_columns(df)
    has_priority_col = "Priority" in df.columns
    records = []
    for _, row in df.iterrows():
        try:
            route_type = str(row.get("RouteType", "Alternative") or "Alternative").strip()

            # Derive priority from RouteType when no Priority column present
            if has_priority_col:
                priority = int(float(row.get("Priority", 99) or 99))
            else:
                priority = _ROUTETYPE_PRIORITY.get(route_type.lower(), 99)

            # EligibleFlag: accept Y/N/True/False/1/0 and numeric (non-zero = eligible)
            raw_ef = row.get("EligibleFlag", "Y")
            ef_str = str(raw_ef).strip().upper()
            if ef_str in ("Y", "YES", "TRUE", "1"):
                eligible = True
            elif ef_str in ("N", "NO", "FALSE", "0"):
                eligible = False
            else:
                # Numeric value (e.g. 0.88) → treat as eligible if > 0
                try:
                    eligible = float(raw_ef) > 0
                except (ValueError, TypeError):
                    eligible = True   # default to eligible if unparseable

            pw = float(row.get("PenaltyWeight", 0) or 0)
            raw_prod   = row.get("Product", None)
            raw_family = row.get("ProductFamily", None)
            prod_val   = None if (raw_prod   is None or str(raw_prod).strip()   in ("", "nan", "None")) else str(raw_prod).strip()
            family_val = None if (raw_family is None or str(raw_family).strip() in ("", "nan", "None")) else str(raw_family).strip()
            records.append(RoutingRecord(
                product=prod_val,
                product_family=family_val,
                work_center=str(row.get("WorkCenter", "")).strip(),
                priority=priority,
                eligible_flag=eligible,
                route_type=route_type,
                penalty_weight=pw,
            ))
        except Exception as e:
            print(f"  [WARN] Skipping routing row due to: {e}  →  {dict(row)}")
    return records


def _norm_product(raw: str) -> str:
    """
    Strip leading zeros from purely numeric product codes.
    e.g. '00011026040' → '11026040'
    Non-numeric codes (e.g. 'SFDC^^...') are returned unchanged.
    """
    if raw.isdigit():
        return raw.lstrip("0") or raw   # keep at least one digit if all-zero
    return raw


def _norm_month(raw: str) -> str:
    """
    Normalise to YYYY-MM. Handles:
      - YYYY-MM          e.g. 2025-01
      - YYYY/MM          e.g. 2025/01
      - Excel serial no. e.g. 46023  (days since 1899-12-30)
      - Any string pandas can parse  e.g. "Jan 2025", "2025-01-15"
    """
    raw = raw.strip()
    # Already YYYY-MM
    if len(raw) == 7 and raw[4] == "-":
        return raw
    # YYYY/MM
    if len(raw) == 7 and raw[4] == "/":
        return raw.replace("/", "-")
    # Excel serial number (pure integer string, e.g. "46023")
    if raw.isdigit():
        try:
            from datetime import timedelta, date
            # Excel epoch = 1899-12-30 (accounts for Excel's leap-year-1900 bug)
            excel_epoch = date(1899, 12, 30)
            actual_date = excel_epoch + timedelta(days=int(raw))
            return f"{actual_date.year}-{actual_date.month:02d}"
        except Exception:
            pass
    # Try pandas general parsing (handles "Jan 2025", "2025-01-15", etc.)
    try:
        ts = pd.to_datetime(raw)
        return f"{ts.year}-{ts.month:02d}"
    except Exception:
        return raw
