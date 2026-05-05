"""
Helpers for report-side workcenter pressure calculations.

These helpers do not change the optimizer solution. They only control how
load percentages are displayed in Excel analysis outputs.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd
from ortools.linear_solver import pywraplp

from app.capacity_basis import MAX_BASIS, PLANNED_BASIS
from app.models import AllocationResult, CapacityRecord, LoadRecord, RoutingRecord


EPSILON = 1e-6

RawCapacityMap = Dict[Tuple[str, str], float]


def build_raw_capacity_map(capacities: List[CapacityRecord]) -> RawCapacityMap:
    return {
        (record.product, record.work_center): record.monthly_capacity_tons
        for record in capacities
        if record.monthly_capacity_tons > EPSILON
    }


def compute_display_capacity_share_pct(
    product: str,
    work_center: str,
    allocated_tons: float,
    raw_capacity_map: RawCapacityMap,
) -> float:
    if allocated_tons <= EPSILON:
        return 0.0
    raw_capacity = raw_capacity_map.get((product, work_center), 0.0)
    if raw_capacity <= EPSILON:
        raise ValueError(
            f"Missing raw capacity definition for product={product}, resource={work_center}."
        )
    return round(100.0 * allocated_tons / raw_capacity, 4)


def build_pressure_load_frame(
    mode: str,
    results: List[AllocationResult],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    months: List[str],
    unmet_capacities: List[CapacityRecord] | None = None,
) -> pd.DataFrame:
    raw_capacity_map = build_raw_capacity_map(capacities)
    unmet_capacity_map = build_raw_capacity_map(unmet_capacities or capacities)
    month_wc_load = _build_internal_load_map(results, raw_capacity_map)
    unmet_by_month_product = _extract_unmet_by_month_product(results)

    if mode.strip().lower() == "modea":
        assigned_unmet_tons = _assign_mode_a_unmet_tons(
            loads=loads,
            unmet_by_month_product=unmet_by_month_product,
        )
    else:
        assigned_unmet_tons = _assign_mode_b_unmet_tons(
            results=results,
            unmet_by_month_product=unmet_by_month_product,
            candidate_capacity_map=unmet_capacity_map,
            display_capacity_map=raw_capacity_map,
        )

    for (month, product, work_center), tons in assigned_unmet_tons.items():
        month_wc_load[(month, work_center)] += _tons_to_load_share(product, work_center, tons, raw_capacity_map)

    workcenters = sorted({work_center for _month, work_center in month_wc_load})
    if not workcenters:
        return pd.DataFrame(columns=["WorkCenter", *months])

    rows = []
    for work_center in workcenters:
        row = {"WorkCenter": work_center}
        for month in months:
            row[month] = month_wc_load.get((month, work_center), 0.0)
        rows.append(row)

    return pd.DataFrame(rows, columns=["WorkCenter", *months])


def build_dashboard_fact_frame(
    mode: str,
    results: List[AllocationResult],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    unmet_capacities: List[CapacityRecord] | None = None,
) -> pd.DataFrame:
    raw_capacity_map = build_raw_capacity_map(capacities)
    unmet_capacity_map = build_raw_capacity_map(unmet_capacities or capacities)
    internal_tons_by_key = _build_internal_tons_map(results)
    unmet_by_month_product = _extract_unmet_by_month_product(results)
    outsourced_by_month_product = _extract_outsourced_by_month_product(results)

    if mode.strip().lower() == "modea":
        assigned_unmet_tons = _assign_mode_a_unmet_tons(
            loads=loads,
            unmet_by_month_product=unmet_by_month_product,
        )
        assigned_outsourced_tons: Dict[Tuple[str, str, str], float] = {}
    else:
        assigned_unmet_tons = _assign_mode_b_unmet_tons(
            results=results,
            unmet_by_month_product=unmet_by_month_product,
            candidate_capacity_map=unmet_capacity_map,
            display_capacity_map=raw_capacity_map,
        )
        assigned_outsourced_tons = _assign_mode_b_outsourced_tons(
            outsourced_by_month_product=outsourced_by_month_product,
            routings=routings,
        )

    fact_by_workcenter_year: Dict[Tuple[str, str], Dict[str, float]] = defaultdict(
        lambda: {
            "Demand_Tons": 0.0,
            "Internal_Tons": 0.0,
            "Outsourced_Tons": 0.0,
            "Unmet_Tons": 0.0,
            "Supplied_Tons": 0.0,
        }
    )

    for (month, _product, work_center), tons in internal_tons_by_key.items():
        year = _month_to_year(month)
        fact_by_workcenter_year[(year, work_center)]["Demand_Tons"] += tons
        fact_by_workcenter_year[(year, work_center)]["Internal_Tons"] += tons
        fact_by_workcenter_year[(year, work_center)]["Supplied_Tons"] += tons

    for (month, _product, work_center), tons in assigned_outsourced_tons.items():
        year = _month_to_year(month)
        fact_by_workcenter_year[(year, work_center)]["Demand_Tons"] += tons
        fact_by_workcenter_year[(year, work_center)]["Outsourced_Tons"] += tons
        fact_by_workcenter_year[(year, work_center)]["Supplied_Tons"] += tons

    for (month, _product, work_center), tons in assigned_unmet_tons.items():
        year = _month_to_year(month)
        fact_by_workcenter_year[(year, work_center)]["Demand_Tons"] += tons
        fact_by_workcenter_year[(year, work_center)]["Unmet_Tons"] += tons

    rows = []
    for year, work_center in sorted(
        fact_by_workcenter_year,
        key=lambda item: (str(item[0]).casefold(), str(item[1]).casefold()),
    ):
        payload = fact_by_workcenter_year[(year, work_center)]
        rows.append(
            {
                "Mode": mode,
                "Year": year,
                "WorkCenter": work_center,
                "Demand_Tons": round(payload["Demand_Tons"], 4),
                "Internal_Tons": round(payload["Internal_Tons"], 4),
                "Outsourced_Tons": round(payload["Outsourced_Tons"], 4),
                "Unmet_Tons": round(payload["Unmet_Tons"], 4),
                "Supplied_Tons": round(payload["Supplied_Tons"], 4),
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Mode",
            "Year",
            "WorkCenter",
            "Demand_Tons",
            "Internal_Tons",
            "Outsourced_Tons",
            "Unmet_Tons",
            "Supplied_Tons",
        ],
    )


def build_pressure_tons_frame(
    mode: str,
    results: List[AllocationResult],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    months: List[str],
    unmet_capacities: List[CapacityRecord] | None = None,
) -> pd.DataFrame:
    raw_capacity_map = build_raw_capacity_map(capacities)
    unmet_capacity_map = build_raw_capacity_map(unmet_capacities or capacities)
    month_wc_tons = _build_internal_workcenter_tons_map(results)
    unmet_by_month_product = _extract_unmet_by_month_product(results)

    if mode.strip().lower() == "modea":
        assigned_unmet_tons = _assign_mode_a_unmet_tons(
            loads=loads,
            unmet_by_month_product=unmet_by_month_product,
        )
    else:
        assigned_unmet_tons = _assign_mode_b_unmet_tons(
            results=results,
            unmet_by_month_product=unmet_by_month_product,
            candidate_capacity_map=unmet_capacity_map,
            display_capacity_map=raw_capacity_map,
        )

    for (month, _product, work_center), tons in assigned_unmet_tons.items():
        month_wc_tons[(month, work_center)] += tons

    workcenters = sorted({work_center for _month, work_center in month_wc_tons}, key=str.casefold)
    if not workcenters:
        return pd.DataFrame(columns=["WorkCenter", *months])

    rows = []
    for work_center in workcenters:
        row = {"WorkCenter": work_center}
        for month in months:
            row[month] = round(month_wc_tons.get((month, work_center), 0.0), 4)
        rows.append(row)
    return pd.DataFrame(rows, columns=["WorkCenter", *months])


def build_unmet_attribution_detail_frame(
    mode: str,
    results: List[AllocationResult],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    unmet_capacities: List[CapacityRecord] | None = None,
) -> pd.DataFrame:
    columns = [
        "Month",
        "PlannerName",
        "Product",
        "ProductFamily",
        "Plant",
        "Source_Resource",
        "Owner_WorkCenter",
        "Capacity_Candidate_WorkCenters",
        "Attributed_WorkCenter",
        "Reference_Demand_Tons",
        "Product_Unmet_Tons",
        "Attributed_Unmet_Tons",
        "Attribution_Rule",
    ]
    unmet_by_month_product = _extract_unmet_by_month_product(results)
    if not unmet_by_month_product:
        return pd.DataFrame(columns=columns)

    if mode.strip().lower() == "modea":
        rows = _build_mode_a_unmet_assignment_rows(
            loads=loads,
            unmet_by_month_product=unmet_by_month_product,
        )
    else:
        rows = _build_mode_b_unmet_assignment_rows(
            results=results,
            loads=loads,
            unmet_by_month_product=unmet_by_month_product,
            candidate_capacity_map=build_raw_capacity_map(unmet_capacities or capacities),
            display_capacity_map=build_raw_capacity_map(capacities),
        )

    if not rows:
        return pd.DataFrame(columns=columns)

    detail_df = pd.DataFrame(rows, columns=columns)
    for numeric_col in ("Reference_Demand_Tons", "Product_Unmet_Tons", "Attributed_Unmet_Tons"):
        detail_df[numeric_col] = pd.to_numeric(detail_df[numeric_col], errors="coerce").fillna(0.0).round(4)
    detail_df = detail_df.sort_values(
        ["Month", "Product", "Plant", "Source_Resource", "PlannerName", "Attributed_WorkCenter"],
        key=lambda col: col.map(lambda value: str(value).casefold()),
        kind="stable",
    ).reset_index(drop=True)
    return detail_df


def build_capacity_compare_heatmap_frames(
    mode: str,
    basis_results: Dict[str, List[AllocationResult]],
    basis_capacities: Dict[str, List[CapacityRecord]],
    loads: List[LoadRecord],
    routings: List[RoutingRecord],
    months: List[str],
    demand_basis: str = PLANNED_BASIS,
    unmet_capacities_by_basis: Dict[str, List[CapacityRecord]] | None = None,
) -> Dict[str, pd.DataFrame]:
    max_load = build_pressure_load_frame(
        mode=mode,
        results=basis_results[MAX_BASIS],
        loads=loads,
        capacities=basis_capacities[MAX_BASIS],
        routings=routings,
        months=months,
        unmet_capacities=(unmet_capacities_by_basis or {}).get(MAX_BASIS),
    )
    planner_load = build_pressure_load_frame(
        mode=mode,
        results=basis_results[PLANNED_BASIS],
        loads=loads,
        capacities=basis_capacities[PLANNED_BASIS],
        routings=routings,
        months=months,
        unmet_capacities=(unmet_capacities_by_basis or {}).get(PLANNED_BASIS),
    )
    demand_tons = build_pressure_tons_frame(
        mode=mode,
        results=basis_results[demand_basis],
        loads=loads,
        capacities=basis_capacities[demand_basis],
        routings=routings,
        months=months,
        unmet_capacities=(unmet_capacities_by_basis or {}).get(demand_basis),
    )

    monthly = _merge_heatmap_metric_frames(
        demand_tons=demand_tons,
        max_load=max_load,
        planner_load=planner_load,
        months=months,
    )
    yearly = _summarize_heatmap_months_to_years(monthly, months)
    return {
        "monthly": monthly,
        "yearly": yearly,
    }


def _build_internal_load_map(
    results: List[AllocationResult],
    raw_capacity_map: RawCapacityMap,
) -> Dict[Tuple[str, str], float]:
    month_wc_load: Dict[Tuple[str, str], float] = defaultdict(float)
    for (month, product, work_center), tons in _build_internal_tons_map(results).items():
        raw_capacity = raw_capacity_map.get((product, work_center), 0.0)
        if raw_capacity <= EPSILON:
            continue
        month_wc_load[(month, work_center)] += tons / raw_capacity
    return month_wc_load


def _build_internal_workcenter_tons_map(
    results: List[AllocationResult],
) -> Dict[Tuple[str, str], float]:
    month_wc_tons: Dict[Tuple[str, str], float] = defaultdict(float)
    for (month, _product, work_center), tons in _build_internal_tons_map(results).items():
        month_wc_tons[(month, work_center)] += tons
    return month_wc_tons


def _build_internal_tons_map(
    results: List[AllocationResult],
) -> Dict[Tuple[str, str, str], float]:
    month_product_wc_tons: Dict[Tuple[str, str, str], float] = defaultdict(float)
    for result in results:
        if result.allocation_type != "Internal":
            continue
        month_product_wc_tons[(result.month, result.product, result.work_center)] += float(result.allocated_tons or 0.0)
    return month_product_wc_tons


def _result_node_key(result: AllocationResult) -> Tuple[str, str, str, str]:
    return (
        str(result.month),
        str(result.product),
        str(result.plant),
        str(result.source_resource),
    )


def _extract_unmet_by_month_product(
    results: List[AllocationResult],
) -> Dict[Tuple[str, str, str, str], float]:
    unmet_by_month_product: Dict[Tuple[str, str, str, str], float] = {}
    for result in results:
        key = _result_node_key(result)
        unmet_by_month_product[key] = max(
            unmet_by_month_product.get(key, 0.0),
            float(result.unmet_tons or 0.0),
        )
    return {
        key: value
        for key, value in unmet_by_month_product.items()
        if value > EPSILON
    }


def _month_to_year(month_value: str) -> str:
    text = str(month_value or "").strip()
    return text[:4] if len(text) >= 4 else text


def _merge_heatmap_metric_frames(
    demand_tons: pd.DataFrame,
    max_load: pd.DataFrame,
    planner_load: pd.DataFrame,
    months: List[str],
) -> pd.DataFrame:
    workcenters = sorted(
        set(demand_tons.get("WorkCenter", pd.Series(dtype=str)).tolist())
        | set(max_load.get("WorkCenter", pd.Series(dtype=str)).tolist())
        | set(planner_load.get("WorkCenter", pd.Series(dtype=str)).tolist()),
        key=str.casefold,
    )
    metric_specs = [
        ("Demand", demand_tons),
        ("Max Load%", max_load),
        ("Planned Load%", planner_load),
    ]

    rows: list[dict[str, object]] = []
    for work_center in workcenters:
        for metric_name, frame in metric_specs:
            frame_row = frame[frame["WorkCenter"] == work_center] if not frame.empty else pd.DataFrame()
            row = {
                "WorkCenter": work_center,
                "Metric": metric_name,
            }
            for month in months:
                row[month] = (
                    float(frame_row.iloc[0][month])
                    if not frame_row.empty and month in frame_row.columns
                    else 0.0
                )
            rows.append(row)
    return pd.DataFrame(rows, columns=["WorkCenter", "Metric", *months])


def _summarize_heatmap_months_to_years(
    monthly_frame: pd.DataFrame,
    months: List[str],
) -> pd.DataFrame:
    years = list(dict.fromkeys(_month_to_year(month) for month in months))
    if monthly_frame.empty:
        return pd.DataFrame(columns=["WorkCenter", "Metric", *years])

    rows: list[dict[str, object]] = []
    for _, record in monthly_frame.iterrows():
        row = {
            "WorkCenter": record["WorkCenter"],
            "Metric": record["Metric"],
        }
        for year in years:
            year_months = [month for month in months if _month_to_year(month) == year]
            values = [float(record[month]) for month in year_months if month in monthly_frame.columns]
            if record["Metric"] == "Demand":
                row[year] = round(sum(values), 4)
            else:
                row[year] = round(sum(values) / len(values), 6) if values else 0.0
        rows.append(row)
    return pd.DataFrame(rows, columns=["WorkCenter", "Metric", *years])


def _extract_outsourced_by_month_product(
    results: List[AllocationResult],
) -> Dict[Tuple[str, str, str, str], float]:
    outsourced_by_month_product: Dict[Tuple[str, str, str, str], float] = {}
    for result in results:
        key = _result_node_key(result)
        outsourced_by_month_product[key] = max(
            outsourced_by_month_product.get(key, 0.0),
            float(result.outsourced_tons or 0.0),
        )
    return {
        key: value
        for key, value in outsourced_by_month_product.items()
        if value > EPSILON
    }


def _split_merged_text(value: str | None) -> List[str]:
    text = str(value or "").strip()
    if not text:
        return []
    parts = [part.strip() for part in text.split("|")]
    return [part for part in parts if part]


def _join_workcenters(work_centers: List[str] | set[str]) -> str:
    return " | ".join(sorted({str(value).strip() for value in work_centers if str(value).strip()}, key=str.casefold))


def _build_month_product_metadata(
    loads: List[LoadRecord],
    results: List[AllocationResult],
) -> Dict[Tuple[str, str, str, str], dict[str, object]]:
    metadata: Dict[Tuple[str, str, str, str], dict[str, object]] = {}

    for load in loads:
        key = (
            str(load.month),
            str(load.product),
            str(load.plant),
            str(load.resource_group_owner or "").strip(),
        )
        payload = metadata.setdefault(
            key,
            {
                "product_family": "",
                "plant": "",
                "planner_name": "",
                "source_resource": str(load.resource_group_owner or "").strip(),
                "reference_demand_tons": 0.0,
            },
        )
        if not payload["product_family"] and load.product_family:
            payload["product_family"] = load.product_family
        if not payload["plant"] and load.plant:
            payload["plant"] = load.plant
        if not payload["planner_name"] and load.planner_name:
            payload["planner_name"] = load.planner_name
        elif load.planner_name:
            payload["planner_name"] = _join_workcenters(
                set(_split_merged_text(str(payload["planner_name"]))) | {load.planner_name}
            )
        payload["reference_demand_tons"] = float(payload["reference_demand_tons"]) + max(float(load.forecast_tons or 0.0), 0.0)

    for result in results:
        key = _result_node_key(result)
        payload = metadata.setdefault(
            key,
            {
                "product_family": "",
                "plant": "",
                "planner_name": "",
                "source_resource": result.source_resource,
                "reference_demand_tons": 0.0,
            },
        )
        if not payload["product_family"] and result.product_family:
            payload["product_family"] = result.product_family
        if not payload["plant"] and result.plant:
            payload["plant"] = result.plant
        if not payload["planner_name"] and result.planner_name:
            payload["planner_name"] = result.planner_name
        if float(payload["reference_demand_tons"]) <= EPSILON:
            payload["reference_demand_tons"] = max(float(payload["reference_demand_tons"]), float(result.demand_tons or 0.0))

    return metadata


def _build_mode_a_unmet_assignment_rows(
    loads: List[LoadRecord],
    unmet_by_month_product: Dict[Tuple[str, str, str, str], float],
) -> List[dict[str, object]]:
    planner_month_product: Dict[Tuple[str, str, str, str], dict[str, object]] = {}
    for load in loads:
        tons = max(float(load.forecast_tons or 0.0), 0.0)
        if tons <= EPSILON:
            continue
        source_resource = str(load.resource_group_owner or "").strip()
        key = (load.month, load.product, load.plant, source_resource)
        bucket = planner_month_product.setdefault(
            key,
            {
                "tons": 0.0,
                "planner_names": set(),
                "product_family": load.product_family or "",
                "plant": load.plant or "",
                "source_resource": source_resource,
            },
        )
        bucket["tons"] = float(bucket["tons"]) + tons
        if load.planner_name:
            bucket["planner_names"].add(load.planner_name)
        if not bucket["product_family"] and load.product_family:
            bucket["product_family"] = load.product_family
        if not bucket["plant"] and load.plant:
            bucket["plant"] = load.plant

    rows: List[dict[str, object]] = []
    for (month, product, plant, source_resource), unmet_tons in unmet_by_month_product.items():
        payload = planner_month_product.get((month, product, plant, source_resource))
        if not payload:
            raise ValueError(
                f"ModeA unmet assignment could not find load node for product={product}, month={month}, "
                f"plant={plant}, resource={source_resource}."
            )
        rows.append(
            {
                "Month": month,
                "PlannerName": _join_workcenters(payload["planner_names"]),
                "Product": product,
                "ProductFamily": payload["product_family"],
                "Plant": payload["plant"],
                "Source_Resource": source_resource,
                "Owner_WorkCenter": source_resource,
                "Capacity_Candidate_WorkCenters": source_resource,
                "Attributed_WorkCenter": source_resource,
                "Reference_Demand_Tons": float(payload["tons"]),
                "Product_Unmet_Tons": unmet_tons,
                "Attributed_Unmet_Tons": unmet_tons,
                "Attribution_Rule": "ModeA source resource workcenter",
            }
        )
    return rows


def _assign_mode_a_unmet_tons(
    loads: List[LoadRecord],
    unmet_by_month_product: Dict[Tuple[str, str], float],
    ) -> Dict[Tuple[str, str, str], float]:
    assigned_tons: Dict[Tuple[str, str, str], float] = defaultdict(float)
    for row in _build_mode_a_unmet_assignment_rows(
        loads=loads,
        unmet_by_month_product=unmet_by_month_product,
    ):
        assigned_tons[(str(row["Month"]), str(row["Product"]), str(row["Attributed_WorkCenter"]))] += float(row["Attributed_Unmet_Tons"])
    return assigned_tons


def _assign_mode_b_unmet_tons(
    results: List[AllocationResult],
    unmet_by_month_product: Dict[Tuple[str, str, str, str], float],
    candidate_capacity_map: RawCapacityMap,
    display_capacity_map: RawCapacityMap,
) -> Dict[Tuple[str, str, str], float]:
    assigned_tons: Dict[Tuple[str, str, str], float] = defaultdict(float)
    for row in _build_mode_b_unmet_assignment_rows(
        results=results,
        loads=[],
        unmet_by_month_product=unmet_by_month_product,
        candidate_capacity_map=candidate_capacity_map,
        display_capacity_map=display_capacity_map,
    ):
        assigned_tons[(str(row["Month"]), str(row["Product"]), str(row["Attributed_WorkCenter"]))] += float(row["Attributed_Unmet_Tons"])
    return assigned_tons


def _build_mode_b_unmet_assignment_rows(
    results: List[AllocationResult],
    loads: List[LoadRecord],
    unmet_by_month_product: Dict[Tuple[str, str, str, str], float],
    candidate_capacity_map: RawCapacityMap,
    display_capacity_map: RawCapacityMap,
) -> List[dict[str, object]]:
    metadata_by_month_product = _build_month_product_metadata(loads, results)
    rows: List[dict[str, object]] = []
    for (month, product, plant, source_resource), unmet_tons in unmet_by_month_product.items():
        metadata = metadata_by_month_product.get(
            (month, product, plant, source_resource),
            {
                "product_family": "",
                "plant": plant,
                "planner_name": "",
                "source_resource": source_resource,
                "reference_demand_tons": 0.0,
            },
        )
        if candidate_capacity_map.get((product, source_resource), 0.0) <= EPSILON:
            raise ValueError(
                f"ModeB unmet assignment could not find baseline capacity resource for "
                f"product={product}, month={month}, plant={plant}, resource={source_resource}."
            )
        rows.append(
            {
                "Month": month,
                "PlannerName": metadata["planner_name"],
                "Product": product,
                "ProductFamily": metadata["product_family"],
                "Plant": metadata["plant"],
                "Source_Resource": source_resource,
                "Owner_WorkCenter": source_resource,
                "Capacity_Candidate_WorkCenters": source_resource,
                "Attributed_WorkCenter": source_resource,
                "Reference_Demand_Tons": float(metadata["reference_demand_tons"]),
                "Product_Unmet_Tons": float(unmet_tons),
                "Attributed_Unmet_Tons": float(unmet_tons),
                "Attribution_Rule": "ModeB stage1 source resource workcenter",
            }
        )
    return rows


def _assign_mode_b_outsourced_tons(
    outsourced_by_month_product: Dict[Tuple[str, str, str, str], float],
    routings: List[RoutingRecord],
) -> Dict[Tuple[str, str, str], float]:
    assigned_tons: Dict[Tuple[str, str, str], float] = defaultdict(float)
    toller_candidates: Dict[str, set[str]] = defaultdict(set)
    for routing in routings:
        if not routing.product:
            continue
        if not routing.eligible_flag:
            continue
        if routing.route_type.strip().lower() != "toller":
            continue
        toller_candidates[routing.product].add(routing.work_center)

    for (month, product, _plant, _source_resource), outsourced_tons in outsourced_by_month_product.items():
        work_centers = toller_candidates.get(product, set())
        if len(work_centers) != 1:
            raise ValueError(
                f"ModeB dashboard attribution requires exactly one eligible product-level Toller route "
                f"for product={product}. Found: {sorted(work_centers) or ['<missing>']}."
            )
        assigned_tons[(month, product, next(iter(work_centers)))] += outsourced_tons

    return assigned_tons


def _solve_mode_b_capacity_only_unmet(
    month: str,
    product_unmet: Dict[str, float],
    candidate_capacity_map: RawCapacityMap,
    display_capacity_map: RawCapacityMap,
    base_load_by_work_center: Dict[str, float],
) -> Dict[str, float]:
    if not product_unmet:
        return {}

    candidates_by_product: Dict[str, List[str]] = {}
    for product in product_unmet:
        candidates = sorted(
            work_center
            for (capacity_product, work_center), raw_capacity in candidate_capacity_map.items()
            if capacity_product == product and raw_capacity > EPSILON
        )
        if not candidates:
            raise ValueError(
                f"ModeB unmet assignment could not find any baseline capacity resource for product={product}, month={month}."
            )
        candidates_by_product[product] = candidates

    solver = pywraplp.Solver.CreateSolver("GLOP")
    if not solver:
        raise RuntimeError("OR-Tools GLOP solver could not be created for ModeB unmet assignment.")

    infinity = solver.infinity()
    peak_load = solver.NumVar(0.0, infinity, f"peak_load_{month}")
    assignment_vars: Dict[Tuple[str, str], pywraplp.Variable] = {}
    work_centers = sorted({work_center for candidates in candidates_by_product.values() for work_center in candidates})

    for product, unmet_tons in product_unmet.items():
        balance = solver.Constraint(unmet_tons, unmet_tons, f"unmet_balance_{month}_{product}")
        for work_center in candidates_by_product[product]:
            variable = solver.NumVar(0.0, unmet_tons, f"assign_{month}_{product}_{work_center}")
            assignment_vars[(product, work_center)] = variable
            balance.SetCoefficient(variable, 1.0)

    for work_center in work_centers:
        constraint = solver.Constraint(-infinity, -base_load_by_work_center.get(work_center, 0.0), f"peak_cap_{month}_{work_center}")
        constraint.SetCoefficient(peak_load, -1.0)
        for product, candidates in candidates_by_product.items():
            if work_center not in candidates:
                continue
            raw_capacity = display_capacity_map.get((product, work_center), 0.0)
            if raw_capacity <= EPSILON:
                raw_capacity = candidate_capacity_map[(product, work_center)]
            constraint.SetCoefficient(assignment_vars[(product, work_center)], 1.0 / raw_capacity)

    objective = solver.Objective()
    objective.SetCoefficient(peak_load, 1.0)
    objective.SetMinimization()

    status = solver.Solve()
    if status not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        raise RuntimeError(f"ModeB unmet assignment solver failed for month {month} with status {status}.")

    assigned_tons: Dict[Tuple[str, str], float] = defaultdict(float)
    for (product, work_center), variable in assignment_vars.items():
        tons = max(variable.solution_value(), 0.0)
        if tons <= EPSILON:
            continue
        assigned_tons[(product, work_center)] += tons

    return assigned_tons


def _tons_to_load_share(
    product: str,
    work_center: str,
    tons: float,
    raw_capacity_map: RawCapacityMap,
) -> float:
    raw_capacity = raw_capacity_map.get((product, work_center), 0.0)
    if raw_capacity <= EPSILON:
        raise ValueError(
            f"Missing capacity definition for product={product}, resource={work_center}."
        )
    return tons / raw_capacity
