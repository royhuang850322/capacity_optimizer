"""
Reusable result-analysis helpers shared by Excel and UI workflows.
"""
from __future__ import annotations

from typing import Any

import pandas as pd


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def compute_service_level(df: pd.DataFrame, demand_col: str, supplied_col: str) -> pd.Series:
    demand = df[demand_col].replace(0, pd.NA)
    return (df[supplied_col] / demand).fillna(0.0)


def format_percent_columns_for_display(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    display_df = df.copy()
    for column in columns:
        if column not in display_df.columns:
            continue
        display_df[column] = (
            pd.to_numeric(display_df[column], errors="coerce")
            .fillna(0.0)
            .map(lambda value: f"{value:.1%}")
        )
    return display_df


def build_result_analysis(
    detail_df: pd.DataFrame,
    wc_load_df: pd.DataFrame,
    run_info_df: pd.DataFrame,
) -> dict[str, Any]:
    if detail_df.empty:
        return {
            "detail": detail_df,
            "run_info": run_info_df,
            "scenario_name": "",
            "mode_name": "",
            "planner_product_month_summary": pd.DataFrame(),
            "planner_summary": pd.DataFrame(),
            "product_month_summary": pd.DataFrame(),
            "monthly_summary": pd.DataFrame(),
            "product_summary": pd.DataFrame(),
            "wc_long": pd.DataFrame(),
            "wc_summary": pd.DataFrame(),
        }

    detail_df = detail_df.copy()
    run_info_df = run_info_df.copy()

    for col in ("Month", "PlannerName", "Product", "ProductFamily", "Plant", "AllocationType", "WorkCenter", "RouteType"):
        if col in detail_df.columns:
            detail_df[col] = detail_df[col].map(_as_text)

    for col in ("Demand_Tons", "Allocated_Tons", "Outsourced_Tons", "Unmet_Tons", "CapacityShare_Pct", "Priority"):
        if col in detail_df.columns:
            detail_df[col] = pd.to_numeric(detail_df[col], errors="coerce").fillna(0.0)

    if "PlannerName" in detail_df.columns and detail_df["PlannerName"].astype(str).str.strip().ne("").any():
        planner_product_month_summary = (
            detail_df.groupby(["Month", "Product", "PlannerName"], as_index=False)
            .agg(
                ProductFamily=("ProductFamily", "first"),
                Plant=("Plant", "first"),
                Demand_Tons=("Demand_Tons", "max"),
                Internal_Tons=("Allocated_Tons", "sum"),
                Outsourced_Tons=("Outsourced_Tons", "sum"),
                Unmet_Tons=("Unmet_Tons", "max"),
            )
            .sort_values(["Month", "PlannerName", "Product"])
        )
    else:
        planner_product_month_summary = (
            detail_df.groupby(["Month", "Product"], as_index=False)
            .agg(
                ProductFamily=("ProductFamily", "first"),
                Plant=("Plant", "first"),
                Demand_Tons=("Demand_Tons", "max"),
                Internal_Tons=("Allocated_Tons", "sum"),
                Outsourced_Tons=("Outsourced_Tons", "sum"),
                Unmet_Tons=("Unmet_Tons", "max"),
            )
            .sort_values(["Month", "Product"])
        )
        planner_product_month_summary.insert(1, "PlannerName", "")

    planner_product_month_summary["Supplied_Tons"] = (
        planner_product_month_summary["Internal_Tons"] + planner_product_month_summary["Outsourced_Tons"]
    )
    planner_product_month_summary["Service_Level"] = compute_service_level(
        planner_product_month_summary,
        "Demand_Tons",
        "Supplied_Tons",
    )

    planner_summary = (
        planner_product_month_summary.groupby("PlannerName", as_index=False)
        .agg(
            Demand_Tons=("Demand_Tons", "sum"),
            Internal_Tons=("Internal_Tons", "sum"),
            Outsourced_Tons=("Outsourced_Tons", "sum"),
            Unmet_Tons=("Unmet_Tons", "sum"),
        )
        .sort_values(["Unmet_Tons", "Demand_Tons"], ascending=[False, False])
    )
    planner_summary["Supplied_Tons"] = planner_summary["Internal_Tons"] + planner_summary["Outsourced_Tons"]
    planner_summary["Service_Level"] = compute_service_level(planner_summary, "Demand_Tons", "Supplied_Tons")

    product_month_summary = (
        planner_product_month_summary.groupby(["Month", "Product"], as_index=False)
        .agg(
            ProductFamily=("ProductFamily", "first"),
            Plant=("Plant", "first"),
            Demand_Tons=("Demand_Tons", "sum"),
            Internal_Tons=("Internal_Tons", "sum"),
            Outsourced_Tons=("Outsourced_Tons", "sum"),
            Unmet_Tons=("Unmet_Tons", "sum"),
        )
        .sort_values(["Month", "Product"])
    )
    product_month_summary["Supplied_Tons"] = (
        product_month_summary["Internal_Tons"] + product_month_summary["Outsourced_Tons"]
    )
    product_month_summary["Service_Level"] = compute_service_level(
        product_month_summary,
        "Demand_Tons",
        "Supplied_Tons",
    )

    monthly_summary = (
        product_month_summary.groupby("Month", as_index=False)
        .agg(
            Demand_Tons=("Demand_Tons", "sum"),
            Internal_Tons=("Internal_Tons", "sum"),
            Outsourced_Tons=("Outsourced_Tons", "sum"),
            Unmet_Tons=("Unmet_Tons", "sum"),
        )
        .sort_values("Month")
    )
    monthly_summary["Supplied_Tons"] = monthly_summary["Internal_Tons"] + monthly_summary["Outsourced_Tons"]
    monthly_summary["Service_Level"] = compute_service_level(monthly_summary, "Demand_Tons", "Supplied_Tons")

    product_summary = (
        planner_product_month_summary.groupby(["Product", "ProductFamily", "Plant"], as_index=False)
        .agg(
            Demand_Tons=("Demand_Tons", "sum"),
            Internal_Tons=("Internal_Tons", "sum"),
            Outsourced_Tons=("Outsourced_Tons", "sum"),
            Unmet_Tons=("Unmet_Tons", "sum"),
        )
        .sort_values(["Unmet_Tons", "Demand_Tons"], ascending=[False, False])
    )
    product_summary["Supplied_Tons"] = product_summary["Internal_Tons"] + product_summary["Outsourced_Tons"]
    product_summary["Service_Level"] = compute_service_level(product_summary, "Demand_Tons", "Supplied_Tons")

    wc_long = pd.DataFrame()
    wc_summary = pd.DataFrame()
    if not wc_load_df.empty and "WorkCenter" in wc_load_df.columns:
        wc_load_df = wc_load_df.copy()
        wc_load_df["WorkCenter"] = wc_load_df["WorkCenter"].map(_as_text)
        month_cols = [col for col in wc_load_df.columns if col != "WorkCenter"]
        for col in month_cols:
            wc_load_df[col] = pd.to_numeric(wc_load_df[col], errors="coerce").fillna(0.0)
        wc_long = wc_load_df.melt(
            id_vars="WorkCenter",
            value_vars=month_cols,
            var_name="Month",
            value_name="LoadPct",
        )
        wc_summary = (
            wc_long.groupby("WorkCenter", as_index=False)
            .agg(
                AvgLoadPct=("LoadPct", "mean"),
                PeakLoadPct=("LoadPct", "max"),
                MinLoadPct=("LoadPct", "min"),
                StdLoadPct=("LoadPct", "std"),
                Over95Months=("LoadPct", lambda values: int((values >= 0.95).sum())),
            )
            .fillna({"StdLoadPct": 0.0})
            .sort_values(["Over95Months", "PeakLoadPct", "AvgLoadPct"], ascending=[False, False, False])
        )

    run_info_map: dict[str, str] = {}
    if not run_info_df.empty and {"Parameter", "Value"}.issubset(run_info_df.columns):
        run_info_map = {
            _as_text(row["Parameter"]): _as_text(row["Value"])
            for _, row in run_info_df.iterrows()
        }

    return {
        "detail": detail_df,
        "run_info": run_info_df,
        "scenario_name": run_info_map.get("Scenario_Name", ""),
        "mode_name": run_info_map.get("Mode", ""),
        "planner_product_month_summary": planner_product_month_summary,
        "planner_summary": planner_summary,
        "product_month_summary": product_month_summary,
        "monthly_summary": monthly_summary,
        "product_summary": product_summary,
        "wc_long": wc_long,
        "wc_summary": wc_summary,
    }


def build_mode_comparison_frame(metrics_by_mode: dict[str, dict[str, Any]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for mode in ("ModeA", "ModeB"):
        metrics = metrics_by_mode.get(mode)
        if not metrics:
            continue
        rows.extend([
            {"Mode": mode, "Metric": "Internal allocated", "Value": metrics["total_internal_allocated"]},
            {"Mode": mode, "Metric": "Outsourced", "Value": metrics["total_outsourced"]},
            {"Mode": mode, "Metric": "Residual unmet", "Value": metrics["total_unmet"]},
        ])
    return pd.DataFrame(rows)


def build_executive_insights(
    preview_mode: str,
    preview_metrics: dict[str, Any],
    metrics_by_mode: dict[str, dict[str, Any]],
    analysis: dict[str, Any],
) -> list[str]:
    lines: list[str] = []
    total_supplied = preview_metrics["total_internal_allocated"] + preview_metrics["total_outsourced"]
    lines.append(
        f"{preview_mode} delivered {total_supplied:,.0f} tons out of {preview_metrics['total_demand']:,.0f}, "
        f"for a service level of {preview_metrics['service_level']:.1f}%."
    )

    monthly_summary = analysis["monthly_summary"]
    if not monthly_summary.empty:
        peak_gap = monthly_summary.sort_values(["Unmet_Tons", "Demand_Tons"], ascending=[False, False]).iloc[0]
        lines.append(
            f"The largest monthly gap is in {peak_gap['Month']}, with {peak_gap['Unmet_Tons']:,.0f} tons unmet "
            f"against {peak_gap['Demand_Tons']:,.0f} tons demand."
        )

    product_summary = analysis["product_summary"]
    if not product_summary.empty:
        top_product = product_summary.iloc[0]
        lines.append(
            f"The highest-risk product is {top_product['Product']} ({top_product['ProductFamily']}), "
            f"with {top_product['Unmet_Tons']:,.0f} tons residual unmet."
        )

    wc_summary = analysis["wc_summary"]
    if not wc_summary.empty:
        bottleneck = wc_summary.iloc[0]
        lines.append(
            f"The tightest work center is {bottleneck['WorkCenter']}, peaking at {bottleneck['PeakLoadPct']:.1%} "
            f"and running above 95% in {int(bottleneck['Over95Months'])} month(s)."
        )

    mode_a = metrics_by_mode.get("ModeA")
    mode_b = metrics_by_mode.get("ModeB")
    if mode_a and mode_b and mode_a.get("scenario_name") == mode_b.get("scenario_name"):
        unmet_delta = mode_a["total_unmet"] - mode_b["total_unmet"]
        service_delta = mode_b["service_level"] - mode_a["service_level"]
        lines.append(
            f"Current-session comparison: ModeB changes residual unmet by {unmet_delta:,.0f} tons "
            f"and service level by {service_delta:+.1f} pct versus ModeA."
        )

    return lines[:4]
