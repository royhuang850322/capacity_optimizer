"""
Data validation for the capacity optimizer.

ModeA:
  - validates load and capacity data
  - checks that every loaded product has at least one capacity record

ModeB:
  - includes all ModeA checks
  - validates Alternative_routing data
  - checks routing coverage against capacity
"""
from collections import defaultdict
from typing import Dict, List, Set, Tuple

from app.models import CapacityRecord, LoadRecord, RoutingRecord, ValidationIssue


def validate(
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    mode: str = "ModeB",
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    require_routing = mode.strip().lower() == "modeb"

    issues.extend(_check_load_records(loads))
    issues.extend(_check_capacity_records(capacities))

    if require_routing:
        if not routings:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingMissing",
                detail="Mode B requires Alternative_routing data.",
            ))
        else:
            issues.extend(_check_routing_records(routings))

    issues.extend(_check_cross_coverage(
        loads=loads,
        capacities=capacities,
        routings=routings,
        require_routing=require_routing,
    ))
    return issues


def _row_ref(record: LoadRecord) -> str:
    parts = []
    if record.source_file:
        parts.append(record.source_file)
    if record.row_num is not None:
        parts.append(f"row {record.row_num}")
    return f"[{' '.join(parts)}] " if parts else ""


def _split_merged_text(value: str | None) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []
    return [part.strip() for part in text.split("|") if part.strip()]


def _routing_match_summary(
    record: LoadRecord,
    product_wcs: Set[str],
    family_wcs: Set[str],
    toller_eligible: bool,
) -> str:
    parts: List[str] = []

    if product_wcs:
        parts.append(
            f"matched Product='{record.product}' -> {sorted(product_wcs)}"
        )

    if family_wcs:
        family = record.product_family or "<blank>"
        parts.append(
            f"matched ProductFamily='{family}' -> {sorted(family_wcs)}"
        )

    if toller_eligible:
        parts.append("matched product-level Toller eligibility")

    if parts:
        return "; ".join(parts)

    if record.product_family:
        return (
            f"checked Product='{record.product}' and "
            f"ProductFamily='{record.product_family}'"
        )
    return f"checked Product='{record.product}'"


def _check_load_records(loads: List[LoadRecord]) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    planner_product_resources: Dict[Tuple[str, str], Set[str]] = defaultdict(set)

    for record in loads:
        ref = _row_ref(record)

        for field_name, value in (
            ("Month", record.month),
            ("Product", record.product),
            ("PlannerName", record.planner_name),
            ("Plant", record.plant),
        ):
            if not value or value.strip() in ("", "nan", "None"):
                issues.append(ValidationIssue(
                    severity="ERROR",
                    check="LoadRequired",
                    detail=f"{ref}Missing required field {field_name}.",
                ))

        if record.month and not _valid_month(record.month):
            issues.append(ValidationIssue(
                severity="ERROR",
                check="LoadMonthFormat",
                detail=f"{ref}Invalid month '{record.month}' for product '{record.product}'.",
            ))

        if record.forecast_tons < 0:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="LoadNegativeTons",
                detail=(
                    f"{ref}Forecast_Tons is negative ({record.forecast_tons}) "
                    f"for product={record.product}, month={record.month}."
                ),
            ))

        if record.forecast_tons == 0:
            issues.append(ValidationIssue(
                severity="WARNING",
                check="LoadZeroTons",
                detail=(
                    f"{ref}Forecast_Tons=0 for product {record.product}, "
                    f"month={record.month}, planner={record.planner_name}."
                ),
            ))

        resources = _split_merged_text(record.resource_group_owner)
        if not resources:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="LoadResourceMissing",
                detail=(
                    f"{ref}Missing Resource for product={record.product}, "
                    f"month={record.month}, planner={record.planner_name}."
                ),
            ))
        planner_product_resources[(record.planner_name, record.product)].update(resources)

    for (planner_name, product), resources in sorted(planner_product_resources.items()):
        if len(resources) > 1:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="LoadPlannerProductMultiResource",
                detail=(
                    f"Planner '{planner_name}' maps product '{product}' to multiple Resources: "
                    f"{sorted(resources)}."
                ),
            ))

    return issues


def _check_capacity_records(capacities: List[CapacityRecord]) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    seen: Set[Tuple[str, str]] = set()

    for record in capacities:
        key = (record.product, record.work_center)
        if key in seen:
            issues.append(ValidationIssue(
                severity="WARNING",
                check="CapacityDuplicate",
                detail=f"Duplicate capacity entry for {record.product} / {record.work_center}.",
            ))
        seen.add(key)

        if record.annual_capacity_tons <= 0:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="CapacityZero",
                detail=f"Annual_Capacity_Tons <= 0 for {record.product} / {record.work_center}.",
            ))

        if not (0.0 < record.utilization_target <= 1.0):
            issues.append(ValidationIssue(
                severity="ERROR",
                check="CapacityUtilRange",
                detail=(
                    f"Utilization_Target {record.utilization_target} is outside (0,1] "
                    f"for {record.product} / {record.work_center}."
                ),
            ))

    return issues


def _check_routing_records(routings: List[RoutingRecord]) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []

    for record in routings:
        if record.product is None and record.product_family is None:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingNoKey",
                detail=f"Routing row has neither Product nor ProductFamily for WC={record.work_center}.",
            ))
        if record.priority < 1:
            issues.append(ValidationIssue(
                severity="WARNING",
                check="RoutingPriority",
                detail=(
                    f"Priority < 1 for WC={record.work_center}, "
                    f"product={record.product or record.product_family}."
                ),
            ))

    return issues


def _check_cross_coverage(
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    require_routing: bool,
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []

    load_products = {record.product for record in loads}
    cap_keys: Set[Tuple[str, str]] = {
        (record.product, record.work_center)
        for record in capacities
    }
    cap_products = {record.product for record in capacities}

    if not require_routing or not routings:
        for product in load_products:
            if product not in cap_products:
                issues.append(ValidationIssue(
                    severity="ERROR",
                    check="NoCoverageCapacity",
                    detail=f"Product '{product}' in load has no capacity record.",
                ))
        issues.extend(_check_planner_resource_capacity_coverage(loads, cap_keys))
        return issues

    cap_wcs_by_product: Dict[str, Set[str]] = {}
    for record in capacities:
        cap_wcs_by_product.setdefault(record.product, set()).add(record.work_center)

    routing_by_product_all: Dict[str, Set[str]] = {}
    routing_by_family_all: Dict[str, Set[str]] = {}
    routing_by_product_internal: Dict[str, Set[str]] = {}
    routing_by_family_internal: Dict[str, Set[str]] = {}
    toller_products: Set[str] = set()
    for routing in routings:
        if routing.product:
            routing_by_product_all.setdefault(routing.product, set()).add(routing.work_center)
            if routing.eligible_flag and routing.route_type.strip().lower() == "toller":
                toller_products.add(routing.product)
            elif routing.eligible_flag:
                routing_by_product_internal.setdefault(routing.product, set()).add(routing.work_center)
        elif routing.product_family:
            routing_by_family_all.setdefault(routing.product_family, set()).add(routing.work_center)
            if routing.eligible_flag and routing.route_type.strip().lower() != "toller":
                routing_by_family_internal.setdefault(routing.product_family, set()).add(routing.work_center)

    for product in load_products:
        if product not in cap_products and product not in toller_products:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="NoCoverageCapacity",
                detail=f"Product '{product}' in load has no capacity record.",
            ))

    issues.extend(_check_planner_resource_capacity_coverage(loads, cap_keys))
    issues.extend(_check_mode_b_product_primary_routes(load_products, routings))
    issues.extend(_check_mode_b_product_toller_routes(load_products, routings))

    for record in loads:
        ref = _row_ref(record)
        product_wcs = routing_by_product_internal.get(record.product, set())
        family_wcs = routing_by_family_internal.get(record.product_family, set())
        toller_eligible = record.product in toller_products
        matched_any = bool(
            routing_by_product_all.get(record.product, set()) or
            routing_by_family_all.get(record.product_family, set())
        )
        eligible_wcs = product_wcs | family_wcs
        routing_summary = _routing_match_summary(record, product_wcs, family_wcs, toller_eligible)
        if not matched_any:
            # No routing row applies to this product/family.
            # ModeB falls back to capacity-only routing in the optimizer.
            continue

        if not eligible_wcs and not toller_eligible:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="NoCoverageRouting",
                detail=(
                    f"{ref}Product '{record.product}' has no eligible routing; "
                    f"{routing_summary}."
                ),
            ))
            continue

        wcs_with_capacity = {
            wc for wc in eligible_wcs
            if (record.product, wc) in cap_keys
        }
        if not wcs_with_capacity and eligible_wcs and not toller_eligible:
            available_cap_wcs = sorted(cap_wcs_by_product.get(record.product, set()))
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingCapacityMismatch",
                detail=(
                    f"{ref}Product '{record.product}' has no matching capacity record; "
                    f"{routing_summary}; combined eligible routes={sorted(eligible_wcs)}; "
                    f"available capacity routes={available_cap_wcs}."
                ),
            ))

    return issues


def _check_planner_resource_capacity_coverage(
    loads: List[LoadRecord],
    cap_keys: Set[Tuple[str, str]],
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    seen: Set[Tuple[str, str]] = set()

    for record in loads:
        for resource in _split_merged_text(record.resource_group_owner):
            key = (record.product, resource)
            if key in seen:
                continue
            seen.add(key)
            if key not in cap_keys:
                issues.append(ValidationIssue(
                    severity="ERROR",
                    check="PlannerResourceMissingInCapacity",
                    detail=(
                        f"Planner load references Product='{record.product}' / Resource='{resource}' "
                        "but no matching master_capacity row exists."
                    ),
                ))
    return issues


def _check_mode_b_product_primary_routes(
    load_products: Set[str],
    routings: List[RoutingRecord],
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    product_level_rows: Dict[str, List[RoutingRecord]] = defaultdict(list)
    for routing in routings:
        if routing.product:
            product_level_rows[routing.product].append(routing)

    for product in sorted(load_products):
        rows = product_level_rows.get(product, [])
        if not rows:
            continue

        primary_wcs = {
            row.work_center
            for row in rows
            if row.eligible_flag and row.route_type.strip().lower() == "primary"
        }
        if len(primary_wcs) == 0:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingPrimaryMissing",
                detail=f"Product '{product}' has product-level routing rows but no eligible Primary route.",
            ))
        elif len(primary_wcs) > 1:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingPrimaryDuplicate",
                detail=f"Product '{product}' has multiple eligible Primary routes: {sorted(primary_wcs)}.",
            ))

    return issues


def _check_mode_b_product_toller_routes(
    load_products: Set[str],
    routings: List[RoutingRecord],
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    product_level_rows: Dict[str, List[RoutingRecord]] = defaultdict(list)
    for routing in routings:
        if routing.product:
            product_level_rows[routing.product].append(routing)

    for product in sorted(load_products):
        rows = product_level_rows.get(product, [])
        if not rows:
            continue

        toller_wcs = {
            row.work_center
            for row in rows
            if row.eligible_flag and row.route_type.strip().lower() == "toller"
        }
        if len(toller_wcs) > 1:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingTollerDuplicate",
                detail=f"Product '{product}' has multiple eligible Toller routes: {sorted(toller_wcs)}.",
            ))

    return issues


def _valid_month(month: str) -> bool:
    if len(month) != 7 or month[4] != "-":
        return False
    try:
        year, month_num = int(month[:4]), int(month[5:])
        return 1 <= month_num <= 12 and 2000 <= year <= 2100
    except ValueError:
        return False


def has_errors(issues: List[ValidationIssue]) -> bool:
    return any(issue.severity == "ERROR" for issue in issues)


def format_issue_report(
    issues: List[ValidationIssue],
    warning_example_limit: int = 5,
) -> List[str]:
    errors = [issue for issue in issues if issue.severity == "ERROR"]
    warnings = [issue for issue in issues if issue.severity == "WARNING"]
    lines: List[str] = []

    if errors:
        lines.append(f"\n  [X] {len(errors)} validation ERROR(s):")
        for issue in errors:
            lines.append(f"    [{issue.check}] {issue.detail}")

    if warnings:
        lines.append(f"\n  [!] {len(warnings)} validation WARNING(s):")
        warnings_by_check: Dict[str, List[ValidationIssue]] = defaultdict(list)
        for issue in warnings:
            warnings_by_check[issue.check].append(issue)

        for check, grouped in sorted(
            warnings_by_check.items(),
            key=lambda item: (-len(item[1]), item[0]),
        ):
            if len(grouped) <= warning_example_limit:
                for issue in grouped:
                    lines.append(f"    [{issue.check}] {issue.detail}")
                continue

            lines.append(f"    [{check}] {len(grouped)} occurrence(s)")
            for issue in grouped[:warning_example_limit]:
                lines.append(f"      e.g. {issue.detail}")
            lines.append(f"      ... and {len(grouped) - warning_example_limit} more")

    if not issues:
        lines.append("  [OK] All validation checks passed.")

    return lines


def print_issues(issues: List[ValidationIssue]) -> None:
    for line in format_issue_report(issues):
        print(line)
