"""
Data validation for the capacity optimizer.

ModeA:
  - validates load and capacity data
  - checks that every loaded product has at least one capacity record

ModeB:
  - includes all ModeA checks
  - requires product-level routing data
  - enforces 1 product -> 1 family
  - allows routing-only resources for Stage 2
"""
from collections import defaultdict
from typing import Dict, List, Set, Tuple

from app.models import CapacityRecord, LoadRecord, RoutingRecord, ValidationIssue


def validate(
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    mode: str = "ModeB",
    routing_capacities: List[CapacityRecord] | None = None,
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    require_routing = mode.strip().lower() == "modeb"

    issues.extend(_check_load_records(loads))
    issues.extend(_check_capacity_records(capacities))
    issues.extend(_check_product_family_consistency(loads, routings))

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
        routing_capacities=routing_capacities,
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

    return issues


def _check_product_family_consistency(
    loads: List[LoadRecord],
    routings: List[RoutingRecord],
) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    family_values_by_product: Dict[str, Set[str]] = defaultdict(set)

    for record in loads:
        product = str(record.product or "").strip()
        family = str(record.product_family or "").strip()
        if product and family:
            family_values_by_product[product].add(family)

    for routing in routings:
        product = str(routing.product or "").strip()
        family = str(routing.product_family or "").strip()
        if product and family:
            family_values_by_product[product].add(family)

    for product, families in sorted(family_values_by_product.items()):
        if len(families) > 1:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="ProductFamilyConflict",
                detail=f"Product '{product}' maps to multiple ProductFamily values: {sorted(families)}.",
            ))

    return issues


def _check_routing_records(routings: List[RoutingRecord]) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []

    for record in routings:
        if not record.product:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingProductRequired",
                detail=(
                    f"Routing row for WC={record.work_center} must provide Product. "
                    "ProductFamily-only routing is no longer used by ModeB."
                ),
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
    routing_capacities: List[CapacityRecord] | None,
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

    routing_capacities = routing_capacities or capacities
    routing_cap_keys: Set[Tuple[str, str]] = {
        (record.product, record.work_center)
        for record in routing_capacities
    }

    product_level_rows: Dict[str, List[RoutingRecord]] = defaultdict(list)
    toller_products: Set[str] = set()
    internal_route_wcs: Dict[str, Set[str]] = defaultdict(set)

    for routing in routings:
        if not routing.product:
            continue
        product_level_rows[routing.product].append(routing)
        if not routing.eligible_flag:
            continue
        if routing.route_type.strip().lower() == "toller":
            toller_products.add(routing.product)
        else:
            internal_route_wcs[routing.product].add(routing.work_center)

    issues.extend(_check_planner_resource_capacity_coverage(loads, cap_keys))
    issues.extend(_check_mode_b_product_toller_routes(load_products, routings))

    for product in sorted(load_products):
        has_stage1_capacity = product in cap_products
        has_stage2_internal = bool(internal_route_wcs.get(product))
        has_toller = product in toller_products

        if not has_stage1_capacity and not has_stage2_internal and not has_toller:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="NoCoverageCapacity",
                detail=(
                    f"Product '{product}' has no Stage 1 capacity, no eligible internal routing, "
                    "and no eligible Toller route."
                ),
            ))

    for product, route_wcs in sorted(internal_route_wcs.items()):
        missing_capacity = sorted(
            wc for wc in route_wcs
            if (product, wc) not in routing_cap_keys
        )
        if missing_capacity:
            issues.append(ValidationIssue(
                severity="ERROR",
                check="RoutingCapacityMissing",
                detail=(
                    f"Product '{product}' has eligible product-level routing to resources without capacity rows: "
                    f"{missing_capacity}."
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
