"""
Monthly capacity optimisation using OR-Tools GLOP (linear programming).

ModeA:
  - capacity only
  - each demand node is scoped to month + product + plant + source resource
  - Stage 1 may only consume the matching source resource from master_capacity

ModeB:
  - solve source capacity, product-level routing, Toller, and Unmet in one global LP
  - objective priority is Unmet first, then Toller, then route preference
  - stage trace fields are retained for report compatibility
"""
from __future__ import annotations

import calendar
from datetime import datetime, timedelta
from typing import Dict, List, Set, Tuple

from ortools.linear_solver import pywraplp

from app.models import AllocationResult, CapacityRecord, LoadRecord, RoutingRecord


BIG_M = 1_000_000_000.0
SETUP_TRIGGER_PENALTY = 1_000_000.0
SETUP_HOURS_PENALTY = 1_000.0
SETUP_TONS_PENALTY = 1.0
TOLLER_PENALTY = 100_000.0
ALTERNATIVE_PENALTY = 1_000.0
PRIMARY_PENALTY = 10.0
CAPACITY_BASE_PENALTY = 0.0
PRIORITY_BASE_PENALTY = 10
EPSILON = 1e-6
REPORT_DATA_DECIMALS = 10

DemandKey = Tuple[str, str, str, str]  # month, product, plant, source_resource
DemandMap = Dict[DemandKey, float]
EffCapMap = Dict[Tuple[str, str], float]
SetupShareMap = Dict[Tuple[str, str], float]
SetupTonsMap = Dict[Tuple[str, str], float]
SetupHoursMap = Dict[Tuple[str, str], float]
SetupGroupKey = Tuple[str, str, str]  # product_family, plant, work_center
SetupGroupShareMap = Dict[SetupGroupKey, float]
SetupGroupTonsMap = Dict[SetupGroupKey, float]
SetupGroupHoursMap = Dict[SetupGroupKey, float]
NodeMetaMap = Dict[DemandKey, Tuple[str, str]]  # product_family, merged planner names
EligibleRoute = Tuple[str, int, str, float]
RouteMap = Dict[DemandKey, List[EligibleRoute]]
GlobalRoute = Tuple[str, int, str, float, str, float]


def _report_number(value: float) -> float:
    return round(float(value or 0.0), REPORT_DATA_DECIMALS)


def _create_mip_solver() -> pywraplp.Solver:
    solver = pywraplp.Solver.CreateSolver("SCIP")
    if not solver:
        solver = pywraplp.Solver.CreateSolver("CBC_MIXED_INTEGER_PROGRAMMING")
    if not solver:
        raise RuntimeError("OR-Tools mixed-integer solver could not be created.")
    solver.SuppressOutput()
    return solver


def run_optimization_mode_a(
    months: List[str],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    verbose: bool = False,
) -> List[AllocationResult]:
    demand, node_meta = _build_demand(loads)

    all_results: List[AllocationResult] = []
    for month in months:
        eff_cap = _build_eff_cap(capacities, month)
        setup_share, setup_equiv_tons, setup_hours = _build_setup_maps(capacities, month)
        eligible = {
            demand_key: _build_capacity_only_routes(demand_key, eff_cap)
            for demand_key in demand
        }
        month_nodes = _nodes_in_month(month, demand)
        month_demand = _slice_demand(demand, month, month_nodes)
        phase_results, residual = _run_lp_for_nodes(
            month=month,
            nodes=month_nodes,
            demand=month_demand,
            full_demand=demand,
            node_meta=node_meta,
            eff_cap=eff_cap,
            setup_share=setup_share,
            setup_equiv_tons=setup_equiv_tons,
            setup_hours=setup_hours,
            eligible=eligible,
            wc_limits={},
            verbose=verbose,
        )

        month_results = list(phase_results)
        month_results.extend(_build_unmet_rows(residual, demand, node_meta, allocation_source=""))
        _apply_final_balances(
            month_results,
            final_unmet=residual,
        )
        all_results.extend(month_results)

    return all_results


def run_optimization_mode_b(
    months: List[str],
    loads: List[LoadRecord],
    baseline_capacities: List[CapacityRecord],
    routing_capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    verbose: bool = False,
) -> Tuple[List[AllocationResult], Set[str]]:
    demand, node_meta = _build_demand(loads)
    toller_products = _eligible_toller_products(routings)

    all_results: List[AllocationResult] = []
    for month in months:
        monthly_nodes = _nodes_in_month(month, demand)
        if not monthly_nodes:
            continue

        baseline_eff_cap = _build_eff_cap(baseline_capacities, month)
        routing_eff_cap = _build_eff_cap(routing_capacities, month)
        routing_setup_share, routing_setup_equiv_tons, routing_setup_hours = _build_setup_maps(routing_capacities, month)
        month_results = _run_global_routing_lp_for_nodes(
            month=month,
            nodes=monthly_nodes,
            demand=_slice_demand(demand, month, monthly_nodes),
            full_demand=demand,
            node_meta=node_meta,
            baseline_eff_cap=baseline_eff_cap,
            routing_eff_cap=routing_eff_cap,
            setup_share=routing_setup_share,
            setup_equiv_tons=routing_setup_equiv_tons,
            setup_hours=routing_setup_hours,
            routings=routings,
            verbose=verbose,
        )
        all_results.extend(month_results)

    return all_results, toller_products


def run_optimization(months, loads, capacities, routings, verbose=False):
    results, _ = run_optimization_mode_b(
        months,
        loads,
        capacities,
        capacities,
        routings,
        verbose,
    )
    return results


def _merge_meta_text(existing: str, incoming: str) -> str:
    values: Dict[str, str] = {}
    for raw_value in (existing, incoming):
        text = str(raw_value or "").strip()
        if not text:
            continue
        values.setdefault(text.casefold(), text)
    return " | ".join(values[key] for key in sorted(values))


def _demand_key_from_load(record: LoadRecord) -> DemandKey:
    return (
        record.month,
        record.product,
        record.plant,
        str(record.resource_group_owner or "").strip(),
    )


def _demand_key_from_result(result: AllocationResult) -> DemandKey:
    return (
        result.month,
        result.product,
        result.plant,
        result.source_resource,
    )


def _build_demand(loads: List[LoadRecord]) -> Tuple[DemandMap, NodeMetaMap]:
    demand: DemandMap = {}
    node_meta: NodeMetaMap = {}
    for record in loads:
        forecast_tons = max(record.forecast_tons, 0.0)
        key = _demand_key_from_load(record)
        demand[key] = demand.get(key, 0.0) + forecast_tons
        if key not in node_meta:
            node_meta[key] = (record.product_family, record.planner_name)
            continue

        existing_family, existing_planners = node_meta[key]
        node_meta[key] = (
            _merge_meta_text(existing_family, record.product_family),
            _merge_meta_text(existing_planners, record.planner_name),
        )
    return demand, node_meta


def _setup_group_key(
    *,
    product_family: str,
    product: str,
    plant: str,
    work_center: str,
) -> SetupGroupKey:
    family = str(product_family or "").strip() or str(product or "").strip()
    return (family, str(plant or "").strip(), str(work_center or "").strip())


def _setup_group_key_for_demand(
    demand_key: DemandKey,
    work_center: str,
    node_meta: NodeMetaMap,
) -> SetupGroupKey:
    _month, product, plant, _source_resource = demand_key
    product_family, _planner_names = node_meta.get(demand_key, ("", ""))
    return _setup_group_key(
        product_family=product_family,
        product=product,
        plant=plant,
        work_center=work_center,
    )


def _build_eff_cap(capacities: List[CapacityRecord], month: str | None = None) -> EffCapMap:
    if month:
        capacities = _select_effective_capacity_records(capacities, month)
    eff_cap: EffCapMap = {}
    for record in capacities:
        eff_cap[(record.product, record.work_center)] = record.effective_monthly_capacity_tons
    return eff_cap


def _build_setup_maps(
    capacities: List[CapacityRecord],
    month: str,
) -> Tuple[SetupShareMap, SetupTonsMap, SetupHoursMap]:
    month_hours = _month_hours(month)
    setup_share: SetupShareMap = {}
    setup_equiv_tons: SetupTonsMap = {}
    setup_hours_map: SetupHoursMap = {}
    for record in _select_effective_capacity_records(capacities, month):
        setup_hours = float(record.setup_hours or 0.0)
        if setup_hours <= EPSILON:
            continue
        key = (record.product, record.work_center)
        setup_share[key] = setup_hours / month_hours
        setup_equiv_tons[key] = setup_hours * record.setup_reference_monthly_capacity_tons / month_hours
        setup_hours_map[key] = setup_hours
    return setup_share, setup_equiv_tons, setup_hours_map


def _month_hours(month: str) -> float:
    year = int(str(month)[:4])
    month_num = int(str(month)[5:7])
    return float(calendar.monthrange(year, month_num)[1] * 24)


def _select_effective_capacity_records(
    capacities: List[CapacityRecord],
    month: str,
) -> List[CapacityRecord]:
    month_key = _month_to_index(month)
    grouped: Dict[Tuple[str, str], List[CapacityRecord]] = {}
    for record in capacities:
        grouped.setdefault((record.product, record.work_center), []).append(record)

    selected: List[CapacityRecord] = []
    for records in grouped.values():
        concrete_matches = [
            record
            for record in records
            if _capacity_window_covers(record, month_key)
        ]
        if concrete_matches:
            selected.append(concrete_matches[0])
            continue

        default_matches = [
            record
            for record in records
            if _is_default_capacity_window(record)
        ]
        if default_matches:
            selected.append(default_matches[0])
    return selected


def _capacity_window_covers(record: CapacityRecord, month_key: int) -> bool:
    if _is_default_capacity_window(record):
        return False
    parsed = _capacity_window(record)
    if parsed is None:
        return False
    start_month, end_month = parsed
    return start_month <= month_key <= end_month


def _is_default_capacity_window(record: CapacityRecord) -> bool:
    start = _capacity_date_text(record.effective_from)
    end = _capacity_date_text(record.effective_to)
    if not start and not end:
        return True
    return start == "99999" and end == "99999"


def _capacity_window(record: CapacityRecord) -> Tuple[int, int] | None:
    start = _capacity_date_text(record.effective_from)
    end = _capacity_date_text(record.effective_to)
    if not start or not end or start == "99999" or end == "99999":
        return None
    try:
        return _parse_capacity_month(start), _parse_capacity_month(end)
    except ValueError:
        return None


def _parse_capacity_month(value: str) -> int:
    text = _capacity_date_text(value)
    if not text:
        raise ValueError("empty capacity month")
    if text == "99999":
        raise ValueError("default capacity marker is not a concrete month")
    if text.isdigit() and len(text) > 4:
        base = datetime(1899, 12, 30)
        dt = base + timedelta(days=int(text))
        return dt.year * 12 + dt.month
    normalized = text.replace("/", "-").replace(".", "-")
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(normalized, fmt)
            return dt.year * 12 + dt.month
        except ValueError:
            pass
    dt = datetime.fromisoformat(normalized)
    return dt.year * 12 + dt.month


def _capacity_date_text(value) -> str:
    text = str(value or "").strip()
    if text.casefold() in {"", "nan", "none", "nat"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _month_to_index(month: str) -> int:
    year, month_num = int(str(month)[:4]), int(str(month)[5:7])
    return year * 12 + month_num


def _build_capacity_only_routes(
    demand_key: DemandKey,
    eff_cap: EffCapMap,
) -> List[EligibleRoute]:
    product = demand_key[1]
    source_resource = demand_key[3]
    cap = eff_cap.get((product, source_resource), 0.0)
    if cap <= 0 or not source_resource:
        return []
    return [(source_resource, 1, "Capacity", 1.0)]


def _build_mode_b_stage2_routes(
    all_demand_keys: Set[DemandKey],
    eff_cap: EffCapMap,
    routings: List[RoutingRecord],
) -> Tuple[RouteMap, Set[str]]:
    product_rows: Dict[str, List[RoutingRecord]] = {}
    toller_products: Set[str] = set()

    for routing in routings:
        if not routing.product:
            continue
        product_rows.setdefault(routing.product, []).append(routing)
        if routing.eligible_flag and _is_toller_route(routing.route_type):
            toller_products.add(routing.product)

    routes_by_product: Dict[str, List[EligibleRoute]] = {}
    for product, rows in product_rows.items():
        routes: Dict[str, Tuple[int, str, float]] = {}
        for routing in rows:
            if not routing.eligible_flag or _is_toller_route(routing.route_type):
                continue
            wc = routing.work_center
            if not wc or eff_cap.get((product, wc), 0.0) <= 0:
                continue
            candidate = (routing.priority, routing.route_type, _route_penalty(routing))
            existing = routes.get(wc)
            if existing is None or candidate[0] < existing[0]:
                routes[wc] = candidate

        routes_by_product[product] = [
            (wc, priority, route_type, penalty)
            for wc, (priority, route_type, penalty) in sorted(
                routes.items(),
                key=lambda item: (item[1][0], item[0]),
            )
        ]

    return (
        {
            demand_key: list(routes_by_product.get(demand_key[1], []))
            for demand_key in all_demand_keys
        },
        toller_products,
    )


def _run_global_routing_lp_for_nodes(
    month: str,
    nodes: List[DemandKey],
    demand: DemandMap,
    full_demand: DemandMap,
    node_meta: NodeMetaMap,
    baseline_eff_cap: EffCapMap,
    routing_eff_cap: EffCapMap,
    setup_share: SetupShareMap,
    setup_equiv_tons: SetupTonsMap,
    setup_hours: SetupHoursMap,
    routings: List[RoutingRecord],
    verbose: bool = False,
) -> List[AllocationResult]:
    nodes = [demand_key for demand_key in nodes if demand_key in demand]
    if not nodes:
        return []

    routes = _build_global_routes(nodes, baseline_eff_cap, routing_eff_cap, routings)
    toller_products = _eligible_toller_products(routings)

    solver = _create_mip_solver()
    inf = solver.infinity()

    x: Dict[Tuple[DemandKey, str, str], pywraplp.Variable] = {}
    setup_used: Dict[SetupGroupKey, pywraplp.Variable] = {}
    setup_group_share: SetupGroupShareMap = {}
    setup_group_equiv_tons: SetupGroupTonsMap = {}
    setup_group_hours: SetupGroupHoursMap = {}
    outsource: Dict[DemandKey, pywraplp.Variable] = {}
    unmet: Dict[DemandKey, pywraplp.Variable] = {}

    for demand_key in nodes:
        demand_tons = demand[demand_key]
        unmet[demand_key] = solver.NumVar(0.0, demand_tons, f"unmet_{hash(demand_key)}")
        if demand_key[1] in toller_products:
            outsource[demand_key] = solver.NumVar(0.0, demand_tons, f"out_{hash(demand_key)}")
        for wc, _priority, _route_type, _penalty, allocation_source, cap in routes.get(demand_key, []):
            if cap > 0:
                x[(demand_key, wc, allocation_source)] = solver.NumVar(
                    0.0,
                    demand_tons,
                    f"x_{hash(demand_key)}_{wc}_{allocation_source}",
                )

    for demand_key in nodes:
        demand_tons = demand[demand_key]
        constraint = solver.Constraint(demand_tons, demand_tons, f"dem_{hash(demand_key)}")
        constraint.SetCoefficient(unmet[demand_key], 1.0)
        if demand_key in outsource:
            constraint.SetCoefficient(outsource[demand_key], 1.0)
        for wc, _priority, _route_type, _penalty, allocation_source, _cap in routes.get(demand_key, []):
            variable = x.get((demand_key, wc, allocation_source))
            if variable is not None:
                constraint.SetCoefficient(variable, 1.0)

    setup_group_variables: Dict[SetupGroupKey, List[Tuple[pywraplp.Variable, str, str]]] = {}
    for (demand_key, wc, _allocation_source), variable in x.items():
        setup_key = _setup_group_key_for_demand(demand_key, wc, node_meta)
        setup_group_variables.setdefault(setup_key, []).append((variable, demand_key[1], wc))

    for setup_key, variables in setup_group_variables.items():
        group_share = max((setup_share.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        if group_share <= EPSILON:
            continue
        group_equiv_tons = max((setup_equiv_tons.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        group_hours = max((setup_hours.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        setup_group_share[setup_key] = group_share
        setup_group_equiv_tons[setup_key] = group_equiv_tons
        setup_group_hours[setup_key] = group_hours
        setup_name = abs(hash((month, *setup_key)))
        setup_var = solver.IntVar(0.0, 1.0, f"setup_{setup_name}")
        setup_used[setup_key] = setup_var
        link = solver.Constraint(-inf, 0.0, f"setup_link_{setup_name}")
        for variable, _product, _wc in variables:
            link.SetCoefficient(variable, 1.0)
        link.SetCoefficient(setup_var, -sum(demand.values()))

    all_wcs = {wc for _demand_key, wc, _source in x}
    for wc in all_wcs:
        constraint = solver.Constraint(-inf, 1.0, f"cap_{wc}")
        for demand_key in nodes:
            for route_wc, _priority, _route_type, _penalty, allocation_source, cap in routes.get(demand_key, []):
                if route_wc != wc:
                    continue
                variable = x.get((demand_key, wc, allocation_source))
                if variable is not None and cap > 0:
                    constraint.SetCoefficient(variable, 1.0 / cap)
        for setup_key, setup_var in setup_used.items():
            _family, _plant, setup_wc = setup_key
            if setup_wc == wc:
                constraint.SetCoefficient(setup_var, setup_group_share.get(setup_key, 0.0))

    objective = solver.Objective()
    objective.SetMinimization()
    for demand_key, variable in unmet.items():
        objective.SetCoefficient(variable, BIG_M)
    for demand_key, variable in outsource.items():
        objective.SetCoefficient(variable, TOLLER_PENALTY)
    for key, variable in setup_used.items():
        objective.SetCoefficient(
            variable,
            SETUP_TRIGGER_PENALTY
            + setup_group_hours.get(key, 0.0) * SETUP_HOURS_PENALTY
            + setup_group_equiv_tons.get(key, 0.0) * SETUP_TONS_PENALTY,
        )
    for (demand_key, wc, allocation_source), variable in x.items():
        route = _find_global_route(demand_key, wc, allocation_source, routes)
        if route is None:
            continue
        _wc, priority, route_type, penalty, _allocation_source, _cap = route
        objective.SetCoefficient(variable, _global_route_objective_penalty(route_type, priority, penalty))

    status = solver.Solve()
    if status not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        print(f"  [WARN] Solver status {status} for month {month} Global-Routing")

    internal_results: List[AllocationResult] = []
    capacity_base_allocated: Dict[DemandKey, float] = {}
    final_unmet: Dict[DemandKey, float] = {}
    final_outsourced: Dict[DemandKey, float] = {}

    for demand_key in nodes:
        demand_month, product, plant, source_resource = demand_key
        display_demand = full_demand.get(demand_key, demand[demand_key])
        family, planner_names = node_meta.get(demand_key, ("", ""))
        final_unmet[demand_key] = max(unmet[demand_key].solution_value(), 0.0)
        if demand_key in outsource:
            final_outsourced[demand_key] = max(outsource[demand_key].solution_value(), 0.0)

        for wc, priority, route_type, _penalty, allocation_source, cap in routes.get(demand_key, []):
            variable = x.get((demand_key, wc, allocation_source))
            if variable is None:
                continue
            allocated_tons = max(variable.solution_value(), 0.0)
            stored_allocated = _report_number(allocated_tons)
            if allocated_tons < EPSILON or stored_allocated <= 0.0:
                continue
            if allocation_source == "Capacity_Base":
                capacity_base_allocated[demand_key] = capacity_base_allocated.get(demand_key, 0.0) + allocated_tons
            internal_results.append(AllocationResult(
                month=demand_month,
                product=product,
                product_family=family,
                plant=plant,
                source_resource=source_resource,
                allocation_type="Internal",
                work_center=wc,
                route_type=route_type,
                priority=priority,
                demand_tons=_report_number(display_demand),
                allocated_tons=stored_allocated,
                outsourced_tons=0.0,
                unmet_tons=0.0,
                capacity_share_pct=_report_number(100.0 * allocated_tons / cap),
                planner_name=planner_names,
                allocation_source=allocation_source,
                capacity_used_tons=stored_allocated,
            ))

    _apply_setup_to_internal_results(
        internal_results,
        setup_used=setup_used,
        setup_share=setup_group_share,
        setup_equiv_tons=setup_group_equiv_tons,
        setup_hours=setup_group_hours,
    )

    outsource_rows = _build_outsource_rows(final_outsourced, full_demand, node_meta)
    unmet_rows = _build_unmet_rows(final_unmet, full_demand, node_meta)
    month_results = [*internal_results, *outsource_rows, *unmet_rows]

    residual_after_capacity = {
        demand_key: max(demand[demand_key] - capacity_base_allocated.get(demand_key, 0.0), 0.0)
        for demand_key in nodes
    }
    residual_after_routing = {
        demand_key: final_unmet.get(demand_key, 0.0)
        for demand_key in nodes
    }
    _apply_stage_trace(
        month_results,
        month=month,
        residual_after_capacity=residual_after_capacity,
        residual_after_routing=residual_after_routing,
    )
    _apply_final_balances(
        month_results,
        final_unmet=final_unmet,
        final_outsourced=final_outsourced,
    )

    if verbose:
        total_unmet = sum(final_unmet.values())
        total_outsourced = sum(final_outsourced.values())
        print(
            f"  {month} [Global-Routing] {len(nodes)} nodes | "
            f"unmet = {total_unmet:,.1f} tons | outsourced = {total_outsourced:,.1f} tons | "
            f"obj = {objective.Value():,.2f}"
        )

    return month_results


def _build_global_routes(
    nodes: List[DemandKey],
    baseline_eff_cap: EffCapMap,
    routing_eff_cap: EffCapMap,
    routings: List[RoutingRecord],
) -> Dict[DemandKey, List[GlobalRoute]]:
    route_rows_by_product: Dict[str, List[RoutingRecord]] = {}
    for routing in routings:
        if routing.product and routing.eligible_flag and not _is_toller_route(routing.route_type):
            route_rows_by_product.setdefault(routing.product, []).append(routing)

    routes: Dict[DemandKey, List[GlobalRoute]] = {}
    for demand_key in nodes:
        product = demand_key[1]
        source_resource = demand_key[3]
        route_map: Dict[Tuple[str, str], GlobalRoute] = {}

        baseline_cap = baseline_eff_cap.get((product, source_resource), 0.0)
        if source_resource and baseline_cap > 0:
            route_map[(source_resource, "Capacity_Base")] = (
                source_resource,
                1,
                "Capacity",
                1.0,
                "Capacity_Base",
                baseline_cap,
            )

        for routing in route_rows_by_product.get(product, []):
            wc = routing.work_center
            cap = routing_eff_cap.get((product, wc), 0.0)
            if not wc or cap <= 0:
                continue
            if wc == source_resource and baseline_cap > 0:
                continue
            route = (
                wc,
                routing.priority,
                routing.route_type,
                _route_penalty(routing),
                "Routing_Reroute",
                cap,
            )
            key = (wc, "Routing_Reroute")
            existing = route_map.get(key)
            if existing is None or route[1] < existing[1]:
                route_map[key] = route

        routes[demand_key] = sorted(
            route_map.values(),
            key=lambda route: (_global_route_objective_penalty(route[2], route[1], route[3]), route[0], route[4]),
        )
    return routes


def _eligible_toller_products(routings: List[RoutingRecord]) -> Set[str]:
    return {
        routing.product
        for routing in routings
        if routing.product and routing.eligible_flag and _is_toller_route(routing.route_type)
    }


def _find_global_route(
    demand_key: DemandKey,
    wc: str,
    allocation_source: str,
    routes: Dict[DemandKey, List[GlobalRoute]],
) -> GlobalRoute | None:
    for route in routes.get(demand_key, []):
        if route[0] == wc and route[4] == allocation_source:
            return route
    return None


def _global_route_objective_penalty(route_type: str, priority: int, penalty: float) -> float:
    normalized = str(route_type or "").strip().lower()
    if normalized == "capacity":
        base = CAPACITY_BASE_PENALTY
    elif normalized == "alternative":
        base = ALTERNATIVE_PENALTY
    else:
        base = PRIMARY_PENALTY
    return base + max(float(penalty or 0.0), 0.0) + max(priority - 1, 0) * 0.01


def _apply_setup_to_internal_results(
    results: List[AllocationResult],
    *,
    setup_used: Dict[SetupGroupKey, pywraplp.Variable],
    setup_share: SetupGroupShareMap,
    setup_equiv_tons: SetupGroupTonsMap,
    setup_hours: SetupGroupHoursMap,
) -> None:
    applied: Set[SetupGroupKey] = set()
    for result in results:
        if result.allocation_type != "Internal":
            continue
        key = _setup_group_key(
            product_family=result.product_family,
            product=result.product,
            plant=result.plant,
            work_center=result.work_center,
        )
        result.capacity_used_tons = _report_number(result.allocated_tons)
        setup_var = setup_used.get(key)
        if setup_var is None or setup_var.solution_value() < 0.5 or key in applied:
            continue
        applied.add(key)
        setup_equiv = setup_equiv_tons.get(key, 0.0)
        result.setup_applied = True
        result.setup_hours = _report_number(setup_hours.get(key, 0.0))
        result.setup_equivalent_tons_by_max = _report_number(setup_equiv)
        result.capacity_used_tons = _report_number(result.allocated_tons + setup_equiv)
        if setup_share.get(key, 0.0) > EPSILON:
            result.capacity_share_pct = _report_number(
                float(result.capacity_share_pct or 0.0) + 100.0 * setup_share.get(key, 0.0)
            )


def _route_penalty(routing: RoutingRecord) -> float:
    if routing.penalty_weight > 0:
        return routing.penalty_weight
    return float(PRIORITY_BASE_PENALTY ** (routing.priority - 1))


def _nodes_in_month(month: str, demand: DemandMap) -> List[DemandKey]:
    return [demand_key for demand_key in demand if demand_key[0] == month]


def _slice_demand(
    demand: DemandMap,
    month: str,
    nodes: List[DemandKey],
) -> DemandMap:
    return {
        demand_key: demand[demand_key]
        for demand_key in nodes
        if demand_key[0] == month and demand.get(demand_key, 0.0) > EPSILON
    }


def _residual_to_demand(residual: Dict[DemandKey, float]) -> DemandMap:
    return {
        demand_key: tons
        for demand_key, tons in residual.items()
        if tons > EPSILON
    }


def _run_lp_for_nodes(
    month: str,
    nodes: List[DemandKey],
    demand: DemandMap,
    full_demand: DemandMap,
    node_meta: NodeMetaMap,
    eff_cap: EffCapMap,
    setup_share: SetupShareMap,
    setup_equiv_tons: SetupTonsMap,
    setup_hours: SetupHoursMap,
    eligible: RouteMap,
    wc_limits: Dict[str, float],
    allocation_source: str = "",
    verbose: bool = False,
    phase_label: str = "",
) -> Tuple[List[AllocationResult], Dict[DemandKey, float]]:
    nodes = [demand_key for demand_key in nodes if demand_key in demand]
    if not nodes:
        return [], {}

    solver = _create_mip_solver()
    inf = solver.infinity()

    x: Dict[Tuple[DemandKey, str], pywraplp.Variable] = {}
    setup_used: Dict[SetupGroupKey, pywraplp.Variable] = {}
    setup_group_share: SetupGroupShareMap = {}
    setup_group_equiv_tons: SetupGroupTonsMap = {}
    setup_group_hours: SetupGroupHoursMap = {}
    unmet: Dict[DemandKey, pywraplp.Variable] = {}

    for demand_key in nodes:
        product = demand_key[1]
        demand_tons = demand[demand_key]
        unmet[demand_key] = solver.NumVar(0.0, demand_tons, f"unmet_{hash(demand_key)}")
        for wc, _priority, _route_type, _penalty in eligible.get(demand_key, []):
            cap = eff_cap.get((product, wc), 0.0)
            if cap > 0:
                x[(demand_key, wc)] = solver.NumVar(0.0, demand_tons, f"x_{hash(demand_key)}_{wc}")

    for demand_key in nodes:
        demand_tons = demand[demand_key]
        constraint = solver.Constraint(demand_tons, demand_tons, f"dem_{hash(demand_key)}")
        constraint.SetCoefficient(unmet[demand_key], 1.0)
        for wc, _priority, _route_type, _penalty in eligible.get(demand_key, []):
            if (demand_key, wc) in x:
                constraint.SetCoefficient(x[(demand_key, wc)], 1.0)

    setup_group_variables: Dict[SetupGroupKey, List[Tuple[pywraplp.Variable, str, str]]] = {}
    for (demand_key, wc), variable in x.items():
        setup_key = _setup_group_key_for_demand(demand_key, wc, node_meta)
        setup_group_variables.setdefault(setup_key, []).append((variable, demand_key[1], wc))

    for setup_key, variables in setup_group_variables.items():
        group_share = max((setup_share.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        if group_share <= EPSILON:
            continue
        group_equiv_tons = max((setup_equiv_tons.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        group_hours = max((setup_hours.get((product, wc), 0.0) for _variable, product, wc in variables), default=0.0)
        setup_group_share[setup_key] = group_share
        setup_group_equiv_tons[setup_key] = group_equiv_tons
        setup_group_hours[setup_key] = group_hours
        setup_name = abs(hash((month, *setup_key)))
        setup_var = solver.IntVar(0.0, 1.0, f"setup_{setup_name}")
        setup_used[setup_key] = setup_var
        link = solver.Constraint(-inf, 0.0, f"setup_link_{setup_name}")
        for variable, _product, _wc in variables:
            link.SetCoefficient(variable, 1.0)
        link.SetCoefficient(setup_var, -sum(demand.values()))

    all_wcs = {wc for _demand_key, wc in x}
    for wc in all_wcs:
        limit = wc_limits.get(wc, 1.0)
        if limit <= 0:
            for demand_key in nodes:
                if (demand_key, wc) in x:
                    x[(demand_key, wc)].SetUb(0.0)
            continue

        constraint = solver.Constraint(-inf, limit, f"cap_{wc}")
        for demand_key in nodes:
            product = demand_key[1]
            if (demand_key, wc) in x:
                constraint.SetCoefficient(x[(demand_key, wc)], 1.0 / eff_cap[(product, wc)])
        for setup_key, setup_var in setup_used.items():
            _family, _plant, setup_wc = setup_key
            if setup_wc == wc:
                constraint.SetCoefficient(setup_var, setup_group_share.get(setup_key, 0.0))

    objective = solver.Objective()
    objective.SetMinimization()
    for demand_key in nodes:
        objective.SetCoefficient(unmet[demand_key], BIG_M)
    for key, variable in setup_used.items():
        objective.SetCoefficient(
            variable,
            SETUP_TRIGGER_PENALTY
            + setup_group_hours.get(key, 0.0) * SETUP_HOURS_PENALTY
            + setup_group_equiv_tons.get(key, 0.0) * SETUP_TONS_PENALTY,
        )
    for (demand_key, wc), variable in x.items():
        penalty = _get_penalty(demand_key, wc, eligible)
        product = demand_key[1]
        objective.SetCoefficient(variable, penalty / eff_cap[(product, wc)])

    status = solver.Solve()
    if status not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        print(f"  [WARN] Solver status {status} for month {month} {phase_label}")

    results: List[AllocationResult] = []
    residual: Dict[DemandKey, float] = {}
    total_unmet = 0.0

    for demand_key in nodes:
        demand_month, product, plant, source_resource = demand_key
        phase_demand_tons = demand[demand_key]
        display_demand = full_demand.get(demand_key, phase_demand_tons)
        family, planner_names = node_meta.get(demand_key, ("", ""))
        unmet_tons = max(unmet[demand_key].solution_value(), 0.0)
        residual[demand_key] = unmet_tons
        total_unmet += unmet_tons

        for wc, priority, route_type, _penalty in eligible.get(demand_key, []):
            if (demand_key, wc) not in x:
                continue
            allocated_tons = max(x[(demand_key, wc)].solution_value(), 0.0)
            stored_allocated = _report_number(allocated_tons)
            if allocated_tons < EPSILON or stored_allocated <= 0.0:
                continue
            cap = eff_cap[(product, wc)]
            results.append(AllocationResult(
                month=demand_month,
                product=product,
                product_family=family,
                plant=plant,
                source_resource=source_resource,
                allocation_type="Internal",
                work_center=wc,
                route_type=route_type,
                priority=priority,
                demand_tons=_report_number(display_demand),
                allocated_tons=stored_allocated,
                outsourced_tons=0.0,
                unmet_tons=0.0,
                capacity_share_pct=_report_number(100.0 * allocated_tons / cap),
                planner_name=planner_names,
                allocation_source=allocation_source,
                capacity_used_tons=stored_allocated,
            ))

    _apply_setup_to_internal_results(
        results,
        setup_used=setup_used,
        setup_share=setup_group_share,
        setup_equiv_tons=setup_group_equiv_tons,
        setup_hours=setup_group_hours,
    )

    if verbose:
        label = f"[{phase_label}] " if phase_label else ""
        print(
            f"  {month} {label}{len(nodes)} nodes | "
            f"remaining = {total_unmet:,.1f} tons | obj = {objective.Value():,.2f}"
        )

    return results, residual


def _build_unmet_rows(
    residual: Dict[DemandKey, float],
    full_demand: DemandMap,
    node_meta: NodeMetaMap,
    allocation_source: str = "Unmet",
) -> List[AllocationResult]:
    rows: List[AllocationResult] = []
    for demand_key, unmet_tons in residual.items():
        if unmet_tons <= EPSILON:
            continue
        month, product, plant, source_resource = demand_key
        family, planner_names = node_meta.get(demand_key, ("", ""))
        rows.append(AllocationResult(
            month=month,
            product=product,
            product_family=family,
            plant=plant,
            source_resource=source_resource,
            allocation_type="Unmet",
            work_center="[UNALLOCATED]",
            route_type="N/A",
            priority=99,
            demand_tons=_report_number(full_demand.get(demand_key, unmet_tons)),
            allocated_tons=0.0,
            outsourced_tons=0.0,
            unmet_tons=_report_number(unmet_tons),
            capacity_share_pct=0.0,
            planner_name=planner_names,
            allocation_source=allocation_source,
        ))
    return rows


def _build_outsource_rows(
    residual: Dict[DemandKey, float],
    full_demand: DemandMap,
    node_meta: NodeMetaMap,
    allocation_source: str = "Toller",
) -> List[AllocationResult]:
    rows: List[AllocationResult] = []
    for demand_key, outsourced_tons in residual.items():
        if outsourced_tons <= EPSILON:
            continue
        month, product, plant, source_resource = demand_key
        family, planner_names = node_meta.get(demand_key, ("", ""))
        rows.append(AllocationResult(
            month=month,
            product=product,
            product_family=family,
            plant=plant,
            source_resource=source_resource,
            allocation_type="Outsourced",
            work_center="[OUTSOURCED]",
            route_type="Toller",
            priority=99,
            demand_tons=_report_number(full_demand.get(demand_key, outsourced_tons)),
            allocated_tons=0.0,
            outsourced_tons=_report_number(outsourced_tons),
            unmet_tons=0.0,
            capacity_share_pct=0.0,
            planner_name=planner_names,
            allocation_source=allocation_source,
        ))
    return rows


def _apply_stage_trace(
    results: List[AllocationResult],
    month: str,
    residual_after_capacity: Dict[DemandKey, float],
    residual_after_routing: Dict[DemandKey, float],
) -> None:
    for result in results:
        if result.month != month:
            continue
        demand_key = _demand_key_from_result(result)
        result.residual_after_capacity_tons = _report_number(residual_after_capacity.get(demand_key, 0.0))
        result.residual_after_routing_tons = _report_number(residual_after_routing.get(demand_key, 0.0))


def _apply_final_balances(
    results: List[AllocationResult],
    final_unmet: Dict[DemandKey, float],
    final_outsourced: Dict[DemandKey, float] | None = None,
) -> None:
    final_outsourced = final_outsourced or {}
    for result in results:
        demand_key = _demand_key_from_result(result)
        result.unmet_tons = _report_number(final_unmet.get(demand_key, 0.0))
        if result.allocation_type == "Outsourced":
            result.outsourced_tons = _report_number(final_outsourced.get(demand_key, result.outsourced_tons))
        else:
            result.outsourced_tons = _report_number(result.outsourced_tons)


def _accumulate_wc_used(
    wc_used: Dict[str, float],
    results: List[AllocationResult],
    eff_cap: EffCapMap,
) -> None:
    for result in results:
        if result.allocation_type != "Internal":
            continue
        cap = eff_cap.get((result.product, result.work_center), 0.0)
        if cap <= 0:
            continue
        wc_used[result.work_center] = (
            wc_used.get(result.work_center, 0.0) +
            result.allocated_tons / cap
        )


def _remaining_wc_limits(wc_used: Dict[str, float]) -> Dict[str, float]:
    return {
        wc: max(0.0, 1.0 - used)
        for wc, used in wc_used.items()
    }


def _is_toller_route(route_type: str) -> bool:
    return route_type.strip().lower() == "toller"


def _get_penalty(
    demand_key: DemandKey,
    wc: str,
    eligible: RouteMap,
) -> float:
    for route_wc, _priority, _route_type, penalty in eligible.get(demand_key, []):
        if route_wc == wc:
            return penalty
    return float(PRIORITY_BASE_PENALTY ** 2)
