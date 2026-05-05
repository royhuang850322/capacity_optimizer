"""
Monthly capacity optimisation using OR-Tools GLOP (linear programming).

ModeA:
  - capacity only
  - each demand node is scoped to month + product + plant + source resource
  - Stage 1 may only consume the matching source resource from master_capacity

ModeB:
  - Stage 1: run the same source-resource baseline as ModeA
  - Stage 2: reroute only the Stage 1 residual through product-level routing
  - Stage 3: classify the remaining residual as Toller or Unmet
"""
from __future__ import annotations

from typing import Dict, List, Set, Tuple

from ortools.linear_solver import pywraplp

from app.models import AllocationResult, CapacityRecord, LoadRecord, RoutingRecord


BIG_M = 1_000_000.0
PRIORITY_BASE_PENALTY = 10
EPSILON = 1e-6

DemandKey = Tuple[str, str, str, str]  # month, product, plant, source_resource
DemandMap = Dict[DemandKey, float]
EffCapMap = Dict[Tuple[str, str], float]
NodeMetaMap = Dict[DemandKey, Tuple[str, str]]  # product_family, merged planner names
EligibleRoute = Tuple[str, int, str, float]
RouteMap = Dict[DemandKey, List[EligibleRoute]]


def run_optimization_mode_a(
    months: List[str],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    verbose: bool = False,
) -> List[AllocationResult]:
    demand, node_meta = _build_demand(loads)
    eff_cap = _build_eff_cap(capacities)
    eligible = {
        demand_key: _build_capacity_only_routes(demand_key, eff_cap)
        for demand_key in demand
    }

    all_results: List[AllocationResult] = []
    for month in months:
        month_nodes = _nodes_in_month(month, demand)
        month_demand = _slice_demand(demand, month, month_nodes)
        phase_results, residual = _run_lp_for_nodes(
            month=month,
            nodes=month_nodes,
            demand=month_demand,
            full_demand=demand,
            node_meta=node_meta,
            eff_cap=eff_cap,
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
    baseline_eff_cap = _build_eff_cap(baseline_capacities)
    routing_eff_cap = _build_eff_cap(routing_capacities)

    stage1_routes = {
        demand_key: _build_capacity_only_routes(demand_key, baseline_eff_cap)
        for demand_key in demand
    }
    stage2_routes, toller_products = _build_mode_b_stage2_routes(
        all_demand_keys=set(demand),
        eff_cap=routing_eff_cap,
        routings=routings,
    )

    all_results: List[AllocationResult] = []
    for month in months:
        monthly_nodes = _nodes_in_month(month, demand)
        if not monthly_nodes:
            continue

        month_results: List[AllocationResult] = []

        stage1_results, residual_after_capacity = _run_lp_for_nodes(
            month=month,
            nodes=monthly_nodes,
            demand=_slice_demand(demand, month, monthly_nodes),
            full_demand=demand,
            node_meta=node_meta,
            eff_cap=baseline_eff_cap,
            eligible=stage1_routes,
            wc_limits={},
            allocation_source="Capacity_Base",
            verbose=verbose,
            phase_label="Capacity-Base",
        )
        month_results.extend(stage1_results)

        wc_used_after_stage1: Dict[str, float] = {}
        _accumulate_wc_used(wc_used_after_stage1, stage1_results, routing_eff_cap)

        stage2_nodes = [
            demand_key
            for demand_key, tons in residual_after_capacity.items()
            if tons > EPSILON and stage2_routes.get(demand_key)
        ]
        stage2_results, stage2_residual = _run_lp_for_nodes(
            month=month,
            nodes=stage2_nodes,
            demand=_residual_to_demand(residual_after_capacity),
            full_demand=demand,
            node_meta=node_meta,
            eff_cap=routing_eff_cap,
            eligible=stage2_routes,
            wc_limits=_remaining_wc_limits(wc_used_after_stage1),
            allocation_source="Routing_Reroute",
            verbose=verbose,
            phase_label="Routing-Reroute",
        )
        month_results.extend(stage2_results)

        residual_after_routing = dict(residual_after_capacity)
        for demand_key in stage2_nodes:
            residual_after_routing[demand_key] = stage2_residual.get(demand_key, 0.0)

        final_unmet: Dict[DemandKey, float] = {}
        final_outsourced: Dict[DemandKey, float] = {}
        outsource_residual: Dict[DemandKey, float] = {}
        unmet_residual: Dict[DemandKey, float] = {}
        for demand_key, residual_tons in residual_after_routing.items():
            if residual_tons <= EPSILON:
                continue
            product = demand_key[1]
            if product in toller_products:
                final_outsourced[demand_key] = residual_tons
                outsource_residual[demand_key] = residual_tons
            else:
                final_unmet[demand_key] = residual_tons
                unmet_residual[demand_key] = residual_tons

        month_results.extend(_build_outsource_rows(outsource_residual, demand, node_meta))
        month_results.extend(_build_unmet_rows(unmet_residual, demand, node_meta))

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


def _build_eff_cap(capacities: List[CapacityRecord]) -> EffCapMap:
    eff_cap: EffCapMap = {}
    for record in capacities:
        eff_cap[(record.product, record.work_center)] = record.effective_monthly_capacity_tons
    return eff_cap


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
    eligible: RouteMap,
    wc_limits: Dict[str, float],
    allocation_source: str = "",
    verbose: bool = False,
    phase_label: str = "",
) -> Tuple[List[AllocationResult], Dict[DemandKey, float]]:
    nodes = [demand_key for demand_key in nodes if demand_key in demand]
    if not nodes:
        return [], {}

    solver = pywraplp.Solver.CreateSolver("GLOP")
    if not solver:
        raise RuntimeError("OR-Tools GLOP solver could not be created.")
    solver.SuppressOutput()
    inf = solver.infinity()

    x: Dict[Tuple[DemandKey, str], pywraplp.Variable] = {}
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

    objective = solver.Objective()
    objective.SetMinimization()
    for demand_key in nodes:
        objective.SetCoefficient(unmet[demand_key], BIG_M)
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
            rounded_allocated = round(allocated_tons, 4)
            if allocated_tons < EPSILON or rounded_allocated <= 0.0:
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
                demand_tons=round(display_demand, 4),
                allocated_tons=rounded_allocated,
                outsourced_tons=0.0,
                unmet_tons=0.0,
                capacity_share_pct=round(100.0 * allocated_tons / cap, 2),
                planner_name=planner_names,
                allocation_source=allocation_source,
            ))

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
            demand_tons=round(full_demand.get(demand_key, unmet_tons), 4),
            allocated_tons=0.0,
            outsourced_tons=0.0,
            unmet_tons=round(unmet_tons, 4),
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
            demand_tons=round(full_demand.get(demand_key, outsourced_tons), 4),
            allocated_tons=0.0,
            outsourced_tons=round(outsourced_tons, 4),
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
        result.residual_after_capacity_tons = round(
            residual_after_capacity.get(demand_key, 0.0),
            4,
        )
        result.residual_after_routing_tons = round(
            residual_after_routing.get(demand_key, 0.0),
            4,
        )


def _apply_final_balances(
    results: List[AllocationResult],
    final_unmet: Dict[DemandKey, float],
    final_outsourced: Dict[DemandKey, float] | None = None,
) -> None:
    final_outsourced = final_outsourced or {}
    for result in results:
        demand_key = _demand_key_from_result(result)
        result.unmet_tons = round(final_unmet.get(demand_key, 0.0), 4)
        if result.allocation_type == "Outsourced":
            result.outsourced_tons = round(final_outsourced.get(demand_key, result.outsourced_tons), 4)
        else:
            result.outsourced_tons = round(result.outsourced_tons, 4)


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
