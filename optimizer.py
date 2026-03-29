"""
Monthly capacity optimisation using OR-Tools GLOP (linear programming).

ModeA:
  - capacity only
  - each (product, work center) row from master_capacity is eligible

ModeB:
  - routing aware
  - regular products are allocated in two internal passes:
      1. Primary / Capacity routes
      2. Alternative / lower-priority internal routes
  - products with a product-level Toller route are scheduled after all other products
  - residual demand for Toller products is converted to outsourced tons
"""
from __future__ import annotations

from typing import Dict, List, Set, Tuple

from ortools.linear_solver import pywraplp

from models import AllocationResult, CapacityRecord, LoadRecord, RoutingRecord


BIG_M = 1_000_000.0
PRIORITY_BASE_PENALTY = 10
EPSILON = 1e-6

DemandMap = Dict[Tuple[str, str], float]
EffCapMap = Dict[Tuple[str, str], float]
ProductMetaMap = Dict[str, Tuple[str, str]]
EligibleRoute = Tuple[str, int, str, float]
RouteMap = Dict[str, List[EligibleRoute]]


def run_optimization_mode_a(
    months: List[str],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    verbose: bool = False,
) -> List[AllocationResult]:
    demand, product_meta = _build_demand(loads)
    eff_cap = _build_eff_cap(capacities)
    eligible = {
        product: _build_capacity_only_routes(product, eff_cap)
        for _month, product in demand
    }

    all_results: List[AllocationResult] = []
    for month in months:
        month_products = _products_in_month(month, demand)
        month_demand = _slice_demand(demand, month, month_products)
        phase_results, residual = _run_lp_for_products(
            month=month,
            products=month_products,
            demand=month_demand,
            full_demand=demand,
            product_meta=product_meta,
            eff_cap=eff_cap,
            eligible=eligible,
            wc_limits={},
            verbose=verbose,
        )

        month_results = list(phase_results)
        month_results.extend(_build_unmet_rows(month, residual, demand, product_meta))
        _apply_final_balances(
            month_results,
            final_unmet={(month, product): tons for product, tons in residual.items()},
        )
        all_results.extend(month_results)

    return all_results


def run_optimization_mode_b(
    months: List[str],
    loads: List[LoadRecord],
    capacities: List[CapacityRecord],
    routings: List[RoutingRecord],
    verbose: bool = False,
) -> Tuple[List[AllocationResult], Set[str]]:
    demand, product_meta = _build_demand(loads)
    eff_cap = _build_eff_cap(capacities)
    primary_routes, secondary_routes, toller_products = _build_mode_b_routes(
        product_meta=product_meta,
        demand=demand,
        eff_cap=eff_cap,
        routings=routings,
    )

    all_results: List[AllocationResult] = []
    for month in months:
        monthly_products = _products_in_month(month, demand)
        if not monthly_products:
            continue

        regular_products = [product for product in monthly_products if product not in toller_products]
        toller_only_products = [product for product in monthly_products if product in toller_products]

        wc_used: Dict[str, float] = {}
        month_results: List[AllocationResult] = []

        regular_primary_results, regular_primary_residual = _run_lp_for_products(
            month=month,
            products=regular_products,
            demand=_slice_demand(demand, month, regular_products),
            full_demand=demand,
            product_meta=product_meta,
            eff_cap=eff_cap,
            eligible=primary_routes,
            wc_limits={},
            verbose=verbose,
            phase_label="Regular-Primary",
        )
        month_results.extend(regular_primary_results)
        _accumulate_wc_used(wc_used, regular_primary_results, eff_cap)

        regular_alt_results, regular_final_residual = _run_lp_for_products(
            month=month,
            products=regular_products,
            demand=_residual_to_demand(month, regular_primary_residual),
            full_demand=demand,
            product_meta=product_meta,
            eff_cap=eff_cap,
            eligible=secondary_routes,
            wc_limits=_remaining_wc_limits(wc_used),
            verbose=verbose,
            phase_label="Regular-Alternative",
        )
        month_results.extend(regular_alt_results)
        _accumulate_wc_used(wc_used, regular_alt_results, eff_cap)

        toller_primary_results, toller_primary_residual = _run_lp_for_products(
            month=month,
            products=toller_only_products,
            demand=_slice_demand(demand, month, toller_only_products),
            full_demand=demand,
            product_meta=product_meta,
            eff_cap=eff_cap,
            eligible=primary_routes,
            wc_limits=_remaining_wc_limits(wc_used),
            verbose=verbose,
            phase_label="Toller-Primary",
        )
        month_results.extend(toller_primary_results)
        _accumulate_wc_used(wc_used, toller_primary_results, eff_cap)

        toller_alt_results, toller_final_residual = _run_lp_for_products(
            month=month,
            products=toller_only_products,
            demand=_residual_to_demand(month, toller_primary_residual),
            full_demand=demand,
            product_meta=product_meta,
            eff_cap=eff_cap,
            eligible=secondary_routes,
            wc_limits=_remaining_wc_limits(wc_used),
            verbose=verbose,
            phase_label="Toller-Alternative",
        )
        month_results.extend(toller_alt_results)
        _accumulate_wc_used(wc_used, toller_alt_results, eff_cap)

        month_results.extend(_build_unmet_rows(month, regular_final_residual, demand, product_meta))
        month_results.extend(_build_outsource_rows(month, toller_final_residual, demand, product_meta))

        final_unmet = {
            (month, product): tons
            for product, tons in regular_final_residual.items()
        }
        final_outsourced = {
            (month, product): tons
            for product, tons in toller_final_residual.items()
        }
        _apply_final_balances(month_results, final_unmet, final_outsourced)

        all_results.extend(month_results)

    return all_results, toller_products


def run_optimization(months, loads, capacities, routings, verbose=False):
    results, _ = run_optimization_mode_b(months, loads, capacities, routings, verbose)
    return results


def _merge_meta_text(existing: str, incoming: str) -> str:
    values: Dict[str, str] = {}
    for raw_value in (existing, incoming):
        text = str(raw_value or "").strip()
        if not text:
            continue
        values.setdefault(text.casefold(), text)
    return " | ".join(values[key] for key in sorted(values))


def _build_demand(loads: List[LoadRecord]) -> Tuple[DemandMap, ProductMetaMap]:
    demand: DemandMap = {}
    product_meta: ProductMetaMap = {}
    for record in loads:
        forecast_tons = max(record.forecast_tons, 0.0)
        key = (record.month, record.product)
        demand[key] = demand.get(key, 0.0) + forecast_tons
        if record.product not in product_meta:
            product_meta[record.product] = (record.product_family, record.plant)
            continue

        existing_family, existing_plant = product_meta[record.product]
        product_meta[record.product] = (
            _merge_meta_text(existing_family, record.product_family),
            _merge_meta_text(existing_plant, record.plant),
        )
    return demand, product_meta


def _build_eff_cap(capacities: List[CapacityRecord]) -> EffCapMap:
    eff_cap: EffCapMap = {}
    for record in capacities:
        eff_cap[(record.product, record.work_center)] = record.effective_monthly_capacity_tons
    return eff_cap


def _build_capacity_only_routes(
    product: str,
    eff_cap: EffCapMap,
) -> List[EligibleRoute]:
    return sorted(
        (
            (wc, 1, "Capacity", 1.0)
            for (prod, wc), cap in eff_cap.items()
            if prod == product and cap > 0
        ),
        key=lambda item: item[0],
    )


def _build_mode_b_routes(
    product_meta: ProductMetaMap,
    demand: DemandMap,
    eff_cap: EffCapMap,
    routings: List[RoutingRecord],
) -> Tuple[RouteMap, RouteMap, Set[str]]:
    routing_by_product: Dict[str, List[RoutingRecord]] = {}
    routing_by_family: Dict[str, List[RoutingRecord]] = {}
    for routing in routings:
        if routing.product:
            routing_by_product.setdefault(routing.product, []).append(routing)
        elif routing.product_family:
            routing_by_family.setdefault(routing.product_family, []).append(routing)

    primary_routes: RouteMap = {}
    secondary_routes: RouteMap = {}
    toller_products: Set[str] = set()

    all_products = {product for _month, product in demand}
    for product in all_products:
        family, _plant = product_meta.get(product, ("", ""))
        family_rows = routing_by_family.get(family, [])
        product_rows = routing_by_product.get(product, [])
        matched_any = bool(family_rows or product_rows)

        if any(_is_toller_route(row.route_type) and row.eligible_flag for row in product_rows):
            toller_products.add(product)

        if matched_any:
            internal_routes = _build_internal_routes(
                product=product,
                eff_cap=eff_cap,
                family_rows=family_rows,
                product_rows=product_rows,
            )
        else:
            internal_routes = _build_capacity_only_routes(product, eff_cap)

        primary_routes[product] = [
            route
            for route in internal_routes
            if _is_primary_like(route[2])
        ]
        secondary_routes[product] = [
            route
            for route in internal_routes
            if not _is_primary_like(route[2])
        ]

    return primary_routes, secondary_routes, toller_products


def _build_internal_routes(
    product: str,
    eff_cap: EffCapMap,
    family_rows: List[RoutingRecord],
    product_rows: List[RoutingRecord],
) -> List[EligibleRoute]:
    routes: Dict[str, Tuple[int, str, float]] = {}
    blocked_wcs: Set[str] = set()

    for routing in family_rows:
        _merge_family_route(
            product=product,
            eff_cap=eff_cap,
            routes=routes,
            blocked_wcs=blocked_wcs,
            routing=routing,
        )

    for routing in product_rows:
        _override_product_route(
            product=product,
            eff_cap=eff_cap,
            routes=routes,
            blocked_wcs=blocked_wcs,
            routing=routing,
        )

    return [
        (wc, priority, route_type, penalty)
        for wc, (priority, route_type, penalty) in sorted(
            routes.items(),
            key=lambda item: (item[1][0], item[0]),
        )
    ]


def _route_penalty(routing: RoutingRecord) -> float:
    if routing.penalty_weight > 0:
        return routing.penalty_weight
    return float(PRIORITY_BASE_PENALTY ** (routing.priority - 1))


def _merge_family_route(
    product: str,
    eff_cap: EffCapMap,
    routes: Dict[str, Tuple[int, str, float]],
    blocked_wcs: Set[str],
    routing: RoutingRecord,
) -> None:
    if _is_toller_route(routing.route_type):
        return

    wc = routing.work_center
    if not wc:
        return
    if not routing.eligible_flag:
        blocked_wcs.add(wc)
        routes.pop(wc, None)
        return
    if wc in blocked_wcs or (product, wc) not in eff_cap:
        return

    candidate = (routing.priority, routing.route_type, _route_penalty(routing))
    existing = routes.get(wc)
    if existing is None or candidate[0] < existing[0]:
        routes[wc] = candidate


def _override_product_route(
    product: str,
    eff_cap: EffCapMap,
    routes: Dict[str, Tuple[int, str, float]],
    blocked_wcs: Set[str],
    routing: RoutingRecord,
) -> None:
    if _is_toller_route(routing.route_type):
        return

    wc = routing.work_center
    if not wc:
        return
    if not routing.eligible_flag:
        blocked_wcs.add(wc)
        routes.pop(wc, None)
        return

    blocked_wcs.discard(wc)
    if (product, wc) not in eff_cap:
        return
    routes[wc] = (routing.priority, routing.route_type, _route_penalty(routing))


def _products_in_month(month: str, demand: DemandMap) -> List[str]:
    return [product for (bucket, product) in demand if bucket == month]


def _slice_demand(
    demand: DemandMap,
    month: str,
    products: List[str],
) -> DemandMap:
    return {
        (month, product): demand[(month, product)]
        for product in products
        if (month, product) in demand and demand[(month, product)] > EPSILON
    }


def _residual_to_demand(month: str, residual: Dict[str, float]) -> DemandMap:
    return {
        (month, product): tons
        for product, tons in residual.items()
        if tons > EPSILON
    }


def _run_lp_for_products(
    month: str,
    products: List[str],
    demand: DemandMap,
    full_demand: DemandMap,
    product_meta: ProductMetaMap,
    eff_cap: EffCapMap,
    eligible: RouteMap,
    wc_limits: Dict[str, float],
    verbose: bool = False,
    phase_label: str = "",
) -> Tuple[List[AllocationResult], Dict[str, float]]:
    products = [product for product in products if (month, product) in demand]
    if not products:
        return [], {}

    solver = pywraplp.Solver.CreateSolver("GLOP")
    if not solver:
        raise RuntimeError("OR-Tools GLOP solver could not be created.")
    solver.SuppressOutput()
    inf = solver.infinity()

    x: Dict[Tuple[str, str], pywraplp.Variable] = {}
    unmet: Dict[str, pywraplp.Variable] = {}

    for product in products:
        demand_tons = demand[(month, product)]
        unmet[product] = solver.NumVar(0.0, demand_tons, f"unmet_{product}")
        for wc, _priority, _route_type, _penalty in eligible.get(product, []):
            cap = eff_cap.get((product, wc), 0.0)
            if cap > 0:
                x[(product, wc)] = solver.NumVar(0.0, demand_tons, f"x_{product}_{wc}")

    for product in products:
        demand_tons = demand[(month, product)]
        constraint = solver.Constraint(demand_tons, demand_tons, f"dem_{product}")
        constraint.SetCoefficient(unmet[product], 1.0)
        for wc, _priority, _route_type, _penalty in eligible.get(product, []):
            if (product, wc) in x:
                constraint.SetCoefficient(x[(product, wc)], 1.0)

    all_wcs = {wc for _product, wc in x}
    for wc in all_wcs:
        limit = wc_limits.get(wc, 1.0)
        if limit <= 0:
            for product in products:
                if (product, wc) in x:
                    x[(product, wc)].SetUb(0.0)
            continue

        constraint = solver.Constraint(-inf, limit, f"cap_{wc}")
        for product in products:
            if (product, wc) in x:
                constraint.SetCoefficient(x[(product, wc)], 1.0 / eff_cap[(product, wc)])

    objective = solver.Objective()
    objective.SetMinimization()
    for product in products:
        objective.SetCoefficient(unmet[product], BIG_M)
    for (product, wc), variable in x.items():
        penalty = _get_penalty(product, wc, eligible)
        objective.SetCoefficient(variable, penalty / eff_cap[(product, wc)])

    status = solver.Solve()
    if status not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        print(f"  [WARN] Solver status {status} for month {month} {phase_label}")

    results: List[AllocationResult] = []
    residual: Dict[str, float] = {}
    total_unmet = 0.0

    for product in products:
        phase_demand_tons = demand[(month, product)]
        display_demand = full_demand.get((month, product), phase_demand_tons)
        family, plant = product_meta.get(product, ("", ""))
        unmet_tons = max(unmet[product].solution_value(), 0.0)
        residual[product] = unmet_tons
        total_unmet += unmet_tons

        for wc, priority, route_type, _penalty in eligible.get(product, []):
            if (product, wc) not in x:
                continue
            allocated_tons = max(x[(product, wc)].solution_value(), 0.0)
            rounded_allocated = round(allocated_tons, 4)
            if allocated_tons < EPSILON or rounded_allocated <= 0.0:
                continue
            cap = eff_cap[(product, wc)]
            results.append(AllocationResult(
                month=month,
                product=product,
                product_family=family,
                plant=plant,
                allocation_type="Internal",
                work_center=wc,
                route_type=route_type,
                priority=priority,
                demand_tons=round(display_demand, 4),
                allocated_tons=rounded_allocated,
                outsourced_tons=0.0,
                unmet_tons=0.0,
                capacity_share_pct=round(100.0 * allocated_tons / cap, 2),
            ))

    if verbose:
        label = f"[{phase_label}] " if phase_label else ""
        print(
            f"  {month} {label}{len(products)} products | "
            f"remaining = {total_unmet:,.1f} tons | obj = {objective.Value():,.2f}"
        )

    return results, residual


def _build_unmet_rows(
    month: str,
    residual: Dict[str, float],
    full_demand: DemandMap,
    product_meta: ProductMetaMap,
) -> List[AllocationResult]:
    rows: List[AllocationResult] = []
    for product, unmet_tons in residual.items():
        if unmet_tons <= EPSILON:
            continue
        family, plant = product_meta.get(product, ("", ""))
        rows.append(AllocationResult(
            month=month,
            product=product,
            product_family=family,
            plant=plant,
            allocation_type="Unmet",
            work_center="[UNALLOCATED]",
            route_type="N/A",
            priority=99,
            demand_tons=round(full_demand.get((month, product), unmet_tons), 4),
            allocated_tons=0.0,
            outsourced_tons=0.0,
            unmet_tons=round(unmet_tons, 4),
            capacity_share_pct=0.0,
        ))
    return rows


def _build_outsource_rows(
    month: str,
    residual: Dict[str, float],
    full_demand: DemandMap,
    product_meta: ProductMetaMap,
) -> List[AllocationResult]:
    rows: List[AllocationResult] = []
    for product, outsourced_tons in residual.items():
        if outsourced_tons <= EPSILON:
            continue
        family, plant = product_meta.get(product, ("", ""))
        rows.append(AllocationResult(
            month=month,
            product=product,
            product_family=family,
            plant=plant,
            allocation_type="Outsourced",
            work_center="[OUTSOURCED]",
            route_type="Toller",
            priority=99,
            demand_tons=round(full_demand.get((month, product), outsourced_tons), 4),
            allocated_tons=0.0,
            outsourced_tons=round(outsourced_tons, 4),
            unmet_tons=0.0,
            capacity_share_pct=0.0,
        ))
    return rows


def _apply_final_balances(
    results: List[AllocationResult],
    final_unmet: Dict[Tuple[str, str], float],
    final_outsourced: Dict[Tuple[str, str], float] | None = None,
) -> None:
    final_outsourced = final_outsourced or {}
    for result in results:
        key = (result.month, result.product)
        result.unmet_tons = round(final_unmet.get(key, 0.0), 4)
        if result.allocation_type == "Outsourced":
            result.outsourced_tons = round(final_outsourced.get(key, result.outsourced_tons), 4)
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


def _is_primary_like(route_type: str) -> bool:
    return route_type.strip().lower() in {"primary", "capacity"}


def _get_penalty(
    product: str,
    wc: str,
    eligible: RouteMap,
) -> float:
    for route_wc, _priority, _route_type, penalty in eligible.get(product, []):
        if route_wc == wc:
            return penalty
    return float(PRIORITY_BASE_PENALTY ** 2)
