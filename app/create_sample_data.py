"""
Generate synthetic demo data for the Capacity Optimizer.

Outputs CSV files under Data_Input/:
  - planner1_load.csv
  - planner2_load.csv
  - planner3_load.csv
  - planner4_load.csv
  - master_capacity.csv
  - master_routing.csv

The dataset is intentionally fictional and unrelated to any real customer data.
"""
from __future__ import annotations

import argparse
import math
import random
import shutil
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path

import pandas as pd

from app.runtime_paths import resolve_runtime_paths

RUNTIME_PATHS = resolve_runtime_paths()
BASE_DIR = RUNTIME_PATHS.app_install_dir
OUT_DIR = RUNTIME_PATHS.sample_data_dir
MONTH_COUNT = 72
START_YEAR = 2027
START_MONTH = 1
UTILIZATION_TARGET = 1.0
BASELINE_CAPACITY_FACTOR = 4.95
ROUTING_PRIMARY_MAX_FACTOR = 5.05
ROUTING_PRIMARY_PLANNER_FACTOR = 4.82
ROUTING_ALTERNATIVE_MAX_FACTOR = 2.05
ROUTING_ALTERNATIVE_PLANNER_FACTOR = 1.74
ROUTING_TOLLER_MAX_FACTOR = 1.58
ROUTING_TOLLER_PLANNER_FACTOR = 1.34

SCENARIOS = ("Baseline", "Expansion", "Lean")

STANDARD_VARIANT = "standard"
BOTTLENECK_VARIANT = "bottleneck"


@dataclass(frozen=True)
class FamilyConfig:
    planner_file: str
    planner_name: str
    family: str
    family_code: str
    plants: tuple[str, str]
    primary_resource: str
    primary_alias: str
    alternative_resource: str | None
    alternative_alias: str | None
    toller_resource: str | None
    toller_alias: str | None
    base_min: float
    base_max: float
    season_amp: float
    trend_min: float
    trend_max: float


@dataclass(frozen=True)
class VariantConfig:
    name: str
    output_dir: Path
    family_demand_scale: dict[str, float]
    resource_capacity_bias: dict[str, float]
    resource_demand_profiles: dict[str, tuple[float, float, float]] | None = None
    alternative_route_scale: float = 1.0
    altless_families: frozenset[str] = frozenset()


@dataclass(frozen=True)
class ProductOverride:
    demand_multiplier: float | None = None
    baseline_capacity_factor: float | None = None
    include_routing: bool | None = None
    include_alternative: bool | None = None
    include_toller: bool | None = None
    primary_resource: str | None = None
    primary_alias: str | None = None
    alternative_resource: str | None = None
    alternative_alias: str | None = None
    expansion_factor: float | None = None
    lean_factor: float | None = None
    routing_primary_max_factor: float | None = None
    routing_primary_planner_factor: float | None = None
    routing_alternative_max_factor: float | None = None
    routing_alternative_planner_factor: float | None = None
    routing_toller_max_factor: float | None = None
    routing_toller_planner_factor: float | None = None


FAMILY_CONFIGS: list[FamilyConfig] = [
    FamilyConfig("planner1_load.csv", "NorthDesk", "AURORA FLEX", "AUF", ("PLT-A1", "PLT-A2"), "North Reactor A-1200L", "NRA", "Central Mix E-1500L", "CME", "Partner Cell J-External", "PCJ", 58, 96, 0.08, -0.02, 0.10),
    FamilyConfig("planner1_load.csv", "NorthDesk", "AURORA PRIME", "AUP", ("PLT-A1", "PLT-A3"), "North Reactor B-1200L", "NRB", "Central Mix E-1500L", "CME", "Partner Cell J-External", "PCJ", 52, 88, 0.09, 0.00, 0.14),
    FamilyConfig("planner1_load.csv", "NorthDesk", "NOVA BOND", "NVB", ("PLT-B1", "PLT-B2"), "East Blend C-900L", "EBC", "Central Mix F-1500L", "CMF", None, None, 42, 84, 0.07, -0.03, 0.08),
    FamilyConfig("planner1_load.csv", "NorthDesk", "NOVA SEAL", "NVS", ("PLT-B1", "PLT-B3"), "East Blend D-900L", "EBD", "Central Mix F-1500L", "CMF", None, None, 38, 74, 0.10, 0.02, 0.16),

    FamilyConfig("planner2_load.csv", "HarborDesk", "TERRA CORE", "TRC", ("PLT-C1", "PLT-C2"), "South Line G-700L", "SLG", "East Blend C-900L", "EBC", "Partner Cell J-External", "PCJ", 46, 90, 0.08, -0.01, 0.13),
    FamilyConfig("planner2_load.csv", "HarborDesk", "TERRA SHIELD", "TRS", ("PLT-C1", "PLT-C3"), "South Line H-700L", "SLH", "East Blend D-900L", "EBD", None, None, 40, 82, 0.09, 0.01, 0.12),
    FamilyConfig("planner2_load.csv", "HarborDesk", "PULSE GEL", "PLG", ("PLT-D1", "PLT-D2"), "Central Mix E-1500L", "CME", "North Reactor A-1200L", "NRA", None, None, 34, 68, 0.11, -0.04, 0.07),
    FamilyConfig("planner2_load.csv", "HarborDesk", "PULSE FOAM", "PLF", ("PLT-D1", "PLT-D3"), "Central Mix F-1500L", "CMF", "North Reactor B-1200L", "NRB", None, None, 32, 72, 0.10, 0.03, 0.18),

    FamilyConfig("planner3_load.csv", "CanyonDesk", "LUMEN COAT", "LMC", ("PLT-E1", "PLT-E2"), "East Blend C-900L", "EBC", "South Line G-700L", "SLG", None, None, 36, 78, 0.07, -0.01, 0.10),
    FamilyConfig("planner3_load.csv", "CanyonDesk", "LUMEN BASE", "LMB", ("PLT-E1", "PLT-E3"), "East Blend D-900L", "EBD", "South Line H-700L", "SLH", None, None, 30, 64, 0.08, 0.00, 0.09),
    FamilyConfig("planner3_load.csv", "CanyonDesk", "ORBIT MIX", "ORM", ("PLT-F1", "PLT-F2"), "North Reactor A-1200L", "NRA", "Central Mix E-1500L", "CME", "Partner Cell J-External", "PCJ", 48, 92, 0.09, 0.02, 0.15),
    FamilyConfig("planner3_load.csv", "CanyonDesk", "ORBIT SEAL", "ORS", ("PLT-F1", "PLT-F3"), "North Reactor B-1200L", "NRB", "Central Mix F-1500L", "CMF", None, None, 42, 86, 0.10, -0.02, 0.11),

    FamilyConfig("planner4_load.csv", "VertexDesk", "VECTOR BOND", "VCB", ("PLT-G1", "PLT-G2"), "Central Mix E-1500L", "CME", "South Line G-700L", "SLG", None, None, 44, 84, 0.08, 0.01, 0.12),
    FamilyConfig("planner4_load.csv", "VertexDesk", "VECTOR GUARD", "VCG", ("PLT-G1", "PLT-G3"), "Central Mix F-1500L", "CMF", "South Line H-700L", "SLH", None, None, 38, 76, 0.09, -0.03, 0.08),
    FamilyConfig("planner4_load.csv", "VertexDesk", "CASCADE FORM", "CSF", ("PLT-H1", "PLT-H2"), "East Blend C-900L", "EBC", "North Reactor A-1200L", "NRA", None, None, 40, 80, 0.10, 0.03, 0.17),
    FamilyConfig("planner4_load.csv", "VertexDesk", "CASCADE FILL", "CSL", ("PLT-H1", "PLT-H3"), "East Blend D-900L", "EBD", "North Reactor B-1200L", "NRB", None, None, 34, 70, 0.11, -0.01, 0.10),
]


VARIANT_CONFIGS: dict[str, VariantConfig] = {
    STANDARD_VARIANT: VariantConfig(
        name=STANDARD_VARIANT,
        output_dir=OUT_DIR,
        family_demand_scale={},
        resource_capacity_bias={
            "North Reactor A-1200L": 0.98,
            "North Reactor B-1200L": 0.96,
            "East Blend C-900L": 1.01,
            "East Blend D-900L": 0.97,
            "South Line G-700L": 1.04,
            "South Line H-700L": 1.10,
            "Central Mix E-1500L": 1.02,
            "Central Mix F-1500L": 1.05,
            "Partner Cell J-External": 1.00,
        },
    ),
    BOTTLENECK_VARIANT: VariantConfig(
        name=BOTTLENECK_VARIANT,
        output_dir=BASE_DIR / "Data_Input_Set2",
        family_demand_scale={
            "TERRA CORE": 0.72,
            "TERRA SHIELD": 0.68,
            "LUMEN COAT": 0.78,
            "LUMEN BASE": 0.72,
            "VECTOR BOND": 0.82,
            "VECTOR GUARD": 0.80,
            "CASCADE FORM": 0.78,
            "CASCADE FILL": 0.80,
            "NOVA BOND": 0.82,
            "AURORA FLEX": 1.26,
            "AURORA PRIME": 1.14,
            "NOVA SEAL": 1.24,
            "ORBIT MIX": 1.22,
            "ORBIT SEAL": 1.04,
            "PULSE GEL": 1.16,
            "PULSE FOAM": 1.16,
        },
        resource_capacity_bias={
            "Central Mix F-1500L": 0.72,
            "Central Mix E-1500L": 0.78,
            "North Reactor A-1200L": 0.70,
            "North Reactor B-1200L": 1.15,
            "East Blend C-900L": 1.45,
            "East Blend D-900L": 0.72,
            "South Line G-700L": 1.45,
            "South Line H-700L": 2.20,
            "Partner Cell J-External": 1.35,
        },
        resource_demand_profiles={
            "East Blend D-900L": (0.22, 0.06, 1.20),
            "North Reactor A-1200L": (0.36, 0.08, 0.10),
            "Central Mix F-1500L": (0.40, 0.10, 2.10),
            "Central Mix E-1500L": (0.42, 0.10, 2.80),
            "East Blend C-900L": (0.55, 0.12, 0.80),
            "North Reactor B-1200L": (0.60, 0.10, 1.70),
            "South Line G-700L": (1.00, 0.14, 3.30),
            "South Line H-700L": (1.00, 0.12, 4.00),
        },
        alternative_route_scale=0.90,
        altless_families=frozenset({
            "LUMEN COAT",
            "LUMEN BASE",
            "VECTOR BOND",
            "VECTOR GUARD",
        }),
    ),
}


PRODUCT_OVERRIDES: dict[str, ProductOverride] = {
    # ModeA should clear this product even in Expansion.
    "TRS-02": ProductOverride(
        demand_multiplier=0.88,
        baseline_capacity_factor=7.10,
        include_routing=False,
        expansion_factor=1.08,
        lean_factor=0.83,
    ),
    # ModeA/ModeB should keep part internal and leave the rest as Unmet, with no routing help.
    "AUP-04": ProductOverride(
        demand_multiplier=1.12,
        baseline_capacity_factor=0.56,
        include_routing=False,
        primary_resource="North Reactor B-1200L Annex",
        primary_alias="NRB-X",
        include_alternative=False,
        include_toller=False,
        expansion_factor=1.16,
        lean_factor=0.82,
    ),
    # ModeA should absorb part of the demand internally; ModeB should send only the residual to Toller.
    "AUF-01": ProductOverride(
        demand_multiplier=1.06,
        baseline_capacity_factor=0.68,
        include_routing=True,
        primary_resource="North Reactor A-1200L Batch Cell",
        primary_alias="NRA-B",
        include_alternative=False,
        include_toller=True,
        expansion_factor=1.15,
        lean_factor=0.84,
        routing_primary_max_factor=0.70,
        routing_primary_planner_factor=0.68,
        routing_toller_max_factor=1.10,
        routing_toller_planner_factor=1.05,
    ),
    # ModeA overflows; ModeB cannot absorb enough internally and sends the residual to Toller.
    "AUF-05": ProductOverride(
        demand_multiplier=1.22,
        baseline_capacity_factor=4.12,
        include_routing=True,
        include_alternative=False,
        include_toller=True,
        expansion_factor=1.21,
        lean_factor=0.84,
        routing_primary_max_factor=4.20,
        routing_primary_planner_factor=4.12,
        routing_toller_max_factor=2.15,
        routing_toller_planner_factor=1.92,
    ),
    # ModeA overflows; this product is intentionally absent from routing and remains Unmet in ModeB.
    "NVS-04": ProductOverride(
        demand_multiplier=1.24,
        baseline_capacity_factor=4.18,
        include_routing=False,
        expansion_factor=1.22,
        lean_factor=0.81,
    ),
    # ModeA should cover about 60% internally; ModeB should cover the remaining share through routing reroute.
    "LMB-03": ProductOverride(
        demand_multiplier=1.16,
        baseline_capacity_factor=0.71,
        include_routing=True,
        include_alternative=True,
        include_toller=False,
        primary_resource="East Blend D-900L Pilot Cell",
        primary_alias="EBD-P",
        alternative_resource="West Flex Reactor Q-1000L",
        alternative_alias="WFR",
        expansion_factor=1.18,
        lean_factor=0.82,
        routing_primary_max_factor=0.73,
        routing_primary_planner_factor=0.71,
        routing_alternative_max_factor=0.64,
        routing_alternative_planner_factor=0.59,
    ),
    # ModeA overflows; ModeB reroutes part of the residual but still leaves Unmet with no toller.
    "VCB-03": ProductOverride(
        demand_multiplier=1.18,
        baseline_capacity_factor=4.16,
        include_routing=True,
        include_alternative=True,
        include_toller=False,
        expansion_factor=1.20,
        lean_factor=0.80,
        routing_primary_max_factor=4.24,
        routing_primary_planner_factor=4.16,
        routing_alternative_max_factor=1.18,
        routing_alternative_planner_factor=1.02,
    ),
}


def _month_starts(year: int, month: int, count: int) -> list[date]:
    values: list[date] = []
    y = year
    m = month
    for _ in range(count):
        values.append(date(y, m, 1))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return values


def _excel_serial(value: date) -> int:
    excel_epoch = date(1899, 12, 30)
    return (value - excel_epoch).days


def _variant_family_configs(variant: VariantConfig) -> list[FamilyConfig]:
    configs: list[FamilyConfig] = []
    for family in FAMILY_CONFIGS:
        if family.family in variant.altless_families:
            configs.append(replace(family, alternative_resource=None, alternative_alias=None))
        else:
            configs.append(family)
    return configs


def _product_specs(
    family_configs: list[FamilyConfig],
    family_demand_scale: dict[str, float] | None = None,
) -> list[dict]:
    specs: list[dict] = []
    family_demand_scale = family_demand_scale or {}
    for family_index, family in enumerate(family_configs):
        rng = random.Random(1_000 + family_index)
        for product_index in range(1, 7):
            product_code = f"{family.family_code}-{product_index:02d}"
            profile_rng = random.Random(9_000 + family_index * 20 + product_index)
            primary_resource = family.primary_resource
            primary_alias = family.primary_alias
            alternative_resource = family.alternative_resource
            alternative_alias = family.alternative_alias
            toller_resource = family.toller_resource
            toller_alias = family.toller_alias
            base_tons = round(rng.uniform(family.base_min, family.base_max), 2)
            base_tons *= family_demand_scale.get(family.family, 1.0)
            base_tons *= profile_rng.uniform(0.92, 1.12)

            include_routing = profile_rng.random() > 0.22
            include_alternative = bool(alternative_resource) and include_routing and profile_rng.random() > 0.18
            include_toller = bool(toller_resource) and include_routing and profile_rng.random() > 0.42

            baseline_capacity_factor = BASELINE_CAPACITY_FACTOR * profile_rng.uniform(0.86, 1.18)
            routing_primary_max_factor = baseline_capacity_factor * profile_rng.uniform(0.99, 1.08)
            routing_primary_planner_factor = max(
                baseline_capacity_factor,
                routing_primary_max_factor * profile_rng.uniform(0.91, 0.98),
            )

            routing_alternative_max_factor = 0.0
            routing_alternative_planner_factor = 0.0
            if include_alternative:
                routing_alternative_max_factor = (
                    ROUTING_ALTERNATIVE_MAX_FACTOR
                    * profile_rng.uniform(0.72, 1.38)
                )
                routing_alternative_planner_factor = (
                    routing_alternative_max_factor
                    * profile_rng.uniform(0.78, 0.93)
                )

            routing_toller_max_factor = 0.0
            routing_toller_planner_factor = 0.0
            if include_toller:
                routing_toller_max_factor = (
                    ROUTING_TOLLER_MAX_FACTOR
                    * profile_rng.uniform(0.82, 1.26)
                )
                routing_toller_planner_factor = (
                    routing_toller_max_factor
                    * profile_rng.uniform(0.81, 0.93)
                )

            scenario_factors = {
                "Baseline": 1.0,
                "Expansion": round(profile_rng.uniform(1.08, 1.23), 4),
                "Lean": round(profile_rng.uniform(0.77, 0.90), 4),
            }

            event_factors: dict[int, float] = {}
            for _ in range(profile_rng.randint(1, 3)):
                month_index = profile_rng.randrange(MONTH_COUNT)
                event_factors[month_index] = round(
                    event_factors.get(month_index, 1.0) * profile_rng.uniform(1.08, 1.24),
                    4,
                )
            for _ in range(profile_rng.randint(0, 2)):
                month_index = profile_rng.randrange(MONTH_COUNT)
                event_factors[month_index] = round(
                    event_factors.get(month_index, 1.0) * profile_rng.uniform(0.84, 0.95),
                    4,
                )

            override = PRODUCT_OVERRIDES.get(product_code)
            if override:
                if override.demand_multiplier is not None:
                    base_tons *= override.demand_multiplier
                if override.baseline_capacity_factor is not None:
                    baseline_capacity_factor = override.baseline_capacity_factor
                if override.include_routing is not None:
                    include_routing = override.include_routing
                if override.primary_resource is not None:
                    primary_resource = override.primary_resource
                    primary_alias = override.primary_alias or primary_alias
                if override.alternative_resource is not None:
                    alternative_resource = override.alternative_resource
                    alternative_alias = override.alternative_alias or alternative_alias
                if override.include_alternative is not None:
                    include_alternative = override.include_alternative and bool(alternative_resource)
                if override.include_toller is not None:
                    include_toller = override.include_toller and bool(toller_resource)
                if override.expansion_factor is not None:
                    scenario_factors["Expansion"] = override.expansion_factor
                if override.lean_factor is not None:
                    scenario_factors["Lean"] = override.lean_factor
                if override.routing_primary_max_factor is not None:
                    routing_primary_max_factor = override.routing_primary_max_factor
                if override.routing_primary_planner_factor is not None:
                    routing_primary_planner_factor = override.routing_primary_planner_factor
                if override.routing_alternative_max_factor is not None:
                    routing_alternative_max_factor = override.routing_alternative_max_factor
                if override.routing_alternative_planner_factor is not None:
                    routing_alternative_planner_factor = override.routing_alternative_planner_factor
                if override.routing_toller_max_factor is not None:
                    routing_toller_max_factor = override.routing_toller_max_factor
                if override.routing_toller_planner_factor is not None:
                    routing_toller_planner_factor = override.routing_toller_planner_factor

            if not include_routing:
                include_alternative = False
                include_toller = False
            if not include_alternative:
                routing_alternative_max_factor = 0.0
                routing_alternative_planner_factor = 0.0
            if not include_toller:
                routing_toller_max_factor = 0.0
                routing_toller_planner_factor = 0.0

            specs.append({
                "planner_file": family.planner_file,
                "planner_name": family.planner_name,
                "product": product_code,
                "family": family.family,
                "plant": family.plants[(product_index - 1) % len(family.plants)],
                "primary_resource": primary_resource,
                "primary_alias": primary_alias,
                "alternative_resource": alternative_resource,
                "alternative_alias": alternative_alias,
                "toller_resource": toller_resource,
                "toller_alias": toller_alias,
                "base_tons": round(base_tons, 2),
                "season_amp": family.season_amp + rng.uniform(-0.015, 0.015),
                "phase": rng.uniform(0.0, math.tau),
                "trend": rng.uniform(family.trend_min, family.trend_max),
                "pulse_amp": profile_rng.uniform(0.02, 0.08),
                "pulse_phase": profile_rng.uniform(0.0, math.tau),
                "scenario_factors": scenario_factors,
                "event_factors": event_factors,
                "include_routing": include_routing,
                "include_alternative": include_alternative,
                "include_toller": include_toller,
                "baseline_capacity_factor": baseline_capacity_factor,
                "routing_primary_max_factor": routing_primary_max_factor,
                "routing_primary_planner_factor": routing_primary_planner_factor,
                "routing_alternative_max_factor": routing_alternative_max_factor,
                "routing_alternative_planner_factor": routing_alternative_planner_factor,
                "routing_toller_max_factor": routing_toller_max_factor,
                "routing_toller_planner_factor": routing_toller_planner_factor,
                "noise_seed": 5_000 + family_index * 20 + product_index,
            })
    return specs


def _build_load_rows(
    specs: list[dict],
    months: list[date],
    variant: VariantConfig,
) -> dict[str, list[dict]]:
    planner_rows: dict[str, list[dict]] = {
        "planner1_load.csv": [],
        "planner2_load.csv": [],
        "planner3_load.csv": [],
        "planner4_load.csv": [],
    }
    resource_demand_profiles = variant.resource_demand_profiles or {}

    for spec in specs:
        rng = random.Random(spec["noise_seed"])
        for month_index, month_value in enumerate(months):
            excel_month = _excel_serial(month_value)
            seasonality = 1.0 + spec["season_amp"] * math.sin((month_index / 12.0) * math.tau + spec["phase"])
            trend_factor = 1.0 + spec["trend"] * (month_index / max(len(months) - 1, 1))
            pulse_factor = 1.0 + spec["pulse_amp"] * math.sin((month_index / 6.0) * math.tau + spec["pulse_phase"])
            demand_profile = resource_demand_profiles.get(spec["primary_resource"])
            demand_factor = 1.0
            if demand_profile:
                base, amplitude, phase = demand_profile
                demand_factor = max(
                    0.10,
                    base + amplitude * math.sin((month_index / 12.0) * math.tau + phase),
                )

            month_event_factor = spec["event_factors"].get(month_index, 1.0)
            monthly_base = spec["base_tons"] * seasonality * trend_factor * pulse_factor * demand_factor * month_event_factor
            for scenario_name in SCENARIOS:
                scenario_factor = spec["scenario_factors"][scenario_name]
                noise = 1.0 + rng.uniform(-0.06, 0.06)
                tons = max(
                    6.0,
                    monthly_base * scenario_factor * noise,
                )
                planner_rows[spec["planner_file"]].append({
                    "Month": excel_month,
                    "PlannerName": spec["planner_name"],
                    "Product": spec["product"],
                    "ProductFamily": spec["family"],
                    "Plant": spec["plant"],
                    "Forecast_Tons": round(tons, 4),
                    "Resource": spec["primary_resource"],
                    "Scenario Version": scenario_name,
                    "Comment": spec["primary_alias"],
                })

    return planner_rows


def _build_capacity_rows(
    specs: list[dict],
    variant: VariantConfig,
) -> list[dict]:
    rows: list[dict] = []
    for spec in specs:
        effective_monthly = (
            spec["base_tons"]
            * spec["baseline_capacity_factor"]
            * variant.resource_capacity_bias.get(spec["primary_resource"], 1.0)
        )
        annual_capacity = round((effective_monthly * 12.0) / UTILIZATION_TARGET, 2)
        rows.append({
            "Product": spec["product"],
            "Product Family": spec["family"],
            "Resource": spec["primary_resource"],
            "Annual Capacity Tons": annual_capacity,
            "Utilization Target": UTILIZATION_TARGET,
        })
    return rows


def _build_routing_rows(specs: list[dict], variant: VariantConfig) -> list[dict]:
    rows: list[dict] = []
    for spec in specs:
        if not spec["include_routing"]:
            continue

        route_defs = [
            (
                "Primary",
                spec["primary_resource"],
                spec["routing_primary_max_factor"],
                spec["routing_primary_planner_factor"],
            ),
        ]
        if spec["include_alternative"] and spec["alternative_resource"]:
            route_defs.append((
                "Alternative",
                spec["alternative_resource"],
                spec["routing_alternative_max_factor"] * variant.alternative_route_scale,
                spec["routing_alternative_planner_factor"] * variant.alternative_route_scale,
            ))
        if spec["include_toller"] and spec["toller_resource"]:
            route_defs.append((
                "Toller",
                spec["toller_resource"],
                spec["routing_toller_max_factor"],
                spec["routing_toller_planner_factor"],
            ))

        for route_type, resource, max_factor, planner_factor in route_defs:
            max_capacity = round(
                spec["base_tons"]
                * max_factor
                * variant.resource_capacity_bias.get(resource, 1.0)
                * 12.0,
                2,
            )
            planner_capacity = round(
                spec["base_tons"]
                * planner_factor
                * variant.resource_capacity_bias.get(resource, 1.0)
                * 12.0,
                2,
            )
            rows.append({
                "Product": spec["product"],
                "Product Family": spec["family"],
                "Resource": resource,
                "Max Capacity Ton": max_capacity,
                "Planner Capacity Ton": planner_capacity,
                "EligibleFalg": "Y",
                "Router Type": route_type,
            })
    return rows


def _write_csv(path: Path, rows: list[dict], columns: list[str]) -> None:
    if path.exists():
        path.chmod(0o666)
    df = pd.DataFrame(rows, columns=columns)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"  Written: {path.name:<20} {len(df):>6} rows")


def _copy_input_guide(output_dir: Path) -> None:
    guide_name = "DATA_INPUT_GUIDE_CN.md"
    source = OUT_DIR / guide_name
    target = output_dir / guide_name
    if not source.exists() or source.resolve() == target.resolve():
        return
    shutil.copy2(source, target)
    print(f"  Copied : {guide_name}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate synthetic demo input data.")
    parser.add_argument(
        "--variant",
        choices=sorted(VARIANT_CONFIGS),
        default=STANDARD_VARIANT,
        help="Dataset profile to generate. 'standard' writes Data_Input, 'bottleneck' writes Data_Input_Set2.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Optional override for the output folder.",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    variant = VARIANT_CONFIGS[args.variant]
    output_dir = Path(args.output_dir).resolve() if args.output_dir else variant.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    months = _month_starts(START_YEAR, START_MONTH, MONTH_COUNT)
    family_configs = _variant_family_configs(variant)
    specs = _product_specs(family_configs, family_demand_scale=variant.family_demand_scale)
    planner_rows = _build_load_rows(specs, months, variant)
    capacity_rows = _build_capacity_rows(specs, variant)
    routing_rows = _build_routing_rows(specs, variant)

    for planner_name in sorted(planner_rows):
        _write_csv(
            output_dir / planner_name,
            planner_rows[planner_name],
            [
                "Month",
                "PlannerName",
                "Product",
                "ProductFamily",
                "Plant",
                "Forecast_Tons",
                "Resource",
                "Scenario Version",
                "Comment",
            ],
        )

    _write_csv(
        output_dir / "master_capacity.csv",
        capacity_rows,
        [
            "Product",
            "Product Family",
            "Resource",
            "Annual Capacity Tons",
            "Utilization Target",
        ],
    )
    _write_csv(
        output_dir / "master_routing.csv",
        routing_rows,
        [
            "Product",
            "Product Family",
            "Resource",
            "Max Capacity Ton",
            "Planner Capacity Ton",
            "EligibleFalg",
            "Router Type",
        ],
    )

    _copy_input_guide(output_dir)
    print(f"\nSynthetic demo data created in: {output_dir}")


if __name__ == "__main__":
    main()
