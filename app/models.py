"""
Data models for the Chemical Capacity Optimizer.
All capacity and demand are in metric TONS (not hours).
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Config:
    input_load_folder: str
    input_master_folder: str
    output_folder: str
    output_file_name: str
    scenario_name: str
    start_month: str        # YYYY-MM
    horizon_months: int = 60
    run_timestamp: Optional[str] = None
    notes: Optional[str] = None
    run_mode: str = "ModeB"
    direct_mode: bool = True
    verbose: bool = False
    skip_validation_errors: bool = False
    project_root_folder: str = ""
    license_status: str = ""
    license_id: Optional[str] = None
    license_type: Optional[str] = None
    licensed_to: Optional[str] = None
    license_expiry: Optional[str] = None
    license_binding_mode: Optional[str] = None
    license_machine_label: Optional[str] = None


@dataclass
class LoadRecord:
    """Forecast demand submitted by a planner, in tons per month."""
    month: str              # YYYY-MM
    planner_name: str
    product: str
    product_family: str
    plant: str
    forecast_tons: float
    scenario: Optional[str] = None
    resource_group_owner: Optional[str] = None
    scenario_version: Optional[str] = None
    comment: Optional[str] = None
    source_file: Optional[str] = None   # 来源文件名
    row_num: Optional[int] = None       # 原始文件中的行号（含表头=第1行，数据从第2行起）


@dataclass
class CapacityRecord:
    """
    Annual production capacity for a specific Product on a specific WorkCenter.
    Different products can have different throughput rates on the same line.
    """
    product: str
    work_center: str
    annual_capacity_tons: float
    utilization_target: float   # decimal, e.g. 0.88 for 88%
    effective_from: Optional[str] = None
    effective_to: Optional[str] = None

    @property
    def monthly_capacity_tons(self) -> float:
        return self.annual_capacity_tons / 12.0

    @property
    def effective_monthly_capacity_tons(self) -> float:
        return self.monthly_capacity_tons * self.utilization_target


@dataclass
class RoutingRecord:
    """
    Routing eligibility and priority for Product (or ProductFamily) → WorkCenter.
    Priority 1 = most preferred.  Higher number = lower priority = higher penalty.
    """
    work_center: str
    priority: int               # 1 = primary, 2 = first alt, 3 = toller, …
    eligible_flag: bool
    route_type: str             # Primary / Alternative / Toller
    product: Optional[str] = None
    product_family: Optional[str] = None
    penalty_weight: float = 1.0  # override; if 0 uses auto-derived from priority


@dataclass
class AllocationResult:
    """One row in the optimizer output table."""
    month: str
    product: str
    product_family: str
    plant: str
    allocation_type: str
    work_center: str
    route_type: str
    priority: int
    demand_tons: float
    allocated_tons: float
    outsourced_tons: float
    unmet_tons: float
    capacity_share_pct: float   # fraction of that WC's monthly capacity consumed
    planner_name: str = ""


@dataclass
class ValidationIssue:
    severity: str       # ERROR / WARNING
    check: str
    detail: str
