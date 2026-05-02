import os
import shutil
import unittest
import uuid
from contextlib import contextmanager
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from app.i18n import localize_column_name, localize_sheet_name, localize_value
from app.modeb_customer_case_report import (
    find_latest_modeb_report,
    generate_modeb_customer_case_report,
    load_modeb_report_context,
    resolve_modeb_report_selection,
)


TEST_TMP_ROOT = os.path.join(os.path.dirname(__file__), "_tmp")
os.makedirs(TEST_TMP_ROOT, exist_ok=True)


@contextmanager
def workspace_tempdir():
    tmpdir = os.path.join(TEST_TMP_ROOT, f"tmp_{uuid.uuid4().hex}")
    os.mkdir(tmpdir)
    try:
        yield tmpdir
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _write_modeb_report(path: Path, input_dir: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    detail_ws = workbook.create_sheet(localize_sheet_name("zh", "Allocation_Detail"))
    detail_headers = [
        "Capacity_Basis",
        localize_column_name("zh", "Month"),
        localize_column_name("zh", "PlannerName"),
        localize_column_name("zh", "Product"),
        localize_column_name("zh", "ProductFamily"),
        localize_column_name("zh", "Plant"),
        localize_column_name("zh", "AllocationType"),
        localize_column_name("zh", "WorkCenter"),
        localize_column_name("zh", "Demand_Tons"),
        localize_column_name("zh", "Allocated_Tons"),
        localize_column_name("zh", "Outsourced_Tons"),
        localize_column_name("zh", "Unmet_Tons"),
        localize_column_name("zh", "CapacityShare_Pct"),
        "Allocation_Source",
        "Residual_After_Capacity_Tons",
        "Residual_After_Routing_Tons",
        localize_column_name("zh", "RouteType"),
        localize_column_name("zh", "Priority"),
        localize_column_name("zh", "Year"),
    ]
    for col_index, header in enumerate(detail_headers, start=1):
        detail_ws.cell(3, col_index).value = header

    detail_rows = [
        ["Max", "2027-01", "PlannerA", "P1", "F1", "PLT1", localize_value("zh", "Internal"), "WC1", 100.0, 60.0, 0.0, 0.0, 60.0, "Capacity_Base", 40.0, 40.0, "Capacity", 1, 2027],
        ["Max", "2027-01", "PlannerA", "P1", "F1", "PLT1", localize_value("zh", "Outsourced"), "TOL1", 100.0, 0.0, 40.0, 0.0, 0.0, "Toller", 40.0, 0.0, "Toller", 1, 2027],
        ["Max", "2027-01", "PlannerA", "P2", "F2", "PLT2", localize_value("zh", "Internal"), "WC2", 90.0, 90.0, 0.0, 0.0, 100.0, "Capacity_Base", 0.0, 0.0, "Capacity", 1, 2027],
        ["Planner", "2027-01", "PlannerA", "P1", "F1", "PLT1", localize_value("zh", "Internal"), "WC1", 100.0, 60.0, 0.0, 0.0, 60.0, "Capacity_Base", 40.0, 40.0, "Capacity", 1, 2027],
        ["Planner", "2027-01", "PlannerA", "P1", "F1", "PLT1", localize_value("zh", "Internal"), "WC3", 100.0, 30.0, 0.0, 0.0, 30.0, "Routing_Reroute", 40.0, 10.0, "Primary", 1, 2027],
        ["Planner", "2027-01", "PlannerA", "P1", "F1", "PLT1", localize_value("zh", "Unmet"), "[UNALLOCATED]", 100.0, 0.0, 0.0, 10.0, 0.0, "Unmet", 40.0, 10.0, "N/A", 99, 2027],
        ["Planner", "2027-01", "PlannerA", "P2", "F2", "PLT2", localize_value("zh", "Internal"), "WC2", 90.0, 90.0, 0.0, 0.0, 100.0, "Capacity_Base", 0.0, 0.0, "Capacity", 1, 2027],
    ]
    for row_index, row in enumerate(detail_rows, start=4):
        for col_index, value in enumerate(row, start=1):
            detail_ws.cell(row_index, col_index).value = value

    run_info_ws = workbook.create_sheet(localize_sheet_name("zh", "Run_Info"))
    run_info_ws["A1"] = "Capacity_Basis"
    run_info_ws["B1"] = "参数"
    run_info_ws["C1"] = "值"
    run_info_rows = [
        ["Max", "Scenario_Name", "Baseline"],
        ["Max", "Input_Load_Folder", str(input_dir)],
        ["Max", "Input_Master_Folder", str(input_dir)],
        ["Max", "Output_Folder", str(input_dir)],
        ["Max", "Run_Timestamp", "2026-05-02 13:00:00"],
        ["Planner", "Scenario_Name", "Baseline"],
    ]
    for row_index, row in enumerate(run_info_rows, start=2):
        for col_index, value in enumerate(row, start=1):
            run_info_ws.cell(row_index, col_index).value = value

    workbook.save(path)
    workbook.close()


class ModeBCustomerCaseReportTests(unittest.TestCase):
    def test_find_latest_modeb_report_ignores_temp_files(self):
        with workspace_tempdir() as tmpdir:
            output_dir = Path(tmpdir)
            older = output_dir / "capacity_result_ModeB_Baseline_20260101_000000.xlsx"
            newer = output_dir / "capacity_result_ModeB_Baseline_20260102_000000.xlsx"
            temp_lock = output_dir / "~$capacity_result_ModeB_Baseline_20260103_000000.xlsx"
            modea = output_dir / "capacity_result_ModeA_Baseline_20260104_000000.xlsx"
            for path in (older, newer, temp_lock, modea):
                path.write_bytes(b"x")
            os.utime(older, (1, 1))
            os.utime(newer, (2, 2))
            os.utime(temp_lock, (3, 3))
            os.utime(modea, (4, 4))

            latest = find_latest_modeb_report(output_dir)
            selection = resolve_modeb_report_selection(
                output_dir=output_dir,
                manual_report_path=older.name,
                use_latest_report=False,
            )

        self.assertEqual(latest, newer)
        self.assertEqual(selection.selected_path, older.resolve())
        self.assertEqual(selection.latest_path, newer)
        self.assertFalse(selection.is_latest)

    def test_load_modeb_report_context_accepts_localized_workbook(self):
        with workspace_tempdir() as tmpdir:
            input_dir = Path(tmpdir) / "Data_Input"
            input_dir.mkdir()
            report_path = Path(tmpdir) / "capacity_result_ModeB_Baseline_20260502_130000.xlsx"
            _write_modeb_report(report_path, input_dir)

            context = load_modeb_report_context(report_path)

        self.assertEqual(context.scenario_name, "Baseline")
        self.assertEqual(context.available_bases, ("Max", "Planner"))
        self.assertEqual(len(context.detail_rows), 7)
        self.assertEqual(context.input_load_folder, input_dir.resolve())

    def test_generate_modeb_customer_case_report_creates_requested_product_sheets(self):
        with workspace_tempdir() as tmpdir:
            input_dir = Path(tmpdir) / "Data_Input"
            input_dir.mkdir()
            report_path = Path(tmpdir) / "capacity_result_ModeB_Baseline_20260502_130000.xlsx"
            _write_modeb_report(report_path, input_dir)

            pd.DataFrame(
                [
                    {"Month": "2027-01", "PlannerName": "PlannerA", "Product": "P1", "ProductFamily": "F1", "Plant": "PLT1", "Forecast_Tons": 100.0, "ScenarioVersion": "Baseline"},
                    {"Month": "2027-01", "PlannerName": "PlannerA", "Product": "P2", "ProductFamily": "F2", "Plant": "PLT2", "Forecast_Tons": 90.0, "ScenarioVersion": "Baseline"},
                ]
            ).to_csv(input_dir / "planner1_load.csv", index=False)
            pd.DataFrame(
                [
                    {"Product": "P1", "WorkCenter": "WC1", "Annual_Capacity_Tons": 720.0, "Utilization_Target": 1.0},
                    {"Product": "P2", "WorkCenter": "WC2", "Annual_Capacity_Tons": 1080.0, "Utilization_Target": 1.0},
                ]
            ).to_csv(input_dir / "master_capacity.csv", index=False)
            pd.DataFrame(
                [
                    {"Product": "P1", "Resource": "WC3", "Max Capacity Ton": 30.0, "Planner Capacity Ton": 30.0, "EligibleFalg": "Y", "Router Type": "Primary"},
                    {"Product": "P1", "Resource": "TOL1", "Max Capacity Ton": 40.0, "Planner Capacity Ton": 40.0, "EligibleFalg": "Y", "Router Type": "Toller"},
                ]
            ).to_csv(input_dir / "master_routing.csv", index=False)

            output_path = generate_modeb_customer_case_report(
                report_path=report_path,
                products=["P1", "P2"],
                output_dir=tmpdir,
                output_name="product_demo.xlsx",
            )

            workbook = load_workbook(output_path, read_only=True, data_only=True)

        self.assertTrue(output_path.name.startswith("product_demo_"))
        self.assertEqual(workbook.sheetnames, ["总览", "1_P1", "2_P2"])
        summary_ws = workbook["总览"]
        self.assertEqual(summary_ws["A1"].value, "ModeB 产品分析报告")
        self.assertEqual(summary_ws["A5"].value, "产品")
        self.assertEqual(summary_ws["B5"].value, "案例类型")
        product_ws = workbook["1_P1"]
        self.assertEqual(product_ws["A1"].value, "P1 - ModeB 产品分析")
        self.assertEqual(product_ws["A5"].value, "Capacity_Basis")
        self.assertEqual(product_ws["B5"].value, "案例类型")
        workbook.close()
