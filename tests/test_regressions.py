import os
import tempfile
import unittest
from contextlib import contextmanager

import pandas as pd
from click.testing import CliRunner
from openpyxl import load_workbook

from data_loader import _aggregate_load_records, _parse_load_df, load_direct_mode_a
from create_template import main as create_template_main
from data_loader import load_config
from main import _validate_direct_mode_setup
from models import AllocationResult, Config, LoadRecord
from optimizer import _build_demand
from output_writer import write_mode_comparison_summary, write_results
from validator import ValidationIssue, format_issue_report, validate


TEST_TMP_ROOT = os.path.join(os.path.dirname(__file__), "_tmp")
os.makedirs(TEST_TMP_ROOT, exist_ok=True)


@contextmanager
def workspace_tempdir():
    with tempfile.TemporaryDirectory(dir=TEST_TMP_ROOT) as tmpdir:
        yield tmpdir


class RegressionTests(unittest.TestCase):
    def test_load_direct_accepts_lowercase_load_headers(self):
        with workspace_tempdir() as tmpdir:
            pd.DataFrame(
                [
                    {
                        "month": "2025-01",
                        "plannername": "P1",
                        "product": "0001",
                        "productfamily": "F1",
                        "plant": "A",
                        "forecast_tons": 10,
                    }
                ]
            ).to_csv(os.path.join(tmpdir, "planner1_load.csv"), index=False)
            pd.DataFrame(
                [
                    {
                        "Product": "1",
                        "WorkCenter": "WC1",
                        "Annual_Capacity_Tons": 120,
                        "Utilization_Target": 1,
                    }
                ]
            ).to_csv(os.path.join(tmpdir, "master_capacity.csv"), index=False)

            loads, capacities, routings = load_direct_mode_a(tmpdir, tmpdir)

        self.assertEqual(len(loads), 1)
        self.assertEqual(loads[0].product, "1")
        self.assertEqual(len(capacities), 1)
        self.assertEqual(routings, [])

    def test_parse_load_preserves_negative_tons_for_validation(self):
        loads = _parse_load_df(
            pd.DataFrame(
                [
                    {
                        "Month": "2025-01",
                        "PlannerName": "P1",
                        "Product": "0001",
                        "ProductFamily": "F1",
                        "Plant": "A",
                        "Forecast_Tons": -5,
                    }
                ]
            ),
            source_file="demo.csv",
        )

        issues = validate(loads, [], [], mode="ModeA")

        self.assertEqual(loads[0].forecast_tons, -5.0)
        self.assertIn(("ERROR", "LoadNegativeTons"), {(i.severity, i.check) for i in issues})

    def test_negative_duplicates_are_not_aggregated_away(self):
        records = [
            LoadRecord(
                month="2025-01",
                planner_name="P1",
                product="1",
                product_family="F1",
                plant="A",
                forecast_tons=10.0,
                scenario="Case90",
            ),
            LoadRecord(
                month="2025-01",
                planner_name="P1",
                product="1",
                product_family="F1",
                plant="A",
                forecast_tons=-5.0,
                scenario="Case90",
            ),
        ]

        aggregated = _aggregate_load_records(records)

        self.assertEqual(len(aggregated), 2)
        self.assertEqual(sorted(record.forecast_tons for record in aggregated), [-5.0, 10.0])

    def test_aggregate_and_demand_merge_multi_plant_metadata(self):
        aggregated = _aggregate_load_records(
            [
                LoadRecord(
                    month="2025-01",
                    planner_name="P1",
                    product="1",
                    product_family="F1",
                    plant="N029",
                    forecast_tons=3.0,
                    scenario="Case90",
                ),
                LoadRecord(
                    month="2025-01",
                    planner_name="P1",
                    product="1",
                    product_family="F1",
                    plant="C447",
                    forecast_tons=2.0,
                    scenario="Case90",
                ),
                LoadRecord(
                    month="2025-02",
                    planner_name="P1",
                    product="1",
                    product_family="F1",
                    plant="Z100",
                    forecast_tons=4.0,
                    scenario="Case90",
                ),
            ]
        )

        demand, product_meta = _build_demand(aggregated)

        self.assertEqual(len(aggregated), 2)
        self.assertEqual(aggregated[0].plant, "C447 | N029")
        self.assertEqual(demand[("2025-01", "1")], 5.0)
        self.assertEqual(product_meta["1"], ("F1", "C447 | N029 | Z100"))

    def test_format_issue_report_compacts_repeated_warnings(self):
        issues = [
            ValidationIssue(
                severity="WARNING",
                check="LoadZeroTons",
                detail=f"row {idx}",
            )
            for idx in range(1, 8)
        ]

        lines = format_issue_report(issues, warning_example_limit=3)
        report = "\n".join(lines)

        self.assertIn("[LoadZeroTons] 7 occurrence(s)", report)
        self.assertIn("e.g. row 1", report)
        self.assertIn("... and 4 more", report)

    def test_create_template_and_load_control_config(self):
        with workspace_tempdir() as tmpdir:
            project_root = os.path.join(tmpdir, "portable_root")
            template_path = os.path.join(project_root, "Tooling Control Panel", "Capacity_Optimizer_Control.xlsx")
            runner = CliRunner()
            result = runner.invoke(create_template_main, ["--out", template_path])

            self.assertEqual(result.exit_code, 0, msg=result.output)

            workbook = load_workbook(template_path)
            self.assertIn("Deployment_Steps", workbook.sheetnames)
            self.assertEqual(workbook.active.title, "Deployment_Steps")
            deployment_ws = workbook["Deployment_Steps"]
            deployment_text = "\n".join(
                str(cell.value)
                for row in deployment_ws.iter_rows()
                for cell in row
                if cell.value
            )
            self.assertIn("license.json", deployment_text)
            self.assertIn("get_machine_fingerprint.bat", deployment_text)
            self.assertIn("unbound", deployment_text.lower())
            instructions_text = "\n".join(
                str(cell.value)
                for row in workbook["Instructions"].iter_rows()
                for cell in row
                if cell.value
            )
            self.assertIn("license.json", instructions_text)
            self.assertIn("unbound", instructions_text.lower())
            ws = workbook["Control_Panel"]
            row_by_parameter = {
                ws[f"A{row_num}"].value: row_num
                for row_num in range(2, 30)
                if ws[f"A{row_num}"].value
            }
            ws[f"B{row_by_parameter['Start_Year']}"] = 2027
            ws[f"B{row_by_parameter['Start_Month_Num']}"] = 4
            ws[f"B{row_by_parameter['Run_Mode']}"] = "Both"
            ws[f"B{row_by_parameter['Direct_Mode']}"] = "Yes"
            ws[f"B{row_by_parameter['Verbose']}"] = "Yes"
            ws[f"B{row_by_parameter['Skip_Validation_Errors']}"] = "No"
            workbook.save(template_path)
            workbook.close()

            config = load_config(template_path)

        self.assertEqual(config.start_month, "2027-04")
        self.assertEqual(config.run_mode, "Both")
        self.assertTrue(config.direct_mode)
        self.assertTrue(config.verbose)
        self.assertFalse(config.skip_validation_errors)
        self.assertEqual(config.project_root_folder, project_root)
        self.assertEqual(config.input_load_folder, os.path.join(project_root, "Data_Input"))
        self.assertEqual(config.input_master_folder, os.path.join(project_root, "Data_Input"))
        self.assertEqual(config.output_folder, os.path.join(project_root, "output"))

    def test_write_results_creates_excel_report_sheets(self):
        with workspace_tempdir() as tmpdir:
            config = Config(
                input_load_folder=tmpdir,
                input_master_folder=tmpdir,
                output_folder=tmpdir,
                output_file_name="demo.xlsx",
                scenario_name="Baseline",
                start_month="2025-01",
                horizon_months=2,
                run_mode="Both",
                direct_mode=True,
            )
            loads = [
                LoadRecord(
                    month="2025-01",
                    planner_name="PlannerA",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    forecast_tons=100.0,
                )
            ]
            results = [
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Internal",
                    work_center="WC1",
                    route_type="Primary",
                    priority=1,
                    demand_tons=100.0,
                    allocated_tons=60.0,
                    outsourced_tons=0.0,
                    unmet_tons=40.0,
                    capacity_share_pct=60.0,
                ),
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Outsourced",
                    work_center="[UNALLOCATED]",
                    route_type="Toller",
                    priority=3,
                    demand_tons=100.0,
                    allocated_tons=0.0,
                    outsourced_tons=20.0,
                    unmet_tons=20.0,
                    capacity_share_pct=0.0,
                ),
            ]
            issues = [ValidationIssue(severity="WARNING", check="DemoWarning", detail="Example")]

            out_path = write_results(
                results=results,
                loads=loads,
                config=config,
                issues=issues,
                months=["2025-01", "2025-02"],
                mode="ModeB",
                toller_products={"P1"},
                metrics_by_mode={
                    "ModeA": {
                        "total_internal_allocated": 50.0,
                        "total_outsourced": 0.0,
                        "total_unmet": 50.0,
                        "service_level": 50.0,
                        "scenario_name": "Baseline",
                    },
                    "ModeB": {
                        "total_internal_allocated": 60.0,
                        "total_outsourced": 20.0,
                        "total_unmet": 20.0,
                        "service_level": 80.0,
                        "scenario_name": "Baseline",
                    },
                },
            )

            workbook = load_workbook(out_path)

            expected_sheets = {
                "Dashboard",
                "Monthly_Trend",
                "Bottleneck",
                "WC_Heatmap",
                "Product_Risk",
                "Allocation_Detail",
                "Planner_Result_Summary",
                "Planner_Product_Month",
                "Allocation_Summary",
                "Outsource_Summary",
                "Unmet_Summary",
                "WC_Load_Pct",
                "Binary_Feasibility",
                "Validation_Issues",
                "Run_Info",
            }
            self.assertTrue(expected_sheets.issubset(set(workbook.sheetnames)))
            allocation_summary_ws = workbook["Allocation_Summary"]
            allocation_summary_headers = [
                allocation_summary_ws.cell(1, idx).value
                for idx in range(1, allocation_summary_ws.max_column + 1)
            ]
            self.assertIn("PlannerName", allocation_summary_headers)
            outsource_summary_ws = workbook["Outsource_Summary"]
            outsource_summary_headers = [
                outsource_summary_ws.cell(1, idx).value
                for idx in range(1, outsource_summary_ws.max_column + 1)
            ]
            self.assertIn("PlannerName", outsource_summary_headers)
            unmet_summary_ws = workbook["Unmet_Summary"]
            unmet_summary_headers = [
                unmet_summary_ws.cell(1, idx).value
                for idx in range(1, unmet_summary_ws.max_column + 1)
            ]
            self.assertIn("PlannerName", unmet_summary_headers)
            workbook.close()

    def test_write_results_preserves_totals_after_planner_traceability_split(self):
        with workspace_tempdir() as tmpdir:
            config = Config(
                input_load_folder=tmpdir,
                input_master_folder=tmpdir,
                output_folder=tmpdir,
                output_file_name="demo.xlsx",
                scenario_name="Baseline",
                start_month="2025-01",
                horizon_months=2,
                run_mode="ModeA",
                direct_mode=True,
            )
            loads = [
                LoadRecord(
                    month="2025-01",
                    planner_name="PlannerA",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    forecast_tons=60.0,
                ),
                LoadRecord(
                    month="2025-01",
                    planner_name="PlannerB",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    forecast_tons=40.0,
                ),
            ]
            results = [
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Internal",
                    work_center="WC1",
                    route_type="Primary",
                    priority=1,
                    demand_tons=100.0,
                    allocated_tons=60.0,
                    outsourced_tons=0.0,
                    unmet_tons=40.0,
                    capacity_share_pct=60.0,
                ),
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Unmet",
                    work_center="[UNALLOCATED]",
                    route_type="N/A",
                    priority=99,
                    demand_tons=100.0,
                    allocated_tons=0.0,
                    outsourced_tons=0.0,
                    unmet_tons=40.0,
                    capacity_share_pct=0.0,
                ),
            ]

            out_path = write_results(
                results=results,
                loads=loads,
                config=config,
                issues=[],
                months=["2025-01", "2025-02"],
                mode="ModeA",
                toller_products=set(),
                metrics_by_mode={
                    "ModeA": {
                        "total_internal_allocated": 60.0,
                        "total_outsourced": 0.0,
                        "total_unmet": 40.0,
                        "service_level": 60.0,
                        "scenario_name": "Baseline",
                    },
                },
            )

            workbook = load_workbook(out_path, data_only=True)
            detail_ws = workbook["Allocation_Detail"]
            detail_headers = [detail_ws.cell(1, idx).value for idx in range(1, detail_ws.max_column + 1)]
            self.assertIn("PlannerName", detail_headers)

            planner_ws = workbook["Planner_Result_Summary"]
            planner_rows = list(planner_ws.iter_rows(min_row=3, values_only=True))
            header = planner_rows[0]
            planner_idx = header.index("PlannerName")
            demand_idx = header.index("Demand_Tons")
            internal_idx = header.index("Internal_Tons")
            unmet_idx = header.index("Unmet_Tons")
            data_rows = [
                row
                for row in planner_rows[1:]
                if row[planner_idx] not in (None, "") and isinstance(row[demand_idx], (int, float))
            ]

            summary = {
                row[planner_idx]: {
                    "demand": row[demand_idx],
                    "internal": row[internal_idx],
                    "unmet": row[unmet_idx],
                }
                for row in data_rows
            }
            workbook.close()

        self.assertEqual(set(summary), {"PlannerA", "PlannerB"})
        self.assertAlmostEqual(summary["PlannerA"]["demand"], 60.0, places=4)
        self.assertAlmostEqual(summary["PlannerA"]["internal"], 36.0, places=4)
        self.assertAlmostEqual(summary["PlannerA"]["unmet"], 24.0, places=4)
        self.assertAlmostEqual(summary["PlannerB"]["demand"], 40.0, places=4)
        self.assertAlmostEqual(summary["PlannerB"]["internal"], 24.0, places=4)
        self.assertAlmostEqual(summary["PlannerB"]["unmet"], 16.0, places=4)

    def test_load_direct_rejects_unrecognized_planner_files(self):
        with workspace_tempdir() as tmpdir:
            pd.DataFrame(
                [
                    {
                        "Month": "2025-01",
                        "PlannerName": "P1",
                        "Product": "0001",
                        "ProductFamily": "F1",
                        "Plant": "A",
                        "Forecast_Tons": 10,
                    }
                ]
            ).to_csv(os.path.join(tmpdir, "wrong_name.csv"), index=False)
            pd.DataFrame(
                [
                    {
                        "Product": "1",
                        "WorkCenter": "WC1",
                        "Annual_Capacity_Tons": 120,
                        "Utilization_Target": 1,
                    }
                ]
            ).to_csv(os.path.join(tmpdir, "master_capacity.csv"), index=False)

            with self.assertRaises(FileNotFoundError):
                load_direct_mode_a(tmpdir, tmpdir)

    def test_validate_direct_mode_setup_rejects_invalid_project_root(self):
        with workspace_tempdir() as tmpdir:
            os.makedirs(os.path.join(tmpdir, "Data_Input"))
            config = Config(
                project_root_folder=tmpdir,
                input_load_folder=os.path.join(tmpdir, "Data_Input"),
                input_master_folder=os.path.join(tmpdir, "Data_Input"),
                output_folder=os.path.join(tmpdir, "output"),
                output_file_name="demo.xlsx",
                scenario_name="Baseline",
                start_month="2025-01",
                horizon_months=2,
                run_mode="Both",
                direct_mode=True,
            )

            with self.assertRaises(FileNotFoundError):
                _validate_direct_mode_setup(config, ["ModeA", "ModeB"])

    def test_write_mode_comparison_summary_creates_standalone_workbook(self):
        with workspace_tempdir() as tmpdir:
            config = Config(
                input_load_folder=tmpdir,
                input_master_folder=tmpdir,
                output_folder=tmpdir,
                output_file_name="demo.xlsx",
                scenario_name="Baseline",
                start_month="2025-01",
                horizon_months=2,
                run_mode="Both",
                direct_mode=True,
            )
            mode_a_results = [
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Internal",
                    work_center="WC1",
                    route_type="Primary",
                    priority=1,
                    demand_tons=100.0,
                    allocated_tons=50.0,
                    outsourced_tons=0.0,
                    unmet_tons=50.0,
                    capacity_share_pct=50.0,
                ),
            ]
            mode_b_results = [
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Internal",
                    work_center="WC1",
                    route_type="Primary",
                    priority=1,
                    demand_tons=100.0,
                    allocated_tons=60.0,
                    outsourced_tons=0.0,
                    unmet_tons=40.0,
                    capacity_share_pct=60.0,
                ),
                AllocationResult(
                    month="2025-01",
                    product="P1",
                    product_family="F1",
                    plant="PLT1",
                    allocation_type="Outsourced",
                    work_center="[UNALLOCATED]",
                    route_type="Toller",
                    priority=3,
                    demand_tons=100.0,
                    allocated_tons=0.0,
                    outsourced_tons=20.0,
                    unmet_tons=20.0,
                    capacity_share_pct=0.0,
                ),
            ]
            mode_loads = {
                "ModeA": [
                    LoadRecord(
                        month="2025-01",
                        planner_name="PlannerA",
                        product="P1",
                        product_family="F1",
                        plant="PLT1",
                        forecast_tons=100.0,
                    )
                ],
                "ModeB": [
                    LoadRecord(
                        month="2025-01",
                        planner_name="PlannerA",
                        product="P1",
                        product_family="F1",
                        plant="PLT1",
                        forecast_tons=100.0,
                    )
                ],
            }

            out_path = write_mode_comparison_summary(
                mode_results={"ModeA": mode_a_results, "ModeB": mode_b_results},
                mode_loads=mode_loads,
                config=config,
                months=["2025-01", "2025-02"],
                metrics_by_mode={
                    "ModeA": {
                        "total_demand": 100.0,
                        "total_internal_allocated": 50.0,
                        "total_outsourced": 0.0,
                        "total_unmet": 50.0,
                        "service_level": 50.0,
                        "result_rows": 1,
                        "months": 2,
                        "scenario_name": "Baseline",
                    },
                    "ModeB": {
                        "total_demand": 100.0,
                        "total_internal_allocated": 60.0,
                        "total_outsourced": 20.0,
                        "total_unmet": 20.0,
                        "service_level": 80.0,
                        "result_rows": 2,
                        "months": 2,
                        "scenario_name": "Baseline",
                    },
                },
            )

            workbook = load_workbook(out_path)
            expected_sheets = {
                "Executive_Comparison",
                "Monthly_Trend_Compare",
                "Bottleneck_Compare",
                "WC_Heatmap_Compare",
                "Product_Risk_Compare",
                "Planner_Compare",
                "Run_Info",
            }
            self.assertRegex(
                os.path.basename(out_path),
                r"^Summary of Mode A and Mode B_\d{8}_\d{6}\.xlsx$",
            )
            self.assertTrue(expected_sheets.issubset(set(workbook.sheetnames)))
            workbook.close()


if __name__ == "__main__":
    unittest.main()
