import os
import tempfile
import unittest

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


class RegressionTests(unittest.TestCase):
    def test_load_direct_accepts_lowercase_load_headers(self):
        with tempfile.TemporaryDirectory() as tmpdir:
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
        with tempfile.TemporaryDirectory() as tmpdir:
            project_root = os.path.join(tmpdir, "portable_root")
            template_path = os.path.join(project_root, "Tooling Control Panel", "Capacity_Optimizer_Control.xlsx")
            runner = CliRunner()
            result = runner.invoke(create_template_main, ["--out", template_path])

            self.assertEqual(result.exit_code, 0, msg=result.output)

            workbook = load_workbook(template_path)
            self.assertIn("Deployment_Steps", workbook.sheetnames)
            self.assertEqual(workbook.active.title, "Deployment_Steps")
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
        with tempfile.TemporaryDirectory() as tmpdir:
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
                "Allocation_Summary",
                "Outsource_Summary",
                "Unmet_Summary",
                "WC_Load_Pct",
                "Binary_Feasibility",
                "Validation_Issues",
                "Run_Info",
            }
            self.assertTrue(expected_sheets.issubset(set(workbook.sheetnames)))
            workbook.close()

    def test_load_direct_rejects_unrecognized_planner_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
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
        with tempfile.TemporaryDirectory() as tmpdir:
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
        with tempfile.TemporaryDirectory() as tmpdir:
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

            out_path = write_mode_comparison_summary(
                mode_results={"ModeA": mode_a_results, "ModeB": mode_b_results},
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
