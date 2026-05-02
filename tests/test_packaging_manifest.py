from pathlib import Path

from build_support.packaging_manifest import get_target, iter_data_mappings


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_packaging_targets_cover_both_desktop_launchers():
    main_target = get_target("capacity_optimizer")
    modeb_target = get_target("modeb_product_analysis")

    assert main_target.app_name == "CapacityOptimizer"
    assert main_target.entry_script == "CapacityOptimizerLauncher.pyw"
    assert main_target.required_resource_subpaths == ("Data_Input", "docs")
    assert modeb_target.app_name == "ModeBProductAnalysis"
    assert modeb_target.entry_script == "ModeBProductAnalysisLauncher.pyw"
    assert modeb_target.required_resource_subpaths == ()


def test_packaging_targets_resolve_required_resource_mappings():
    main_mappings = iter_data_mappings(PROJECT_ROOT, target_id="capacity_optimizer")
    assert main_mappings
    for source_path, _target_dir in main_mappings:
        assert Path(source_path).exists()

    modeb_mappings = iter_data_mappings(PROJECT_ROOT, target_id="modeb_product_analysis")
    assert modeb_mappings == []
