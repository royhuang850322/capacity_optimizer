from pathlib import Path

from build_support.packaging_manifest import get_target, iter_data_mappings


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_packaging_targets_cover_both_desktop_launchers():
    main_target = get_target("capacity_optimizer")
    modeb_target = get_target("modeb_product_analysis")

    assert main_target.app_name == "CapacityOptimizer"
    assert main_target.entry_script == "CapacityOptimizerLauncher.pyw"
    assert modeb_target.app_name == "ModeBProductAnalysis"
    assert modeb_target.entry_script == "ModeBProductAnalysisLauncher.pyw"


def test_packaging_targets_resolve_required_resource_mappings():
    for target_id in ("capacity_optimizer", "modeb_product_analysis"):
        mappings = iter_data_mappings(PROJECT_ROOT, target_id=target_id)
        assert mappings
        for source_path, _target_dir in mappings:
            assert Path(source_path).exists()
