from __future__ import annotations

from typing import Any

import pandas as pd


def normalize_language(value: str | None) -> str:
    text = str(value or "").strip().lower()
    if text in {"zh", "zh-cn", "cn", "chinese", "中文"}:
        return "zh"
    return "en"


UI_TEXTS: dict[str, dict[str, str]] = {
    "app_title": {"en": "Capacity Optimizer", "zh": "产能优化工具"},
    "launcher_title": {"en": "Capacity Optimizer Launcher", "zh": "产能优化工具启动器"},
    "header_card": {"en": "Header", "zh": "页眉"},
    "header_subtitle": {
        "en": "Enterprise launcher for workspace setup, optimizer runs, and output diagnostics.",
        "zh": "用于工作区初始化、优化运行和输出诊断的企业级启动器。",
    },
    "theme_label": {"en": "Theme", "zh": "主题"},
    "language_label": {"en": "Language", "zh": "语言"},
    "help_button": {"en": "Help", "zh": "帮助"},
    "nav_settings_card": {"en": "Navigation & Settings", "zh": "导航与设置"},
    "nav_home": {"en": "Home", "zh": "首页"},
    "nav_configuration": {"en": "Configuration", "zh": "配置"},
    "nav_license": {"en": "License & Diagnostics", "zh": "许可与诊断"},
    "planning_card": {"en": "Planning", "zh": "计划设置"},
    "runtime_card": {"en": "Runtime", "zh": "运行设置"},
    "actions_card": {"en": "Actions", "zh": "操作"},
    "status_card": {"en": "Run Status", "zh": "运行状态"},
    "config_actions_card": {"en": "Configuration Actions", "zh": "配置操作"},
    "workspace_card": {"en": "Workspace", "zh": "工作区"},
    "license_card": {"en": "License & Diagnostics", "zh": "许可与诊断"},
    "footer_card": {"en": "Footer", "zh": "页脚"},
    "workspace_label": {"en": "Workspace", "zh": "工作区"},
    "data_input_label": {"en": "Data_Input", "zh": "数据输入"},
    "output_label": {"en": "Output", "zh": "输出目录"},
    "scenario_name_label": {"en": "Scenario Name", "zh": "场景名称"},
    "output_file_name_label": {"en": "Output File Name", "zh": "输出文件名"},
    "start_year_label": {"en": "Start Year", "zh": "开始年份"},
    "start_month_label": {"en": "Start Month", "zh": "开始月份"},
    "horizon_months_label": {"en": "Horizon Months", "zh": "滚动月份"},
    "run_mode_label": {"en": "Run Mode", "zh": "运行模式"},
    "verbose_label": {"en": "Verbose", "zh": "详细日志"},
    "skip_validation_label": {"en": "Skip Validation Errors", "zh": "忽略校验错误"},
    "browse_button": {"en": "Browse...", "zh": "浏览..."},
    "run_button": {"en": "Run Optimization", "zh": "运行优化"},
    "open_output_button": {"en": "Open Output Folder", "zh": "打开输出目录"},
    "open_logs_button": {"en": "Open Log Folder", "zh": "打开日志目录"},
    "open_last_log_button": {"en": "Open Latest Log", "zh": "打开最新日志"},
    "open_workspace_button": {"en": "Open Workspace Folder", "zh": "打开工作区目录"},
    "initialize_button": {"en": "Initialize Workspace", "zh": "初始化工作区"},
    "save_settings_button": {"en": "Save Settings", "zh": "保存设置"},
    "generate_fingerprint_button": {"en": "Generate Machine Fingerprint", "zh": "生成机器指纹"},
    "open_requests_button": {"en": "Open License Requests", "zh": "打开许可申请目录"},
    "open_license_folder_button": {"en": "Open License Folder", "zh": "打开许可目录"},
    "open_docs_button": {"en": "Open Workspace Docs", "zh": "打开工作区说明"},
    "status_placeholder": {
        "en": "Runtime status and operation notes appear here.",
        "zh": "这里显示运行状态和操作记录。",
    },
    "status_ready": {
        "en": "Launcher ready. Save settings to confirm the workspace path, then initialize the workspace.",
        "zh": "启动器已就绪。请先保存设置确认工作区路径，再初始化工作区。",
    },
    "status_workspace_ready": {"en": "Workspace Ready", "zh": "工作区已就绪"},
    "status_running": {"en": "Running...", "zh": "运行中..."},
    "status_run_succeeded": {"en": "Run Succeeded", "zh": "运行成功"},
    "status_run_failed": {"en": "Run Failed", "zh": "运行失败"},
    "status_run_started": {"en": "Run started.", "zh": "已开始运行。"},
    "help_title": {"en": "Help", "zh": "帮助"},
    "help_message": {
        "en": "Use Save Settings to confirm the workspace path and run options.\nThen use Initialize Workspace to create the folders under that path.\n\nWorkspace docs folder:\n{docs_dir}",
        "zh": "请先使用“保存设置”确认工作区路径和运行参数。\n然后使用“初始化工作区”在该路径下创建所需文件夹。\n\n工作区说明目录：\n{docs_dir}",
    },
    "saved_title": {"en": "Saved", "zh": "已保存"},
    "saved_message": {"en": "Launcher settings saved:\n{path}", "zh": "启动器设置已保存：\n{path}"},
    "workspace_ready_title": {"en": "Workspace Ready", "zh": "工作区已就绪"},
    "workspace_ready_message": {"en": "Workspace initialization completed.", "zh": "工作区初始化完成。"},
    "machine_fingerprint_title": {"en": "Machine Fingerprint Generated", "zh": "机器指纹已生成"},
    "machine_fingerprint_message": {"en": "Request file created:\n{path}", "zh": "已生成申请文件：\n{path}"},
    "invalid_settings_title": {"en": "Invalid Settings", "zh": "设置无效"},
    "save_failed_title": {"en": "Save Failed", "zh": "保存失败"},
    "run_completed_title": {"en": "Run Completed", "zh": "运行完成"},
    "run_completed_message": {"en": "{message}\n\nLog file:\n{path}", "zh": "{message}\n\n日志文件：\n{path}"},
    "run_failed_title": {"en": "Run Failed", "zh": "运行失败"},
    "run_failed_message": {"en": "{message}\n\nLog file:\n{path}", "zh": "{message}\n\n日志文件：\n{path}"},
    "no_logs_title": {"en": "No Logs Found", "zh": "未找到日志"},
    "no_logs_message": {"en": "No log files found in:\n{path}", "zh": "在以下目录未找到日志文件：\n{path}"},
    "path_not_found_title": {"en": "Path Not Found", "zh": "路径不存在"},
    "path_not_found_message": {"en": "Path does not exist:\n{path}", "zh": "路径不存在：\n{path}"},
    "open_failed_title": {"en": "Open Failed", "zh": "打开失败"},
    "open_failed_message": {"en": "Could not open path:\n{path}", "zh": "无法打开路径：\n{path}"},
    "generate_fingerprint_failed_title": {"en": "Generate Fingerprint Failed", "zh": "生成机器指纹失败"},
    "language_english": {"en": "English", "zh": "英文"},
    "language_chinese": {"en": "Chinese", "zh": "中文"},
    "theme_system": {"en": "System", "zh": "跟随系统"},
    "theme_light": {"en": "Light", "zh": "浅色"},
    "theme_dark": {"en": "Dark", "zh": "深色"},
    "mode_modea": {"en": "ModeA", "zh": "模式A"},
    "mode_modeb": {"en": "ModeB", "zh": "模式B"},
    "mode_both": {"en": "Both", "zh": "同时运行"},
    "yes": {"en": "Yes", "zh": "是"},
    "no": {"en": "No", "zh": "否"},
    "version_label": {"en": "Version: {version}", "zh": "版本：{version}"},
    "license_label": {"en": "License: validated at runtime", "zh": "许可：运行时校验"},
    "workspace_footer": {"en": "Workspace: {path}", "zh": "工作区：{path}"},
}


REPORT_TEXTS: dict[str, dict[str, str]] = {
    "dashboard": {"en": "Dashboard", "zh": "仪表板"},
    "monthly_trend": {"en": "Monthly_Trend", "zh": "月度趋势"},
    "bottleneck": {"en": "Bottleneck", "zh": "瓶颈分析"},
    "wc_heatmap": {"en": "WC_Heatmap", "zh": "工作中心热力图"},
    "product_risk": {"en": "Product_Risk", "zh": "产品风险"},
    "planner_result_summary": {"en": "Planner_Result_Summary", "zh": "计划员结果汇总"},
    "allocation_detail": {"en": "Allocation_Detail", "zh": "分配明细"},
    "planner_product_month": {"en": "Planner_Product_Month", "zh": "计划员产品月份汇总"},
    "allocation_summary": {"en": "Allocation_Summary", "zh": "内部分配汇总"},
    "outsource_summary": {"en": "Outsource_Summary", "zh": "外协汇总"},
    "unmet_summary": {"en": "Unmet_Summary", "zh": "未满足汇总"},
    "binary_feasibility": {"en": "Binary_Feasibility", "zh": "可行性矩阵"},
    "executive_comparison": {"en": "Executive_Comparison", "zh": "综合对比总览"},
    "monthly_trend_compare": {"en": "Monthly_Trend_Compare", "zh": "月度趋势对比"},
    "bottleneck_compare": {"en": "Bottleneck_Compare", "zh": "瓶颈对比"},
    "wc_heatmap_compare": {"en": "WC_Heatmap_Compare", "zh": "工作中心热力图对比"},
    "product_risk_compare": {"en": "Product_Risk_Compare", "zh": "产品风险对比"},
    "planner_compare": {"en": "Planner_Compare", "zh": "计划员对比"},
    "modea_cap_summary": {"en": "ModeA_Cap_Summary", "zh": "模式A产能口径总览"},
    "modea_cap_heatmap": {"en": "ModeA_Cap_Heatmap", "zh": "模式A产能口径热力图"},
    "modeb_cap_summary": {"en": "ModeB_Cap_Summary", "zh": "模式B产能口径总览"},
    "modeb_cap_heatmap": {"en": "ModeB_Cap_Heatmap", "zh": "模式B产能口径热力图"},
    "run_info": {"en": "Run_Info", "zh": "运行信息"},
    "validation_issues": {"en": "Validation_Issues", "zh": "校验问题"},
    "summary_workbook_name": {"en": "Summary of Mode A and Mode B", "zh": "模式A与模式B综合报告"},
    "sheet_title_dashboard_mode": {"en": "Executive Summary - {mode}", "zh": "执行摘要 - {mode}"},
    "sheet_title_dashboard_capacity": {
        "en": "Executive Summary - {mode} Capacity Comparison",
        "zh": "执行摘要 - {mode} 产能口径对比",
    },
    "sheet_title_summary_modes": {"en": "Summary of Mode A and Mode B", "zh": "模式A与模式B综合摘要"},
    "sheet_title_monthly_mode": {"en": "Monthly Trend - {mode}", "zh": "月度趋势 - {mode}"},
    "sheet_title_monthly_capacity": {
        "en": "Monthly Trend - {mode} | Max vs Planned",
        "zh": "月度趋势 - {mode} | Max 对比 Planned",
    },
    "sheet_title_bottleneck_mode": {"en": "Bottleneck - {mode}", "zh": "瓶颈分析 - {mode}"},
    "sheet_title_heatmap_mode": {"en": "{mode} Heatmap", "zh": "{mode} 热力图"},
    "sheet_title_product_risk_mode": {"en": "Product Risk - {mode}", "zh": "产品风险 - {mode}"},
    "sheet_title_planner_summary_mode": {
        "en": "Planner Result Summary - {mode}",
        "zh": "计划员结果汇总 - {mode}",
    },
    "sheet_title_allocation_detail_mode": {"en": "Allocation Detail - {mode}", "zh": "分配明细 - {mode}"},
    "workcenter_filter": {"en": "WorkCenter Filter", "zh": "工作中心筛选"},
    "selection_mode": {"en": "Selection Mode", "zh": "筛选模式"},
    "selection_mode_all": {"en": "All", "zh": "全部"},
    "selection_mode_filtered": {"en": "Filtered", "zh": "筛选"},
    "year": {"en": "Year", "zh": "年份"},
    "supply_mix": {"en": "Supply Mix", "zh": "供给结构"},
    "supply_mix_comparison": {"en": "Supply Mix Comparison", "zh": "供给结构对比"},
    "service_level_comparison": {"en": "Service Level Comparison", "zh": "服务水平对比"},
    "yearly_summary": {"en": "Yearly summary", "zh": "年度汇总"},
    "monthly_detail": {"en": "Monthly detail", "zh": "月度明细"},
    "monthly_comparison_detail": {"en": "Monthly comparison detail", "zh": "月度对比明细"},
    "top_bottleneck_workcenters": {"en": "Top bottleneck workcenters", "zh": "重点瓶颈工作中心"},
    "validation_status_ok": {"en": "Validation Status: OK - No data issues found.", "zh": "校验状态：正常 - 未发现数据问题。"},
    "validation_status_title": {"en": "Data Validation / Issues", "zh": "数据校验 / 问题"},
    "severity": {"en": "Severity", "zh": "级别"},
    "check": {"en": "Check", "zh": "检查项"},
    "detail": {"en": "Detail", "zh": "说明"},
    "unmet_attribution_detail": {"en": "Unmet_Attribution_Detail", "zh": "\u672a\u6ee1\u8db3\u56de\u6302\u660e\u7ec6"},
    "sheet_title_monthly_capacity": {
        "en": "Monthly Trend - {mode} | Max vs Planned",
        "zh": "月度趋势 - {mode} | 最大产能 对比 计划产能",
    },
}


COLUMN_TRANSLATIONS_ZH: dict[str, str] = {
    "Mode": "口径",
    "Year": "年份",
    "Month": "月份",
    "WorkCenter": "工作中心",
    "Metric": "指标",
    "Category": "类别",
    "Tons": "吨位",
    "Demand_Tons": "需求吨位",
    "Internal_Tons": "内部供应吨位",
    "Outsourced_Tons": "外协吨位",
    "Unmet_Tons": "未满足吨位",
    "Supplied_Tons": "已满足吨位",
    "Service_Level": "服务水平",
    "AvgLoadPct": "平均负载率",
    "PeakLoadPct": "峰值负载率",
    "MinLoadPct": "最低负载率",
    "StdLoadPct": "负载波动",
    "Over95Months": "超过95%月份数",
    "PlannerName": "计划员",
    "Product": "产品",
    "ProductFamily": "产品族",
    "Plant": "工厂",
    "AllocationType": "分配类型",
    "RouteType": "路径类型",
    "Priority": "优先级",
    "Allocated_Tons": "已分配吨位",
    "CapacityShare_Pct": "产能占比",
    "Delta": "差异",
    "ModeA": "模式A",
    "ModeB": "模式B",
    "Max": "Max",
    "Planned": "Planned",
    "Basis": "口径",
    "Parameter": "参数",
    "Value": "值",
    "Capacity_Basis": "\u4ea7\u80fd\u53e3\u5f84",
    "Owner_WorkCenter": "\u5f52\u5c5e\u5de5\u4f5c\u4e2d\u5fc3",
    "Capacity_Candidate_WorkCenters": "\u53ef\u56de\u6302\u5de5\u4f5c\u4e2d\u5fc3",
    "Attributed_WorkCenter": "\u56de\u6302\u5de5\u4f5c\u4e2d\u5fc3",
    "Reference_Demand_Tons": "\u53c2\u8003\u9700\u6c42\u5428\u4f4d",
    "Attributed_Unmet_Tons": "\u56de\u6302\u672a\u6ee1\u8db3\u5428\u4f4d",
    "Attribution_Rule": "\u56de\u6302\u89c4\u5219",
    "LoadPct": "负载率",
    "Allocation_Source": "分配来源",
    "Residual_After_Capacity_Tons": "产能分配后剩余吨位",
    "Residual_After_Routing_Tons": "重分配后剩余吨位",
    "Product_Unmet_Tons": "产品总未满足吨位",
    "Scenario_Name": "场景名称",
    "Start_Month": "开始月份",
    "Horizon_Months": "滚动月份数",
    "Input_Load_Folder": "需求输入目录",
    "Input_Master_Folder": "主数据输入目录",
    "Output_Folder": "输出目录",
    "Project_Root_Folder": "项目根目录",
    "Output_FileName": "输出文件名",
    "Run_Timestamp": "运行时间",
    "Run_Mode": "运行模式",
    "Verbose": "详细日志",
    "Skip_Validation_Errors": "跳过校验错误",
    "License_Status": "许可证状态",
    "License_ID": "许可证ID",
    "License_Type": "许可证类型",
    "Licensed_To": "授权对象",
    "License_Expiry": "许可证到期日",
    "License_Binding_Mode": "许可证绑定方式",
    "License_Machine_Label": "许可证机器标签",
    "Notes": "备注",
    "Tool_Version": "工具版本",
    "Severity": "级别",
    "Check": "检查项",
    "Detail": "说明",
    "Max": "最大产能",
    "Planner": "计划产能",
}


VALUE_TRANSLATIONS_ZH: dict[str, str] = {
    "Internal": "内部",
    "Outsourced": "外协",
    "Unmet": "未满足",
    "Demand": "需求",
    "Load%": "负载率",
    "Max Load%": "Max负载率",
    "Planned Load%": "Planned负载率",
    "All": "全部",
    "Filtered": "筛选",
    "OK": "正常",
    "ERROR": "错误",
    "WARNING": "警告",
    "Internal allocated": "内部供应",
    "Residual unmet": "剩余未满足",
    "Selected workcenters": "选中工作中心数",
    "Service level": "服务水平",
    "Total demand": "总需求",
    "ModeA": "模式A",
    "ModeB": "模式B",
    "Max": "Max",
    "Planned": "Planned",
    "ModeA planner owner workcenter": "\u6a21\u5f0fA - \u6309 planner \u5f52\u5c5e\u5de5\u4f5c\u4e2d\u5fc3\u56de\u6302",
    "ModeB baseline capacity workcenter": "\u6a21\u5f0fB - \u56de\u5230 baseline capacity \u5de5\u4f5c\u4e2d\u5fc3",
    "Max": "最大产能",
    "Planner": "计划产能",
    "Max Load%": "最大产能负载率",
    "Planned Load%": "计划产能负载率",
    "Capacity_Base": "基础产能分配",
    "Routing_Reroute": "路径重分配",
    "Primary": "主路径",
    "Alternative": "备选路径",
    "Toller": "外协路径",
    "Capacity": "产能分配",
    "N/A": "不适用",
    "[UNALLOCATED]": "[未分配]",
    "Baseline": "基准",
    "Expansion": "扩张",
    "Lean": "精益",
    "ModeA planner owner workcenter": "模式A - 按计划员归属工作中心回挂",
    "ModeB baseline capacity workcenter": "模式B - 回到基础产能工作中心",
}


SUFFIX_TRANSLATIONS_ZH: dict[str, str] = {
    "_Max": "_最大产能",
    "_Planned": "_计划产能",
    "_ModeA": "_模式A",
    "_ModeB": "_模式B",
    "_Delta": "_差异",
}


SUBSTRING_VALUE_TRANSLATIONS_ZH: list[tuple[str, str]] = [
    ("Allocation_Detail", "分配明细"),
    ("WorkCenter", "工作中心"),
    ("Selection Mode", "筛选模式"),
    ("Filtered", "已筛选"),
    ("All", "全部"),
    ("Scenario", "场景"),
    ("Baseline", "基准"),
    ("Expansion", "扩张"),
    ("Lean", "精益"),
    ("planner", "计划员"),
    ("Planner Result Summary", "计划员结果汇总"),
    ("Planner Product Month Summary", "计划员产品月份汇总"),
    ("Planner Service Level Comparison", "计划员服务水平对比"),
    ("Planner Residual Unmet Comparison", "计划员剩余未满足对比"),
    ("Monthly Service Level Comparison", "月度服务水平对比"),
    ("Max vs Planned", "最大产能 对比 计划产能"),
    ("Max Load%", "最大产能负载率"),
    ("Planned Load%", "计划产能负载率"),
    ("Planned-basis", "计划产能口径"),
    ("Demand", "需求"),
    ("Demand 行", "需求行"),
    ("负载率 行", "负载率行"),
    ("planner 可追溯", "计划员可追溯"),
    ("按 planner 份额", "按计划员份额"),
    ("按 planner 归属工作中心回挂", "按计划员归属工作中心回挂"),
    ("回到 基础产能 工作中心", "回到基础产能工作中心"),
    ("LoadPct", "负载率"),
    ("Load%", "负载率"),
    ("reroute", "重分配"),
    ("baseline capacity", "基础产能"),
    ("overflow", "溢出"),
    ("toller", "外协"),
    ("N/A", "不适用"),
    ("[UNALLOCATED]", "[未分配]"),
    ("Year", "年份"),
    ("Check", "检查项"),
    ("Detail", "说明"),
    ("Severity", "级别"),
]


SHEET_NAME_KEYS: dict[str, str] = {
    "Dashboard": "dashboard",
    "Monthly_Trend": "monthly_trend",
    "Bottleneck": "bottleneck",
    "WC_Heatmap": "wc_heatmap",
    "Product_Risk": "product_risk",
    "Planner_Result_Summary": "planner_result_summary",
    "Allocation_Detail": "allocation_detail",
    "Unmet_Attribution_Detail": "unmet_attribution_detail",
    "Planner_Product_Month": "planner_product_month",
    "Allocation_Summary": "allocation_summary",
    "Outsource_Summary": "outsource_summary",
    "Unmet_Summary": "unmet_summary",
    "Binary_Feasibility": "binary_feasibility",
    "Executive_Comparison": "executive_comparison",
    "Monthly_Trend_Compare": "monthly_trend_compare",
    "Bottleneck_Compare": "bottleneck_compare",
    "WC_Heatmap_Compare": "wc_heatmap_compare",
    "Product_Risk_Compare": "product_risk_compare",
    "Planner_Compare": "planner_compare",
    "ModeA_Cap_Summary": "modea_cap_summary",
    "ModeA_Cap_Heatmap": "modea_cap_heatmap",
    "ModeB_Cap_Summary": "modeb_cap_summary",
    "ModeB_Cap_Heatmap": "modeb_cap_heatmap",
    "Run_Info": "run_info",
    "Validation_Issues": "validation_issues",
}


def ui_text(language: str, key: str, **kwargs: Any) -> str:
    lang = normalize_language(language)
    entry = UI_TEXTS.get(key, {})
    text = entry.get(lang) or entry.get("en") or key
    return text.format(**kwargs) if kwargs else text


def report_text(language: str, key: str, **kwargs: Any) -> str:
    lang = normalize_language(language)
    entry = REPORT_TEXTS.get(key, {})
    text = entry.get(lang) or entry.get("en") or key
    return text.format(**kwargs) if kwargs else text


def localize_mode(language: str, value: str) -> str:
    if normalize_language(language) != "zh":
        return value
    mapping = {"ModeA": "模式A", "ModeB": "模式B", "Both": "同时运行"}
    return mapping.get(value, value)


def localize_sheet_name(language: str, sheet_name: str) -> str:
    if normalize_language(language) != "zh":
        return sheet_name
    key = SHEET_NAME_KEYS.get(sheet_name)
    return report_text(language, key) if key else sheet_name


def localize_value(language: str, value: Any) -> Any:
    if normalize_language(language) != "zh":
        return value
    if not isinstance(value, str):
        return value
    text = value.strip()
    if text in VALUE_TRANSLATIONS_ZH:
        return VALUE_TRANSLATIONS_ZH[text]
    if text in {"ModeA", "ModeB", "Both"}:
        return localize_mode(language, text)
    for source, target in SUBSTRING_VALUE_TRANSLATIONS_ZH:
        text = text.replace(source, target)
    return text


def localize_column_name(language: str, column_name: str) -> str:
    if normalize_language(language) != "zh":
        return column_name
    if column_name in COLUMN_TRANSLATIONS_ZH:
        return COLUMN_TRANSLATIONS_ZH[column_name]
    for suffix, translated_suffix in SUFFIX_TRANSLATIONS_ZH.items():
        if column_name.endswith(suffix):
            base = column_name[: -len(suffix)]
            translated_base = COLUMN_TRANSLATIONS_ZH.get(base, base)
            return f"{translated_base}{translated_suffix}"
    return column_name


def localize_dataframe(language: str, df: pd.DataFrame) -> pd.DataFrame:
    if normalize_language(language) != "zh":
        return df
    localized = df.copy()
    localized.columns = [localize_column_name(language, str(col)) for col in localized.columns]
    for col in localized.columns:
        if pd.api.types.is_object_dtype(localized[col]):
            localized[col] = localized[col].map(lambda value: localize_value(language, value))
    return localized
