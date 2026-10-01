# 待确认排查草稿脚本清单 (共 174 项)

> **说明**：经全量静态审计，以下 174 个脚本均无未提交修改（isGitModified: false），无任何生产业务代码引用，全部为当时排查特定历史问题时遗留的单点实验草稿。按用户授权原则，**本轮暂不执行物理删除，全部保持原位**，待用户审阅后统一确认。

## 模型Prompt与返回格式探测脚本 (16 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/capture_chart_dialog.ps1 | 3357 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/meta_chart.json | 7310 | 否 | test_gen_chart.cjs | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/run_llm_test.cjs | 1453 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_all_intents.cjs | 6262 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_deepseek_variations.cjs | 3492 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_fetch_vars.cjs | 4180 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_flash_chart.cjs | 2548 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_flash_chart_on_user_wb.ps1 | 1150 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_flash_stream.cjs | 2917 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_gen_chart.cjs | 2166 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_quote_normalizer.cjs | 1858 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_user_chart_exact.ps1 | 1342 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_user_intents.cjs | 7297 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_user_new_prompt.cjs | 4058 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_user_prompt_real.cjs | 3321 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |
| scratch/test_user_prompt_vba.cjs | 3304 | 否 | 无 | 针对模型特定图表或Prompt的调用草稿与变体实验，包含硬编码或外部网络请求 | 生成逻辑已固化于 web/src/services/llm.ts，测试沉淀于 test_suite_unit.cjs |

## 其他单点调试脚本 (81 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/check_dumps.ps1 | 502 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/check_running_excel.ps1 | 933 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/debug_scene6.ps1 | 1070 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/dump_grids.ps1 | 2992 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/dump_task2_vba.ps1 | 876 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/extract_all_stages.ps1 | 4203 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/extract_exact_code.ps1 | 1131 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_all_opt.ps1 | 420 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_exact_attempt3.ps1 | 1916 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_exact_error_var1.ps1 | 1543 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_exact_line.ps1 | 1220 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_first_attempt.ps1 | 628 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_line_step_by_step.ps1 | 1018 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/find_start.ps1 | 673 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/get_run_err.ps1 | 900 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/inspect_code_slices.cjs | 447 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/inspect_code_slices2.cjs | 324 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/inspect_proc_14184.ps1 | 3698 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/inspect_sheets.ps1 | 855 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/inspect_user_workbook.ps1 | 2315 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/pinpoint_err.ps1 | 952 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/print_vba_lines.cjs | 396 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/probe_excel_instances.ps1 | 1407 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/reproduce_real_case.ps1 | 1509 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/search_fast.ps1 | 668 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/search_wv2.ps1 | 568 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/spy_runtime_err.ps1 | 3114 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_active_proj.ps1 | 897 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_all_no_comments.ps1 | 1313 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_app_run_args.ps1 | 2768 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_assign_vbproject.ps1 | 614 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_btn_details.ps1 | 1820 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_chinese_sub.ps1 | 958 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_chinese_var.ps1 | 965 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_combined.ps1 | 2889 | 否 | diff_lines.ps1 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_comment_impact.ps1 | 1642 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_dbg.ps1 | 772 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_display_alerts.ps1 | 737 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_entry_detection.ps1 | 7125 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_exact_task1_failure.ps1 | 1855 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_extract_probe.cjs | 3114 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_fc_line.ps1 | 1145 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_feedback_and_two_complex_tasks.ps1 | 16404 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_fix_line.ps1 | 2080 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_multi_proc.ps1 | 1898 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_name_fix.ps1 | 1379 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_new_comp_lines.ps1 | 584 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_param.ps1 | 866 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_param_deep.ps1 | 2702 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_props.ps1 | 1268 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_real_api_history.cjs | 5681 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_real_run.ps1 | 1753 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_req_id.ps1 | 344 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_run_com_ex.ps1 | 848 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_runner_variants.ps1 | 1784 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_sections.ps1 | 1704 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_snippets.ps1 | 3312 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_step_by_step.ps1 | 1679 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_strip_comments.ps1 | 1291 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_subdiv.ps1 | 1230 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_task2_thinking_disabled.ps1 | 5836 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_this_workbook.ps1 | 775 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_three_conditions.ps1 | 6348 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_typo.cjs | 4498 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_undeclared.ps1 | 973 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_underscore_macro.ps1 | 1440 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_user_exact_env.cjs | 4867 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_validator.ps1 | 5629 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_var1_clean.ps1 | 1562 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_vba_runner_error.ps1 | 1114 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_verify_target_wb_arg.ps1 | 1454 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_visible_err.ps1 | 631 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_warehouse_exact_scope_regen_e2e.ps1 | 10531 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_wb_binding_experiment.ps1 | 7514 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_with_history.cjs | 3227 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_wrapper_err_bubble.ps1 | 1064 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_wrapper_isolation_limits.ps1 | 7124 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/test_xlsx_cleanup.ps1 | 2430 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/verify_c3_fix.ps1 | 1381 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/verify_hash.ps1 | 3129 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |
| scratch/verify_task4_exec.ps1 | 2734 | 否 | 无 | 历史一次性单点实验草稿 | 属于阶段性调查草稿，无需保留 |

## Win32窗口与COM交互排查探针 (18 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/check_windows.ps1 | 181 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/dismiss_completion.ps1 | 1548 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/dump_win.ps1 | 2095 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/find_rot_excel.ps1 | 2000 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/get_exact_vbe_highlight.ps1 | 1192 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/get_excel_from_hwnd.ps1 | 3451 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/inspect_vbe_btn.ps1 | 1999 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/print_vbe_slice.ps1 | 681 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/read_vbe_actual_lines.ps1 | 1061 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/read_vbe_lines.ps1 | 998 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_clean_vbe_highlight.ps1 | 1836 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_comments_vbe.ps1 | 1000 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_dialog_dismiss.ps1 | 1892 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_two_workbooks_vbe.ps1 | 1115 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_universal_dismisser.ps1 | 2518 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_unsaved_vbe.ps1 | 1102 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_vbe_cases.ps1 | 1338 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |
| scratch/test_vbe_exact_err.ps1 | 1015 | 否 | 无 | Windows API 对话框弹窗枚举、句柄关闭与消息捕获的碎片脚本 | 必要工具已沉淀并规范化迁移至 tests/tools/ 与 tests/diagnostics/ |

## 编码格式与文本比对草稿 (11 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/compare_ab.ps1 | 1507 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/compare_s1_s2.ps1 | 267 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/diff_lines.ps1 | 760 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/get_gbk.ps1 | 66 | 否 | get_lie_gbk.cjs | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/get_lie_gbk.cjs | 440 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/json_deserialize.ps1 | 2726 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/parse_chunk.ps1 | 993 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/test_crlf_space.ps1 | 1945 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/test_crlf_vs_lf.ps1 | 1827 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/test_meta.json | 1604 | 否 | run_llm_test.cjs, test_dbg.ps1 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |
| scratch/test_proper_utf8.ps1 | 1170 | 否 | 无 | 针对 CRLF 换行、UTF-8 字符集与 JSON 反序列化的中间排查脚本 | 编码处理已内聚至 src/NativeBridge.cs 与 web/src/services/bridge.ts |

## VBE编译与578错误定位探针 (40 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/debug_compile_exact.ps1 | 1059 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/debug_macro_compile.ps1 | 1518 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/find_chart_compile_err.ps1 | 1246 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/find_compile_error_line.ps1 | 1303 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/find_culprit_line.ps1 | 1520 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/find_culprit_split_new.ps1 | 1083 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/find_new_err_line.ps1 | 1195 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/get_compile_err_msg.ps1 | 1398 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/inspect_line41.ps1 | 2022 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_active_proj_compile.ps1 | 1555 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_all_temps_compile.ps1 | 969 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_app_run_syntax_err.ps1 | 785 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_bad_code_578.ps1 | 2619 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_broken_compile.ps1 | 669 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_broken_wb1.ps1 | 1413 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_compile_attempts.ps1 | 982 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_compile_detector.ps1 | 1952 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_compile_generated.ps1 | 1046 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_compile_new.ps1 | 1437 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_compile_vars.ps1 | 860 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_diagnose_compile.ps1 | 1656 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_full_compile.ps1 | 1062 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_full_compile_split_new.ps1 | 666 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_get_err_line.ps1 | 1129 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_line41.ps1 | 1881 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_multi_wb_compile.ps1 | 1403 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_real_compile.ps1 | 1654 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_real_user_wb_compile.ps1 | 1313 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_run_syntax.ps1 | 1661 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_syntax_error_handled.ps1 | 1238 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_syntax_hang.ps1 | 643 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_syntax_investigation.ps1 | 1492 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_typo_compile.ps1 | 849 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_unsaved_compile.ps1 | 2192 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_var1_compile_clean.ps1 | 856 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_var2_compile.ps1 | 856 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_vbe_compile_btn.ps1 | 559 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_vbe_precompile_cases.ps1 | 9714 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_visible_compile.ps1 | 1140 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |
| scratch/test_why_578.ps1 | 1387 | 否 | 无 | 排查特定语法错误行号的二分探针与切片测试，历史调查已闭环 | 核心语法校验逻辑已固化于 src/VbaRunner.cs，单元测试位于 test_suite_unit.cjs |

## 工作簿状态与快照恢复临时脚本 (8 项)

| 相对路径 | 大小 (B) | 未提交修改 | 脚本内调用关系 | 删除理由 | 替代/沉淀情况 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| scratch/inspect_snap_194104.ps1 | 989 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/inspect_snap_content.ps1 | 603 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/probe_budget.cjs | 1904 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/test_snapshot_verify.ps1 | 6924 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/test_state_restore_edge.ps1 | 6711 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/test_task2_with_budget.ps1 | 5957 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/test_warehouse_inventory_scope_cycle.ps1 | 8354 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |
| scratch/verify_budget_and_history.cjs | 6451 | 否 | 无 | 工作簿保存、只读状态与撤销恢复流程的历史验证探针 | 快照引擎已在 src/SnapshotManager.cs 实现，常规断言在 test_suite_unit.cjs |

