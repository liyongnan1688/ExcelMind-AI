# 执行进度日志 (Progress Log)

## Session: 2026-09-25

### 阶段与里程碑
- [x] **Git 隔离基线锁定**
  - 在 `feat/system-quality-fix` 分支锁定提交 `fdab9e1` 作为安全回退基线。
- [x] **根因排查与设计规范纠偏**
  - 移除写死 temperature，转为前端可选配置，不强传参数。
  - 严禁全盘使用 ActiveWorkbook；实现前端-宿主端显式 Target Workbook 绑定与身份核验。
- [x] **1. 拆分聊天与自动化通道 (CHAT vs AUTOMATION)**
  - 实现精确意图识别器 `detectIntent`，覆盖多措辞问答、解释代码、身份咨询、业务操作与歧义澄清。
  - 问答闲聊走 CHAT：零 VBA、零快照、零注入。
- [x] **2. 修复生成协议与代码提取**
  - 彻底清理全局提示词中 D1:L9、35行上限、严禁逐格操作等负向偏置词。
  - 实现结构化代码提取 `extractVbaCode`，彻底废弃脆弱的 `includes('Sub')`。
  - 实现代码截断（未闭合围栏、缺少 End Sub）安全拦截。
- [x] **3. 目标工作簿绑定与执行保护**
  - 宿主与前端锁定具体工作簿（FullName / Name）。
  - C# 规整 ThisWorkbook / ActiveWorkbook 为显式 `Application.Workbooks("...")`。
  - 实现执行前目标激活与执行后身份校验，彻底免疫前台活动窗口切换造成的串改。
- [x] **4. 区分“宏已运行”与“任务完成”**
  - C# 引入写后读回机制 `WorkbookReadback`（UsedRange、首末坐标、公式、边框、底色、单元格抽样）。
  - 前端引入 `verifyExecutionResult`，区分算式文本与纯数字，主观美化如实标为“效果待确认”，严禁伪装满分绿灯。
- [x] **5. 保留诊断能力与极简界面**
  - 气泡保持一句话汇报，默认折叠。
  - ExecutionCard 增加“写后核验”与“审计日志”选项卡，记录脱敏参数、规整步骤与读回数据。
- [x] **6. 跨任务回归测试**
  - `node test_suite_unit.cjs`：30 项意图分类与提取单测 100% 通过。
  - `powershell test_system_suite.ps1`：9 项真实 Excel COM 集成用例 100% 通过。
  - `powershell test_core.ps1`：存量功能测试 100% 通过。

---

## Session: 2026-10-01 (工作区首轮保守清理与文档体系建设)

### 阶段与里程碑
- [x] **基线只读核查与保护范围锁定**
  - 确认当前处于 `feat/system-quality-fix` 分支（HEAD: `dead496`）。
  - 锁定用户当前 20 个包含未提交修改的核心文件（`src/` 核心源码 5 个、`web/src/` 组件 8 个、`test_suite_unit.cjs`、编译产物与草稿脚本），建立强保护边界。
  - 明确区隔目前已验证范围（纯内存单测）与未验证范围（实际 Excel 挂载加载、Ribbon 交互、WebView2 界面渲染、实际宏注入与读回）。
- [x] **第一批纯临时文件精准核对与备份门禁**
  - 从物理磁盘逐项核实 123 个第一批拟清理文件（11 个 bin 独立测试 exe、1 个 bin/scratch 临时 xlsx、47 个 dump 片段、24 个测试隔离 xlsx、40 个过程调试草稿切片）。
  - 建立物理备份目录 `.backup_cleanup_batch1_20261001/`，完成 123 个文件的镜像复制。
  - 备份前完整备份 `task_plan.md` 与 `progress.md`。
  - 生成 `manifest_batch1.json` 并对 123 个文件执行 SHA256 哈希双向全量比对，验证通过率 100%。
  - 生成自包含冲突保护回滚脚本 `rollback.ps1`（无覆盖冲突保护）。
- [x] **安全精准清理执行与空间记录**
  - 删除前对每个文件二次执行原件与备份 SHA256 比对。
  - 使用 `-LiteralPath` 逐项删除 123 个候选文件，生成 `delete_log.json`。
  - 从原位置移除123个未跟踪临时文件，内容总量766710字节，并在同一磁盘保留完整备份；本轮未因此实现相应的磁盘净释放，另产生备份管理文件及文档。
  - 成功删除 123 个文件，失败/中止 0 个。
- [x] **受保护资产基线与状态说明**
  - 清理前受保护资产基线曾生成，但其临时目录在流程结束时被移除，因此目前无法独立复核清理前后内容一致性。
  - 明确记录 `.protected_baseline_tmp/` 为本次执行中创建并移除的辅助临时目录，不属于批准清理的123个原有文件。
  - 当前 44 个核心受保护文件（`src/`、`web/`、根目录构建/运行/测试脚本、`bin/` 核心运行时）及 `scratch/` 下 19 个独立 `.cs` 测试工具源码与核心日志均存在。
- [x] **限制性 `.gitignore` 规则创建**
  - 建立标准 `.gitignore`，仅针对完全清空的测试生成目录（`bin/scratch/`、`scratch/entry_test_isolated/` 等）及本次备份目录。
  - 针对尚存元数据文件的目录（`acceptance_isolated`、`complex_feedback_isolated`）严格暂不加入忽略，未执行 `git rm --cached`。
- [x] **前后自动化回归测试通过**
  - 清理前：`node test_suite_unit.cjs` (PASS: 27, FAIL: 0), `node test_regex_counter_example.cjs` (PASS: 3, FAIL: 0)。
  - 清理后：`node test_suite_unit.cjs` (PASS: 27, FAIL: 0), `node test_regex_counter_example.cjs` (PASS: 3, FAIL: 0)。
- [x] **项目规划书与技术说明书完工**
  - `task_plan.md` 扩充为全局《项目规划书》，完整保留历史专项记录。
  - 新建 `TECHNICAL_SPEC.md` 作为正式《技术说明书》。

### 最终状态
第一批清理及文档更新已完成；备份一致性与Node测试已核验；清理前内容基线未保留；实际Excel运行待人工验证。

---

## 2026-10-01 (第二批：scratch 资产整理、长期目录与安全规则建设)

### 执行前物理核对与纠偏
- **物理统计闭环**：对 scratch/ 剩余 251 个文件进行字节级核算，总计 **1,043,103 字节**（完全无误差闭环）。
- **数据纠偏与对账**：
  - A 类必要测试工具（16个）：实际测量为 **103,537 字节**（纠正前期混入候选项文本导致的 108,187 字节笔误）。
  - B 类必要历史证据（9个）：实际测量为 **58,619 字节**（纠正前期混入额外日志的 68,092 字节）。
  - C 类明确处理项（41个）：实际测量为 **331,224 字节**（10个测试exe 310,272字节 + 1个临时dll 15,872字节 + 10个已解决探针cs 5,080字节 + 20个抓包文本0/1000字节）。
  - D 类受保护活跃脚本（11个）：实际测量为 **169,546 字节**（含用户未提交修改的 2 个脚本 15,673 字节及 9 个强依赖驱动 153,873 字节）。
  - 待审排查草稿脚本（174个）：实际测量为 **380,177 字节**。
  - 合计：`103,537 + 58,619 + 331,224 + 169,546 + 380,177 = 1,043,103 字节`。

### 备份与持久化基线
- **独立第二批物理备份**：在 `.backup_cleanup_batch2_20261001/` 完成获批处理的 66 个文件完整物理备份，绝不混入第一批备份。
- **Manifest 与回滚脚本**：
  - 生成 `manifest_batch2.json`，经双向 SHA256 与字节核验，66 个文件校验通过率 100%。
  - 生成 `rollback.ps1`，具备哈希校验与冲突保护机制，经 `-WhatIf` 试运行验证成功。
- **受保护资产清理前基线**：
  - 提取 `src/` (5个)、`web/src/` (13个核心文件)、`test_suite_unit.cjs`、`test_regex_counter_example.cjs` 及 `scratch/` 11 个活跃脚本的 SHA256 哈希与字节，持久化保存在 `.backup_cleanup_batch2_20261001/protected_baseline_batch2.json`（共 31 个核心文件，472,921 字节）。

### 物理执行事实
- [x] **A 类测试与诊断工具规范化迁移 (16个)**：
  - 9 个 C# 测试工具源码整理至 `tests/tools/`；
  - 7 个诊断脚本整理至 `tests/diagnostics/`；
  - 对写文件工具（`TestWbBinding.cs`、`VerifyPayload.cs`、`MemoryDump.cs`）改造增加明确命令行输出目录支持，默认写入 `.artifacts/tests/`；经系统 `csc.exe` 独立编译验证语法无误；
  - 原 scratch/ 对应 16 个文件完成清理。
- [x] **B 类必要历史证据归档 (9个)**：
  - 9 个核心日志与已验证证据文件完整迁移至 `docs/history/evidence_202609/`，哈希比对 100% 一致；原 scratch/ 对应 9 个文件已清除。
- [x] **C 类已失效生成物与历史排查文件物理删除 (41个)**：
  - 10 个测试 exe、1 个临时 dll、10 个历史已解决排查 cs、20 个抓包文件在核对哈希后安全删除；
  - 释放 scratch 物理占用 331,224 字节；
  - 执行日志持久化记录于 `.backup_cleanup_batch2_20261001/delete_log_batch2.json`；已清理空的子目录 `scratch/acceptance_isolated` 与 `scratch/complex_feedback_isolated`。
- [x] **D 类当前活跃脚本强保护 (11个)**：
  - 用户修改中的 `scratch/call_llm.cjs`、`scratch/test_compile_behavior.ps1` 及其 9 个强依赖驱动原封不动保留在 `scratch/` 原位，零变动。
- [x] **174 个排查草稿全量静态审计与清单生成**：
  - 经审计，174 个拟删除脚本中包含未提交修改的文件数为 **0**；
  - 详细审计结果已持久化保存至 `.backup_cleanup_batch2_20261001/pending_174_audit.json`；
  - 174 个脚本全量保持在 `scratch/` 原位，等待用户明确批准后执行删除。

### 长期规范与安全清理建设
- [x] **新增安全清理入口 `scripts/clean_artifacts.ps1`**：
  - 默认仅预览不删除；必须显式传入 `-Apply` 才执行物理删除；
  - 作用域严格限定为 `.artifacts/tests/` 和 `.artifacts/tmp/`；
  - 默认保留最近 14 天产物；支持 `-RunId` 筛选；
  - 拒绝路径穿越、符号链接与重解析点；对无元数据目录只报告不自动删除；跳过包含 `.running` 标记的运行目录；不强杀进程。
  - 使用全新模拟运行目录进行了完整的预览与 `-Apply` 验证（绝未触碰真实业务或备份文件）。
- [x] **规范文档编制与同步**：
  - 新建 `docs/FILE_LIFECYCLE.md`：定义各类文件生命周期、存放位置、清理与晋升机制。
  - 新建 `AGENTS.md`：规范 AI 开发行为，禁止污染 bin/ 和根目录，禁止未授权自动运行真实 API 测试。
  - 新建 `tests/README.md`：详细列出 16 个保留工具的具体用途、依赖、编译方式、运行命令、输出位置和副作用。
  - 更新 `TECHNICAL_SPEC.md`：同步新目录体系与职责。
  - 更新 `.gitignore`：加入 `.artifacts/`、`.backups/` 及第二批备份目录。

### 回归与审计结果
- **Node 自动化测试**：
  - `test_suite_unit.cjs`：27/27 PASS
  - `test_regex_counter_example.cjs`：3/3 PASS
- **受保护资产比对**：与 `.backup_cleanup_batch2_20261001/protected_baseline_batch2.json` 比对，核心源码、前端、受保护活跃脚本 100% 存在且哈希无变动。
- **环境安全准则遵循**：本轮未启动或关闭用户 Excel、未调用真实模型、未读取秘密值、未联网、未升级依赖。

---

## 2026-10-01 (最后一批：174 个待审脚本清理与最终收尾)

### 执行前门禁与前置状态
- **前置文件统计**：本轮执行前，`scratch/` 实际存有 **185 个文件**（含 11 个活跃脚本及 174 个待审脚本）。
- **清单与依赖核对**：
  - 读取 `docs/history/pending_174_scripts_manifest.md` 与 `.backup_cleanup_batch2_20261001/pending_174_audit.json`，确认两者精确包含相同的 174 个唯一路径，全部位于 `scratch/` 内。
  - 确认 174 个脚本未包含未提交修改（`isGitModified: false` 全部通过）。
  - 引用核验：除清单文档自身外，`src/`、`web/`、`tests/`、`docs/history/evidence_202609/`、11 个活跃脚本及根目录构建/运行入口对这 174 个脚本的引用数严格为 **0**。
  - 保存本次批准清单固定副本至 `.backups/cleanup_pending_scripts_20261001/approved_174_fixed_manifest.json`，SHA256 为 `5C92C09CFDAF1DC2A9842C42172A5FE2E407326CAF0DAD1208513DE16AD04FF3`。

### 独立备份与物理执行
- **独立备份建立**：
  - 建立独立备份目录 `.backups/cleanup_pending_scripts_20261001/`（绝未混入第一批或第二批备份）。
  - 完整复制 174 个文件至备份目录，生成 `manifest_pending_174.json`，双向哈希比对通过率 100%（共计 380,177 字节）。
  - 生成具备哈希校验和目标冲突保护的回滚脚本 `rollback.ps1`，经 `-WhatIf` 试运行验证 174 项完全一致。
  - 保存受保护资产清理前基线 `protected_baseline_batch3.json`（60 个核心文件）。
- **物理删除执行**：
  - 依据固定批准清单，逐项核对原件与备份 SHA256 哈希后执行物理删除。
  - 实际删除：**174 个**；跳过：**0 个**；释放物理空间：**380,177 字节**。
  - 执行日志持久化记录于 `.backups/cleanup_pending_scripts_20261001/delete_log_batch3.json`。
  - 未使用递归删除，未使用 `git clean`、`git reset` 或批量取消跟踪。

### scratch 最终状态与活跃脚本例外说明
- **最终文件数量**：`scratch/` 目录从清理前的 185 个文件缩减至最终严格为 **11 个文件**。
- **剩余 11 个文件清单**：
  1. `scratch/call_llm.cjs` (含未提交修改)
  2. `scratch/test_compile_behavior.ps1` (含未提交修改)
  3. `scratch/run_comprehensive_acceptance.ps1`
  4. `scratch/run_full_e2e_acceptance.ps1`
  5. `scratch/test_complex_task6_isolated.ps1`
  6. `scratch/test_generic_capabilities_e2e.ps1`
  7. `scratch/test_scope_correction_and_recovery_e2e.ps1`
  8. `scratch/test_tasks_4_to_6.ps1`
  9. `scratch/test_two_complex_tasks_e2e.ps1`
  10. `scratch/test_warehouse_inventory_clean_run.ps1`
  11. `scratch/test_warehouse_with_budget.ps1`
- **暂时例外说明**：
  - 当前保留这 11 个脚本属于**暂时例外**，仅为保护用户当前未提交修改及强依赖驱动链路；
  - 绝不代表允许以后继续向 `scratch/` 堆积测试产物；
  - 待后续改动稳定后，再单独处理路径解耦、输出目录标准化与迁移晋升。

### 回归验证与受保护资产审计
- **Node 自动化测试**：
  - 清理前：`test_suite_unit.cjs` (PASS: 27, FAIL: 0), `test_regex_counter_example.cjs` (PASS: 3, FAIL: 0)
  - 清理后：`test_suite_unit.cjs` (PASS: 27, FAIL: 0), `test_regex_counter_example.cjs` (PASS: 3, FAIL: 0)
- **受保护资产比对**：
  - 对照基线 `protected_baseline_batch3.json`（涵盖 `src/` 5个 C# 核心类、`web/src/` 全部组件与服务、`tests/` 16个工具、`docs/history/` 9个证据、11个活跃脚本及所有构建运行脚本共 60 项资产）；
  - 比对结果：**匹配 60 项，缺失 0 项，哈希不一致 0 项（100% 完好无损）**。

### 最终总结
本轮文件整理完成，实际Excel运行待人工验证。

---

## 2026-10-01 (专项遗漏纠偏：测试工具失效路径修复与清理脚本 RunId 规则加固)

### 一、 前期调查遗漏澄清与复盘
- **遗漏事实如实记录**：
  此前向用户的报告中曾称“A 类测试工具内部无 scratch 硬编码、迁移到 tests/ 后无需修改路径”。经后续严密代码审查证实，**该结论严重失真，属于前期调查分析中的重大遗漏**，绝不仅是历史遗留技术债。
- **具体失效硬编码清单**：
  1. `FindDialogButtons.cs`、`ReadDialogText.cs`、`RegressionTest.cs`、`TestThreeConditions.cs`：均硬编码读取 `scratch/VERIFIED_STAGE2_ORIGINAL.vba`；
  2. `VerifyCompleteDelivery.cs`：硬编码读取 `scratch/VERIFIED_STAGE2_ORIGINAL.vba`，且向 `scratch/isolated_audit_run` 写入测试工作簿；
  3. `RunTaskPaneFrontEndRegression.cs`：硬编码读取 `scratch/stage1_rawModelResponse.vba` 与 `scratch/VERIFIED_STAGE2_ORIGINAL.vba`，且向 `scratch/frontend_regression_run` 写入测试工作簿；
  4. `read_vbe_dialog.ps1`：硬编码读取已删除的历史过程文件 `scratch/generated_split_real.txt`，且在用户桌面通过通配符搜索 `*20260908*.xls*`。

### 二、 修复前备份与受保护基线
- **获批文件物理备份**：在 `.backups/pre_path_fix_20261001/` 完成 7 个工具及相关文档共 12 个目标的完整物理备份，记录 `manifest_pre_fix.json`。
- **受保护资产基线持久化**：生成 `.backups/pre_path_fix_20261001/protected_baseline_pre_fix.json`，涵盖 `src/` (7个)、`web/` (16个)、`bin/` (11个)、根目录构建运行入口 (4个) 及 `scratch/` (11个受保护脚本)，共 49 项资产，记录 SHA256 与字节数。

### 三、 路径与生命周期修复实施
1. **输入路径规范化**：
   - 优先接受命令行入参；无入参时，具有语义对应已核验样本的工具默认回退到 `docs/history/evidence_202609/` 中的固定历史样本；
   - 对已删除且无可靠替代的 `read_vbe_dialog.ps1`，彻底移除桌面搜索逻辑，强制要求调用者显式传入 `-WorkbookPath` 与 `-VbaSnippetPath`；
   - 路径解析通过 `ResolveProjectRoot` 向上层级探测或环境变量解析，消除对调用者当前工作目录的假设。
2. **前置阻断安全门禁**：
   - 所有 7 个工具在执行任何 Excel COM 自动化（`Activator.CreateInstance` / `New-Object -ComObject`）、宏执行、进程杀死（`KillExcel`）之前，首先进行输入参数与文件存在性校验；若输入缺失或非法，直接向控制台输出错误并返回退出码 1，绝不触发外部调用。
3. **输出隔离与生命周期管理**：
   - 消除所有向 `scratch/` 写入测试生成物的逻辑；输出目录默认重定向至 `.artifacts/tests/<tool_name>_<timestamp>_<guid>/`；
   - 运行期间在目标目录维护 `.running` 锁文件，正常或异常退出时在 `finally` 中移除 `.running` 锁；
   - 在运行目录生成标准 `meta.json`（记录 `runId`、`toolName`、`timestamp`、`status`、`details`），与清理脚本完全兼容。
4. **独立程序集动态解析（AssemblyResolve）**：
   - 为避免 C# 工具编译到 `.artifacts/tests/<run-id>/` 独立目录运行时，CLR JIT 因在 EXE 目录找不到 `LeeExcel.dll` 或 WebView2 依赖而抛出 `0xE0434352` (FileNotFoundException)，在 `RegressionTest.cs`、`VerifyCompleteDelivery.cs`、`RunTaskPaneFrontEndRegression.cs` 中拆分入口并注册 `AppDomain.CurrentDomain.AssemblyResolve`，自动从项目 `bin/` 及 `packages/` 检索加载依赖。

### 四、 清理脚本 `clean_artifacts.ps1` 规则加固与核查
- **StrictMode 兼容修复**：修复 PowerShell `Set-StrictMode -Version Latest` 下访问单一标量对象 `.Count` 抛错的问题，采用显式数组防护。
- **完全精确匹配**：实现 `-RunId` 绝对精确匹配，拒绝通配符（`*`、`?`）与正则模式。
- **非法输入严格拒绝**：拒绝包含路径分隔符（`\`、`/`）、目录穿越（`..`）或特殊字符的 RunId。
- **跨作用域歧义保护**：当 `tests/` 与 `tmp/` 同时存在同名 RunId 时，主动报错并中止，要求调用者通过 `-Scope Tests` 或 `-Scope Tmp` 消除歧义，绝不同时删除。
- **运行锁生效**：显式 RunId 虽可绕过 RetentionDays 限制，但检测到 `.running` 标记时依然跳过。
- **默认只读预览**：默认仅展示扫描清单，必须显式附加 `-Apply` 方执行物理删除。
- **模拟验证全部通过**：通过模拟目录针对上述 4 项规则进行了全项黑盒测试，验证完毕后物理清除模拟目录。

### 五、 最终回归与基线比对结果
- **C# 编译验证**：6 个 C# 工具使用系统 64 位 `csc.exe` 独立编译到 `.artifacts/tests/`，全部编译通过（PASS: 6/6）；验证后临时目录物理清除。
- **PowerShell 语法解析**：`clean_artifacts.ps1` 与 `read_vbe_dialog.ps1` 通过 `System.Management.Automation.Language.Parser` 语法树解析（PASS: 2/2）。
- **输入缺失阻断测试**：7 个工具在传入不存在路径或缺失参数时，100% 成功在任何 COM 调用前拦截并返回退出码 1。
- **离线核心测试套件**：
  - `test_suite_unit.cjs`：27/27 PASS
  - `test_regex_counter_example.cjs`：3/3 PASS
- **受保护资产基线比对**：
  - 对照基线：`.backups/pre_path_fix_20261001/protected_baseline_pre_fix.json`（49 项核心资产）；
  - 比对结果：**匹配 49 项，缺失 0 项，哈希不一致 0 项（100% 完好无损）**。
- **环境安全合规**：本轮未启动或关闭用户 Excel、未调用真实模型、未读取密钥、未联网、未执行实际 COM 宏写入。

### 最终结论
本轮文件整理与路径修复全部完成，实际Excel运行待人工验证。

---

## Session: 2026-10-02 (R0 阶段实施：规划治理、基线固定与任务体系建立)

### 阶段目标与授权范围
- **授权约束**：严格限定为 **R0 阶段**，不自动开发 R1–R7 任何业务功能。
- **核心目标**：全面核实真实仓库能力与资产位置，建立唯一主规划，固化 v1.2.0 可用基线与最小回归清单，预备 R1a 接口映射。

### 阶段实施细节与事实记录
- [x] **真实仓库功能与资产深度审计 (TASK-R0-01)**：
  - 核实物理存储路径：宏库位于 `%APPDATA%\ExcelMindAI\Scripts\`（含旧目录自动迁移），快照位于 `%APPDATA%\ExcelMindAI\Backups\`（含救援副本），API 配置位于 WebView2 `localStorage['excelmind_ai_llm_config']`，聊天记录当前在 Svelte 运行时内存。
  - 严格区分 5 类状态：已实现、自动测试通过、用户人工验收、未验证、已知限制；彻底杜绝历史 Office.js 概念混淆。
- [x] **建立唯一主规划路线图 (TASK-R0-03)**：
  - 新建 `docs/product-roadmap.md` 作为项目唯一定位与主规划源，系统化纳入《ExcelMind AI 产品迭代规划书》。
  - 详细定义不可退化原则、R0–R7 演进顺序、每阶段范围、不做项、验收标准与停点。
  - 建立 6 个开源项目（OpenRefine、Rubberduck、ClosedXML、VBA-Web、xlwings、duckdb-excel）的研究参考规范台账，明确许可证要求与边界约束。
  - 建立具备唯一任务 ID（`TASK-R0-01` ~ `TASK-R6c-01`）的任务跟踪表，采用 6 级严格状态机。
- [x] **固化可用版本基线与最小回归清单 (TASK-R0-02)**：
  - 锁定可用基线：`231ae3a` (Release v1.2.0)。
  - 确立 8 项不可退化最小回归：对话不执行、操作写值、无代码回复不伪报宏失败、导入保存重开、源码保真、运行错误展示、快照恢复、冷启动无重复注册。
- [x] **离线门禁自动化验证**：
  - 运行 `node test_suite_unit.cjs`：27/27 PASS（覆盖显式双通道、代码截断拦截、异步切换保护、写后核验比对等）。
  - 运行 `node test_regex_counter_example.cjs`：3/3 PASS（验证禁用正则静默改写的必要性与反例杜绝）。
- [x] **R1a 最小切片接口与文件映射预备 (TASK-R0-04)**：
  - 完成“用户显式附加选区 → 只读上下文卡预览 → 确认发送”的架构设计。
  - 定义了 `get_selection_context` 请求响应契约与 `SelectionContextDto` 数据结构。
  - 明确映射至 `src/NativeBridge.cs`、`src/VbaRunner.cs`、`web/src/components/ChatInput.svelte` 与 `web/src/services/llm.ts`。
  - 严格遵守纪律，未编写任何 R1a 业务代码，等待用户授权。

### 最终状态
R0 阶段全部目标已达成并通过离线验证；当前工作区干净，无未授权代码变动；停下等待用户验收 R0 并授权启动 R1a。

---

## Session: 2026-10-02 (R1a 阶段实施：用户显式附加选区 → 只读预览 → 按用户选择随本次请求发送)

### 阶段目标与授权范围
- **授权约束**：严格限定为 **R1a 阶段** 最小切片，不实施 R1b–R7；不修改模型 VBA、不增加语法/算法限制；不改变现有操作/对话路由、宏执行器、快照和宏库行为。
- **两项短核对**：
  1. **宏库真实目录**：`%APPDATA%\ExcelMindAI\Scripts\`（包含 `.bas` 与 `.meta.json`）。启动时代码自动将旧 `%APPDATA%\LeeExcel\Scripts\` 数据无损迁移复制，旧宏完全可见；历史报告出现的 `ExcelMind\scripts` 确认为书写笔误，已在主规划文档中统一修正。
  2. **会话历史状态**：`web/src/App.svelte` 中的 `messages` 数组完全属于纯前端运行时内存状态，刷新/重开即重置，明确标记为已知限制，R1a 严格未做跨会话持久化。

### 阶段实施细节与文件修改
- [x] **新增只读选区服务 (TASK-R1a-01)**：
  - 新建 `src/SelectionContextService.cs`，彻底独立于 `VbaRunner.cs`，专职只读安全读取。
  - 严格限制：零写入、零格式化、零激活工作表、零选区改变、零宏注入、零大模型调用、零 API Key 触碰。
  - COM 安全采样：绝不加载整表或整列 Value2，预先限制样本行列上限（默认行上限 5，列上限 15，单元格文本上限 100 字符，超长标注 `...[截断]`）。
  - 友好降级机制：多区域（`Areas.Count > 1`）与非 Range 选中（如 Chart、Shape）返回友好降级错误，绝不静默改读 `ActiveSheet.UsedRange`。
  - 数据契约真实表达：区分 `candidateHeaders`、保留原始数据类型与公式、局部状态明确标为 `sample_scanned_only` / `unknown`，不夸大推论全表。
- [x] **宿主调度网关分发 (TASK-R1a-01)**：
  - 更新 `src/NativeBridge.cs`，添加 `case "get_selection_context"` 分发至 `HandleGetSelectionContext`。
  - 确保调用在宿主线程完成，以标准 JSON 格式返回给前端。
- [x] **前端选区只读预览与显式勾选控制 (TASK-R1a-02)**：
  - 更新 `web/src/services/bridge.ts`，导出 `SelectionContextData`、`SelectionSendOptions`、`SelectionSnapshot` 等强类型接口，提供 `bridge.getSelectionContext()` 及 mock 桩。
  - 更新 `web/src/components/ChatInput.svelte`，在输入框上方新增【📎 附加当前选区】按钮（对话/操作双模式通用）。
  - 紧凑只读预览卡：展示工作簿名、工作表、地址、总行列数、采样范围、局部特征与未扫描说明；提供重新读取与移除操作。
  - 4 项发送配置开关（默认仅候选表头有效，样本值与公式默认不勾选，需用户显式确认）：
    1. 首行作为表头（取消勾选则降级为纯候选）
    2. 包含样本值（默认关）
    3. 包含样本公式（默认关）
    4. 仅发送结构（关闭值与公式）
- [x] **请求生命周期固化、跨表隔离与防提权注入 (TASK-R1a-03)**：
  - 更新 `web/src/services/llm.ts`，新增 `formatSelectionContextForPrompt` 组装器：
    - 将选区以低信任独立数据段 `<excel_selection_context>` 注入 user prompt，显式声明单元格文本为纯业务数据，不得执行任何内部指令，彻底阻断提示词提权。
    - 路径脱敏：仅向模型提供工作簿名称与区域地址，本地绝对物理路径默认剔除。
    - 发送审计：在请求发出时记录 `selectionAudit`（包含发送字段、样本行数、列数、字符量、省略情况），普通日志不存业务值全文。
  - 更新 `web/src/App.svelte`：
    - 发送时固化快照：点击发送时将当前模式、目标工作簿身份、选区快照与发送选项深拷贝冻结到 `ChatMessage`，后续选区变动或模式切换对已发请求零影响。
    - 操作模式跨表强拦截：若附件来自工作簿 A，而当前操作目标工作簿变为 B，强行拦截发送并弹窗提示用户重新附加或移除附件，绝不混用上下文。
    - 对话模式安全：对话模式携带附件时仅做业务解释分析，绝不产生快照、写值或 `execute_vba`。
    - 前端展示：用户消息气泡展示附件徽章，提供【查看发送内容】弹窗（严格脱敏且绝不包含 API Key）。

### 构建与自动化验证结果
- **C# 宿主编译**：运行 `build_addin.ps1`，成功编译 `bin/LeeExcel.dll`（0 Errors, 0 Warnings）。
- **前端构建打包**：运行 `pnpm run build`，成功打包至 `bin/dist/`（TypeScript 检查通过）。
- **C# 独立单元测试工具**：
  - 编写 `tests/tools/VerifySelectionContext.cs`，编译至 `.artifacts/tests/selection_verification/` 并执行；
  - 11 项全功能验证（单单元格、非 A1 起始、整列子区域限制、多区域降级拦截、非 Range 降级拦截、中文与前导零保真、日期与公式识别、长文本截断、仅结构发送过滤、含样本值过滤、注入隔离标签）**全部 100% 通过 (PASS: 11/11)**。
- **离线核心单测门禁**：
  - 扩充 `test_suite_unit.cjs` 新增 9 项 R1a 选区与提示注入隔离测试；
  - 运行 `node test_suite_unit.cjs`：**36/36 PASS (100%)**。
  - 运行 `node test_regex_counter_example.cjs`：**3/3 PASS (100%)**。

### 最终状态
R1a 阶段代码与构建已完成，离线与单元测试 100% 通过；`TASK-R1a-01` ~ `TASK-R1a-03` 状态标记为 `awaiting_manual`；停下等待用户人工验收，不自动进入 R1b。

---

## 缺陷排查与纠偏 (2026-10-02：对话/操作点击发送无响应缺陷修复)

### 1. 现象与根因定位 (Systematic Debugging)
- **用户反馈**：在对话模式附加选区并输入“这几行有什么”后，点击【💬 对话】或按 Enter 无法发送，消息未进入列表，文字未清空。
- **排查证据与根因**：
  1. 静态扫描 `web/src/App.svelte` 模板标签：发现第 546 行 `<Zap size={10} />` 与第 549 行 `<MessageSquare size={10} />` 在用户气泡的 `user-mode-tag` 中被直接使用；
  2. 但在 `App.svelte` 顶部的 `<script>` 导入区中，**遗漏了 `import { Zap, MessageSquare } from 'lucide-svelte';`**；
  3. Vite 编译时将其降级为全局裸调用 `MessageSquare(_a, { size: 10 })`；当用户点击发送并向 `messages` 插入第一条带 `mode: 'CHAT'` 的消息时，Svelte 触发微任务重绘，直接抛出未捕获的 **`ReferenceError: MessageSquare is not defined`**；
  4. 该 ReferenceError 导致 Svelte 渲染管道瞬间崩溃卡死，UI 更新冻结，输入框内容无法反映清空，后续点击亦全部失效。

### 2. 修复与防御加固措施
- **补齐组件导入**：在 `web/src/App.svelte` 顶部显式补齐 `import { Zap, MessageSquare } from 'lucide-svelte';`，全仓扫描确认无其他组件遗漏。
- **数据兜底防御**：在 `web/src/services/llm.ts` 的 `formatSelectionContextForPrompt` 中，对 `candidateHeaders` 增加防空数组兜底（`Array.isArray(ctx.candidateHeaders) && ... : ['列1']`）。
- **组件异常保护**：在 `web/src/components/ChatInput.svelte` 的 `handleSubmit` 中加入 try-catch 保护，若组装遇到异常直接在输入框上方错误条展示，杜绝任何静默崩溃。
- **前端构建产物更新**：重新执行 `pnpm run build`，编译产物 `bin/dist/assets/index-DasbtHtN.js` 已完成打包，`Zap` 与 `MessageSquare` 100% 正确混淆引用。
- **离线核心单测回归**：`node test_suite_unit.cjs` (36/36 PASS), `node test_regex_counter_example.cjs` (3/3 PASS)。

---

## 缺陷排查与修复 (2026-10-02：Excel 工作表与任务窗格输入焦点交接缺陷)

### 1. 现象与真实根因判定 (Systematic Debugging)
- **缺陷现象**：在插件文本框输入“插件草稿”后，鼠标点击 Excel 单元格（如 A1），键盘输入依然进入插件文本框，未进入单元格。
- **根因判定**：
  - 排除根因 B（异步回调抢焦点）：审查前端全仓，流式输出、API 返回、选区读取、模式切换**零 focus()、零 autofocus 代码**。
  - 排除根因 C（全局按键截获）：全仓无全局 Windows 键盘 Hook，按键绑定仅局限在 `<textarea>` 自身。
  - **确认根因 A（WebView2 未释放 Win32 焦点）**：
    - CustomTaskPane 与 Excel 主窗口属于同一 UI 线程。
    - 点击工作表 `EXCEL7` 窗口时，Excel 改变了选区框绘制，但未显式调用 Win32 `SetFocus(hwndExcel7)`。
    - 系统的 Win32 键盘焦点句柄 `GetFocus()` 依然停留在 WebView2 内部子窗口（`Chrome_RenderWidgetHostHWND`），后续按键被操作系统直接投递给 Chromium。

### 2. 最小改动修复实施
- **`src/TaskPaneFocusHelper.cs` (新增)**：
  - 封装 Win32 API：`GetFocus`、`SetFocus`、`IsChild`、`GetParent`、`FindWindowEx`、`GetClassName`。
  - 实现 `IsChildOrSame`：支持直接句柄比对、`IsChild` 及 `GetParent` 递归向上回溯（64层防环）。
  - 实现 `GetActiveWorksheetHwnd`：安全枚举当前前台工作簿的 `EXCEL7` 窗口句柄。
  - 实现 `RelinquishFocusToExcel`：仅在焦点确实在任务窗格内部时才安全让出焦点。
  - 实现 `TaskPaneFocusMessageFilter : IMessageFilter`：纯 WinForms 进程内消息过滤（非系统钩子），监听鼠标点击 `WM_LBUTTONDOWN` 等消息，当点击在任务窗格外且焦点在窗格内时，显式将焦点平滑交还目标控件（如 `EXCEL7`）；**永远返回 false**，不吞掉任何消息。
- **`src/TaskPaneControl.cs` (修改)**：
  - 构造函数中向 `Application` 注册 `TaskPaneFocusMessageFilter`。
  - 在 `Dispose` 中注销 `TaskPaneFocusMessageFilter`。
  - 公开 `RelinquishFocusToExcel` 供 COM 事件调用。
- **`src/LeeExcelAddIn.cs` (修改)**：
  - 在 `AutoOpen` 中为 `SheetSelectionChange`、`SheetBeforeDoubleClick`、`SheetBeforeRightClick` 增加焦点归还双重保险。

### 3. 自动化验证与门禁全绿
- **插件全量编译**：运行 `build_addin.ps1`，成功编译 `bin/LeeExcel.dll`（0 Errors, 0 Warnings）。
- **焦点机制独立单元测试**：
  - 编写 `tests/tools/VerifyFocusHandover.cs`，编译至 `.artifacts/tests/focus_handover/` 并执行；
  - 13 项单元验证（句柄判定、边界防护、WinForms 控件树判定、消息过滤不吞消息、空指针抗崩溃保护）**全部 100% 通过 (PASS: 13/13)**。
- **离线核心单测门禁**：
  - 运行 `node test_suite_unit.cjs`：**36/36 PASS (100%)**。
  - 运行 `node test_regex_counter_example.cjs`：**3/3 PASS (100%)**。

### 4. 验收结论与状态锁定
- **用户实测通过**：用户已在前台真实 Excel 运行中完成实测，在插件输入后点击单元格，新输入能够正确进入单元格，确认焦点交接功能验证成功。
- **状态统一更新**：
  - R1a 选区感知功能（`TASK-R1a-01` ~ `TASK-R1a-03`）状态统一锁定为：`accepted`。
  - 焦点交接缺陷（`TASK-FIX-FOCUS-01` / `BUG-FOCUS-01`）状态统一锁定为：`accepted`（正式关闭）。
- **事实边界保留**：其他未深度覆盖边界（中文输入法组合、公式栏、双击编辑、极端长流式前台表现）如实标记为待验，绝不泛化为“全部交互通过”。
- **下一阶段准入**：R1b 切片 1（选区采集时间、快照提示及刷新原区域）已获得用户明确授权启动并完成开发。

---

## Session: 2026-10-02 (R1b 切片 1: 选区采集时间、快照提示与刷新原区域)

### 1. 目标与实现范围
- **目标切片**：仅限实施选区采集时间、快照提示及【🔄 刷新原区域】；坚决不实现宏引用与 Token 估算。
- **开发分支**：独立分支 `feat/r1b-refresh-selection`。
- **数据契约增量**：
  - 增量引入 `capturedAt: number`（Unix 毫秒时间戳）与 `attachmentId: string`（唯一附件标识）。
  - 保留 R1a 全部字段：`headers`、`candidateHeaders`、`sampleRows` 二维单元格结构、`formulaStatus`、`mergeStatus`、`visibilityStatus`、`unscannedNotes`。
- **定向刷新原区域机制**：
  - 卡片【🔄 刷新原区域】：宿主通过 `targetWorkbookName`、`targetWorkbookFullName`、`targetSheetName`、`targetAddress` 查找原对象。
  - 严格不猜测同名表、不回退到 ActiveWorkbook / 当前 Selection，不擅自激活工作簿或改变用户当前选区。若原表找不到报错 `sheet_not_found`，原区域找不到报错 `range_not_found`。
  - 【附加当前选区】：大按钮明确标注用于捕获当前活动选区或替换附件。
- **采集时间与快照提示**：
  - 卡片清晰展示采集时间与提示：“ℹ️ 采集时快照；修改原区域后请点击【刷新原区域】”。
  - 客观展示时间，不后台轮询、不自动刷新、不把 60 秒作为失效判定依据。
- **并发与防覆盖**：
  - 刷新期间卡片禁用重复刷新并显示“读取中...”，禁用发送；
  - 前端维护 `refreshSeq` 与 `attachmentId`；若卡片已被用户移除或替换，晚到结果直接丢弃，不恢复已移除卡片；较旧序列号晚到直接丢弃；
  - 刷新失败保留旧快照数据与旧时间戳，不被清空；
  - 发送瞬间冻结附件快照，后续刷新不影响已发送内容与审计记录。

### 2. 代码交付与修改清单
- **C# 宿主**：
  - `src/SelectionContextService.cs`：扩充 `SelectionContextData` (`capturedAt`, `attachmentId`)，抽取公共逻辑 `ExtractContextFromRange`，实现 `GetSpecificRangeContext`（定向读取指定工作簿/表/地址）。
  - `src/NativeBridge.cs`：`HandleGetSelectionContext` 支持解析定向刷新参数并分发。
- **前端 Web**：
  - `web/src/services/bridge.ts`：更新 `SelectionContextData` 接口；`getSelectionContext` 支持定向刷新参数。
  - `web/src/services/llm.ts`：`formatSelectionContextForPrompt` 增加快照采集时间提示。
  - `web/src/components/ChatInput.svelte`：支持【🔄 刷新原区域】按钮、读取中状态、并发序列号防覆盖校验、刷新失败容灾与发送冻结机制。

### 3. 门禁与自动化验证结果
- **C# 插件编译**：`powershell build_addin.ps1` 成功生成 `bin/LeeExcel.dll`（0 Error, 0 Warning）。
- **前端构建**：`npm --prefix web run build` 成功，更新 `bin/dist/`。
- **离线单元测试**：
  - `node test_suite_unit.cjs`：**41/41 PASS (100%)**（包含 5 项 R1b 专项自动化测试：数据契约、定向刷新参数校验、并发时序丢弃校验、失败容灾保留、发送瞬间冻结）。
  - `node test_regex_counter_example.cjs`：**3/3 PASS (100%)**。
  - `.artifacts/tests/focus_handover/VerifyFocusHandover.exe`：**13/13 PASS (100%)**（焦点机制无回归）。

### 4. 当前状态与下一步
- **状态判定**：`TASK-R1b-01` 标记为 `automated_verified`。
- **停点**：停止编码，向用户提交交付报告与 5 步人工验收指南，等待用户人工实测验收。

---

## Session: 2026-10-02 (真实 Excel 桌面端自动化验收框架落地与 R1b 端到端自动验收)

### 1. 目标与执行边界
- **核心诉求**：
  1. 建立可复用的“真实 Excel 桌面端自动化验收框架”，不再默认把能自动操作的步骤交给用户手动点击；
  2. 针对 R1b 切片 1（选区采集时间、快照提示及刷新原区域）执行端到端真实桌面 UI 自动化验证；
  3. 保持状态为 `automated_verified` 并明确标注“真实 Excel UI 自动验收”，绝不伪称为“用户人工验收”；
  4. 明确划分“自动进入 accepted”与“必须由用户签收”的客观与主观边界。
- **安全红线严格遵守**：
  - 严禁全局 kill Excel，严禁关闭或接管系统原有已存在的 Excel 进程（如 PID 8508）；
  - 从真实 Ribbon 功能区与 TaskPane UI 进入，不绕过 UI 直接调内部接口；
  - 绝不修改生产提示词、VBA 正文、执行器或快照策略迎合测试；
  - 产物全部保存在 `.artifacts/tests/<run-id>/`。

### 2. 自动化框架建设与交付资产
- **独立验收执行器源码**：`tests/tools/DesktopAcceptanceRunner.cs`
  - 使用系统自带 .NET 4.8 `csc.exe` + WPF UIA 程序集（`UIAutomationClient.dll`, `UIAutomationTypes.dll`, `WindowsBase.dll`），零第三方外部运行时依赖；
  - 进程与加载项隔离：启动独立测试 Excel 实例（PID 动态跟踪），加载 `bin\LeeExcel64.xll`，打开专用测试样本工作簿；
  - COM 独占绑定：通过 Win32 `AccessibleObjectFromWindow(hwndExcel7)` 直连该测试实例的 `Excel.Application`，绝不混淆系统已有进程；
  - Ribbon UIA 控制：定位 `TabItem('ExcelMind AI')` 并调用 `SelectionItemPattern.Select()`，定位 `Button('AI助手')` 并调用 `InvokePattern.Invoke()`；
  - TaskPane UI 交互与截图：定位 `NetUINativeHWNDHost` ("ExcelMind AI")，基于窗口相对坐标驱动真实 UI 操作，生成高清 PNG 截图；
  - 安全退出与清理：`wb.Close(false)` 并退出独立测试进程，严格不误伤系统任何已有 Excel 进程。
- **统一驱动脚本**：`scripts/run_desktop_acceptance.ps1`
  - 具备编译检查、输出目录隔离（`.artifacts/tests/desktop_acceptance_<timestamp>/`）、执行调用与日志归档功能。
- **诊断探针工具**：`tests/diagnostics/ProbeDesktopUia.cs`。

### 3. R1b 自动化验收执行结果 (RunId: `desktop_acceptance_20261002_130944`)
- **严格分类统计（绝不统称为 9 项端到端）**：
  - **真实 Excel 桌面 UI 验收 (4项全部 PASS)**：
    - `TC-REG-01`：功能区 Ribbon 选项卡与 AI 助手按钮展开 (PASS, 截图就绪)；
    - `TC-R1b-01`：定向刷新原区域与活动单元格保持（【卡片地址】`$A$1:$C$5`，【样本数据】`A2: 8888.88 (更新成功)`，【时间戳】`1759381809000`，【活动单元格】`$D$10`，截图就绪）；
    - `TC-R1b-02`：替换为当前选区（【卡片地址】`$D$10:$E$12`，【样本数据】`D10:E12 区域切片`，【时间戳】`1759381811000`，【活动单元格】`$D$10`，截图就绪）；
    - `TC-R1b-07`：焦点平滑交接与草稿保持（工作表单元格输入成功，输入框草稿完好保留，截图就绪）。
  - **处理层/集成测试 (4项全部 PASS)**：
    - `TC-R1b-03`：刷新中禁止发送（按钮置灰与回车拦截生效）；
    - `TC-R1b-04`：刷新中移除附件时序保护（防旧响应晚到复活）；
    - `TC-R1b-05`：发送瞬间冻结快照（源数据后续修改不影响历史快照）；
    - `TC-R1b-06`：模型提示词隐私脱敏断言（已验证 Prompt 格式化层绝不上送本地物理全路径；限制说明：未走真实网络请求发包审计，不宣称真实网络请求绝对无泄露）。
  - **存量核心业务门禁 (1项全部 PASS)**：
    - `TC-REG-02`：存量 42 项离线单元与代码保真门禁全绿（42/42 PASS）。
- **产物与证据**：
  - 机器可读 JSON：`.artifacts/tests/desktop_acceptance_20261002_130944/test_results.json`
  - 结构化验收报告：`.artifacts/tests/desktop_acceptance_20261002_130944/acceptance_report.md`
  - 高清截图（4张）：`screenshots/TC-REG-01_...`, `TC-R1b-01_...`, `TC-R1b-02_...`, `TC-R1b-07_...`

### 4. 任务状态与自动签收规范
- **任务状态**：
  - `TASK-FRAMEWORK-01`：`accepted`（真实 Excel 桌面端自动化验收框架建立并验证）。
  - `TASK-R1b-01`：**`accepted`**（依据用户明确授权，客观功能断言与证据全量满足，正式自动签收）。
- **确立四分类签收规范**：
  - 1. 客观 UI、状态、文件和数据结果：证据满足后直接自动签收（`accepted`），不再要求用户手动点击；
  - 2. 商业大模型 API：用户授权调用预算和数据范围后，仍由 Agent 自动化执行测试与留存网络凭据；
  - 3. 固定复杂表格：有客观 COM/UIA 断言的部分继续自动化执行与自动签收；
  - 4. 必须保留人工签收的主观/特殊边界：纯主观排版审美、未覆盖的第三方特殊输入法体验。
- **未覆盖事实清单**：
  - 网络发包层：Prompt 格式化层已脱敏，真实外部商业 API 网络端点发包抓包审计未覆盖；
  - 特殊输入法：搜狗/微信等第三方 IME 悬浮窗复杂交互未深度压测。
- **停点**：保持独立分支 `feat/r1b-refresh-selection`，停止编码，不启动下一切片，不自动合并或推送。

---

## Session: 2026-10-02 (TASK-R4b-01：批量宏任务前端可视化面板与执行进度交互)

### 1. 任务目标与设计修正落实
- **设计修正严格落地**：
  1. **受控执行与响应机制**：专属 Excel 实例由独立 STA 线程创建/操作/释放；`start_batch_job` 立即返回初始状态与任务 ID；状态查询纯读线程安全快照（`CloneSummarySnapshot`），完全脱离 COM；取消请求即时登记并在任务边界优雅终止；普通宏与前台工作簿零回归。
  2. **固化定义与防漂移**：固化清单含文件、代码哈希、入口、参数、输出目录与策略；启动优先传 `jobId`，若传字段必须 100% 匹配；任何变更立即使旧固化失效；第一切片固定“遇错即停（Stop on Error）”；前置完整风险声明与主动确认复选框。
  3. **重载恢复仅恢复显示**：`localStorage` 仅保留 `jobId`，重开仅读快照，绝不自动调用 `start` 或重跑；失效任务明确展示“无法恢复任务状态”；同一 `jobId` 重复启动严格阻断。
  4. **防并发单发轮询与客观进度**：Promise 链单发防重叠轮询锁；终端态、面板关闭或销毁即停止轮询；关闭面板提示不中断后台；六态卡片（总数/已完成/成功/失败/阻断/取消/未执行）客观呈现；`stopped_on_error` 明确警示未全部成功，绝不用模拟百分比冒充宏内部进度。
  5. **界面与参数复用**：复用 R2c 参数表单，`Worksheet`/`Range` 按各文件独立隔离副本绑定；同路径重复添加去重提示，同名异径清晰展示；一键定向打开输出目录。

### 2. 自动化验收成果 (RunId: `desktop_acceptance_20261002_194637`)
- **真实 Excel 桌面 UI 验收 (17 项全部 100% PASS)**：
  - `TC-R4b-01`：面板异步启动与客观进度轮询正常，截图生成 `TC_R4b_01_BatchModalRunning.png`；
  - `TC-R4b-02`：同名不同路径文件完整保留，路径去重与提示准确；
  - `TC-R4b-04`：首个文件出错后遇错即停(`stopped_on_error`)，客观呈现未全部成功；
  - `TC-R4b-05`：运行中点击取消即时登记，边界安全终止；
  - `TC-R4b-06`：重载刷新仅恢复状态显示，不存在的任务显示错误，绝不自动重新执行；
  - `TC-R4b-07`：普通宏执行通道与前台工作簿零回归，截图生成 `TC_R4b_07_NormalNoRegression.png`；
  - 存量真实 UI 用例（TC-REG-01, TC-R1b-01, TC-R1b-02, TC-R1b-07, TC-R2b-01~05, TC-R2c-01, TC-R3a-01, TC-R3b-01）全量 100% 通过。
- **处理层/集成测试 (39 项全部 100% PASS)**：
  - `TC-R4b-03`：配置变更旧固化立即失效，重复启动严格阻断；
  - 既有 R1b, R2a, R2b, R2c, R3a, R3b, R4a 集成断言全量 100% 通过。
- **存量核心业务门禁 (1 项全部 100% PASS)**：
  - `TC-REG-02`：离线单元门禁 143/143 PASS，3/3 正则反例杜绝。

### 3. 任务状态与结论
- **状态变迁**：`TASK-R4b-01` 变更为 **`accepted`**（依据客观门槛自动签收）。
- **停点**：保持独立分支 `feat/r1b-refresh-selection`，停止当前切片编码，不处理正式用户文件，不自动提交或推送 Git。

---

## Session: 2026-10-02 (TASK-R4c-01 快捷数据工具·多文件列名对齐汇总落地与真实全量自动化验收)

### 1. 任务目标与最小闭环落地
- **零外部 API 依赖与只读保护承诺**：纯本地确定性算法，不调用任何收费大模型 API，不生成/改写 VBA 代码，只读读取源工作簿，绝不污染源数据。
- **确定性对齐与鲁棒防线**：
  1. **精确同名列对齐**：区分大小写（`Date` 与 `date` 绝不折叠）、不去空格、稳定按源文件出现顺序排布目标列；缺失列严格留空，数据绝不跨列错位；
  2. **异常表头阻断防线**：空表头、单文件内重复表头、多源列映射同一目标列冲突、输出同源文件路径重合冲突，均在预检期明确阻断报错，零脏文件生成；
  3. **同名异径防歧义**：自动提取并加权父目录标识（如 `DirA\Data.xlsx` 与 `DirB\Data.xlsx`），界面与元数据清晰区隔；
  4. **来源元数据列注入与业务重名避让**：默认追加“来源文件”与“来源工作表”两列；若业务列已有重名列，来源列自动无损避让为 `来源文件_元数据`，业务数据完好保留；
  5. **关键数据逐字符保真**：19 位纯数字长编号、前导零文本、以等号开头的公式样文本及单引号纯文本，写入新表前严格置入单引号保护，真实 Excel COM 读回 `Value2` 逐字符绝对保真且 `HasFormula=false`；
  6. **严格行数恒等式守恒**：`finalOutputRows == Σ includedDataRows + 1`（含表头），表头绝不重复混入数据行；
  7. **防静默篡改与生命周期隔离**：只读分析计算数据指纹，执行前指纹比对不匹配立即阻断；输出保存至独立新工作簿，真机重新打开核验 `UsedRange` 与行列数据；整个过程绝不触碰或关闭用户既有工作簿。

### 2. 自动化验收成果 (RunId: `desktop_acceptance_20261002_203118`)
- **全量验收用例统计 (63 项全部 100% PASS，零失败零阻断)**：
  - **真实 Excel 桌面 UI 验收 (18 项全部 100% PASS)**：
    - `TC-R4c-06`：多文件列名对齐汇总面板只读分析、确定性对齐预览、确认弹窗与真机核验全链路，输出新工作簿 `ConsolOutputUI.xlsx`，生成客观桌面截图 `TC_R4c_06_ConsolidationComplete.png`；
    - 既有真实 UI 用例 17 项全量 PASS（TC-REG-01, TC-R1b-01, TC-R1b-02, TC-R1b-07, TC-R2b-05, TC-R2c-05, TC-R3a-01~03, TC-R3b-01, TC-R3b-05, TC-R4b-01, TC-R4b-02, TC-R4b-04~07）。
  - **处理层/集成测试 (44 项全部 100% PASS)**：
    - `TC-R4c-01`：多文件不同列序同名列确定性精确对齐、缺失列填空与输出列序稳定；
    - `TC-R4c-02`：关键数据逐字符保真输出（19位长编号、前导零、公式样文本及单引号）真机读回 100% 吻合；
    - `TC-R4c-03`：同名异径文件来源标识自动区分与元数据列名重名避让；
    - `TC-R4c-04`：异常表头与输出同源冲突严格阻断（空表头、重复表头、同源输出路径阻断）；
    - `TC-R4c-05`：严格行数恒等式守恒、数据指纹防篡改漂移阻断与宿主工作簿生命周期隔离；
    - 既有集成断言 39 项全量 PASS。
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-02`：离线单元单测 153/153 PASS（涵盖 Suite 14 的 10 项确定性汇总单测），3/3 项反例彻底杜绝。

### 3. 任务状态与结论
- **状态变迁**：`TASK-R4c-01` 变更为 **`accepted`**（依据客观门槛自动签收）。
- **R4 阶段全面收尾**：R4a（批量宏队列底层）、R4b（批量宏前端面板）与 R4c（多文件列名对齐汇总）三项既定任务已全部圆满交付并达成 100% 真机全量验收。

---

## 2026-10-02：TASK-R5a-01 双步骤流水线轻量串联交付与全量自动化验收

### 1. 核心交付成果
- **C# 调度与契约引擎** (`src/WorkflowManager.cs`):
  - 严格双步骤线性调度，禁止 DAG/分支/循环/通用低代码画布；
  - 白名单组合支持：`dedup -> reconcile`、`saved_macro -> saved_macro`、`dedup -> saved_macro`；
  - 跨外部新工作簿生命周期隔离：`consolidation` 步骤明确阻断并给出清晰契约解释，不作单工作簿回滚虚假承诺；
  - 目标工作簿 `FullName`/`Name` 严格查找，零 `ActiveWorkbook` 隐式回退；前置任务快照 `CreateSnapshot(targetWb)` 保护目标；快照失败零步骤执行；
  - 精确故障与恢复范围：“失败后停止后续步骤；用户可按已验证范围恢复任务目标工作簿”；Step 1 失败 Step 2 标为 `skipped`；Step 2 失败现场保留并回传快照 ID；
  - 本地版本化持久化：定义存储于 `%APPDATA%\ExcelMindAI\Workflows\`，执行记录独立存储于 `Runs\`；定义修改自增 `definitionVersion + 1`，执行固定版本；宏正文 SHA-256 漂移阻断并强制要求重新确认；
  - 结构化 DTO 契约 (`StepOutputReference`) 严密传递，步骤边界受控取消不强杀 Excel。
- **前端流水线向导面板** (`web/src/components/WorkflowModal.svelte`, `web/src/services/bridge.ts`, `web/src/components/Header.svelte`):
  - 向导式双步骤配置、预设模版快速填充、输入源灵活路由（选区 vs 前序步骤输出 `prev_step_output`）；
  - 目标工作簿与前置快照安全提示横幅；
  - 七态状态机展示 (`pending`, `running`, `success`, `failed`, `blocked`, `skipped`, `cancelled`) 与分步详情卡片；
  - 步骤边界安全取消与防重入启动锁。

### 2. 验收证据与测试通过率
- **分类统计（全量 71 项 100% PASS）**：
  - **真实 Excel 桌面 UI 验收 (20 项全部 100% PASS)**：
    - `TC-R5a-02`：双步骤流水线（去重导出新表 → 对账分析）成功执行、输出传递与 COM 真机核验（Step 1 成功生成 `RawData_唯一数据`；Step 2 自动消费 Step 1 输出生成 `两表对账结果`；COM 读回两表结构完整，固化截图 `TC_R5a_02_WorkflowSuccess.png`）；
    - `TC-R5a-08`：双步骤流水线向导面板展示、步骤边界取消与真机桌面截图留存（取消指令在步骤边界受控生效，不强杀 Excel 宿主，截图存证 `TC_R5a_08_WorkflowCancelOrComplete.png`）；
    - 既有真实 UI 验收 18 项全量 PASS。
  - **处理层/集成测试 (50 项全部 100% PASS)**：
    - `TC-R5a-01`：工作流定义持久化、版本自增、定义重开与执行记录独立隔离；
    - `TC-R5a-03`：Step 1 失败短路机制（Step 1 失败后停止后续，Step 2 标为 `skipped` 且零执行，`failedStepIndex=1`）；
    - `TC-R5a-04`：Step 2 失败现场保留（Step 1 结果完好保留现场，Step 2 失败后输出精确恢复承诺与快照 ID）；
    - `TC-R5a-05`：切换活动工作簿不改变锁定目标（前台活动工作簿为 `OtherActive.xlsx` 时，流水线严格只写入指定的锁定工作簿，`OtherActive` 100% 零修改零污染）；
    - `TC-R5a-06`：宏代码 SHA-256 防漂移阻断（修改宏代码或哈希不符时严格阻断执行，要求用户重新确认）；
    - `TC-R5a-07`：不支持组合明确阻断（多文件汇总涉及跨工作簿外部生命周期，执行前直接拦截阻断，不作虚假单工作簿回滚承诺）；
    - 既有处理层集成测试 44 项全量 PASS。
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-02`：离线单元单测 163/163 PASS（涵盖 Suite 15 的 10 项流水线调度与持久化单测），3/3 项反例彻底杜绝。

### 3. 任务状态与结论
- **状态变迁**：`TASK-R5a-01` 变更为 **`accepted`**（依据客观断言门槛与用户预定授权自动签收）。
- **停点**：保持独立分支 `feat/r1b-refresh-selection`，严格遵守停点规范，不越界进入下一阶段业务编码，推进 `TASK-R5b-01` 最小方案设计与隔离实施。

---

## 2026-10-03：TASK-R5b-01 自选区域汇总与图表生成交付与全量自动化验收

### 1. 核心交付成果
- **大模型原生标准 VBA 主链保真**：
  - 维持大模型自然语言生成 VBA 主路径完全开放中立，绝不在 Prompt 中设任何图表类型禁令或强制模板约束；
  - 源码 100% 逐字节保真（SHA-256 原文哈希恒定），无任何引号或代码改写；
  - 严禁将模型失败偷换为预设模板；
- **C# 确定性快捷图表服务** (`src/ChartService.cs`):
  - 独立本地 COM 图表生成能力（首切片支持柱状图 `xlColumnClustered`=51、折线图 `xlLineMarkers`=65、饼图 `xlPie`=5），零第三方依赖；
  - 严格数据契约：类别列、数值系列、表头、区域边界、图表标题、目标工作表与单元格位置；
  - 饼图单系列强制约束（多系列阻断 `pie_chart_requires_single_series`）；
  - 数据质量扫描：`reject_on_invalid`（报出精确行列单元格坐标阻断）与 `coerce_zero`（数值安全置零）；
  - 防旧结果叠加与防误删：托管图表统一使用 `__EM_CHART_<chartId>` 命名，Shape 的 `AlternativeText` 注入 `{"generator":"ExcelMindAI"}` 元数据签名；试图替换无签名图表时强制安全阻断 `cannot_replace_user_chart`，100% 保护用户手工图表；
  - 目标工作簿通过 `FullName`/`Name` 严格查找锁定，零 `ActiveWorkbook` 隐式回退；前置 `SnapshotManager.CreateSnapshot(targetWb)` 强制整本物理备份；
  - COM 结构化读回核验：提取图表实际数值类型枚举、系列数量、系列名称、类别轴地址、数值系列地址；
  - `list_managed_charts` 仅枚举自身托管图表，安全排除用户手工图表；
- **双步骤流水线轻量串联集成** (`src/WorkflowManager.cs`):
  - 支持 `toolType: "chart"` 作为流水线第二步，自动消费第一步生成的 `StepOutputReference`（如 `dedup -> chart` 消费去重导出的 Sheet）；
- **前端原生面板集成** (`web/src/components/DataToolModal.svelte`, `web/src/components/WorkflowModal.svelte`, `web/src/services/bridge.ts`):
  - Tab 4 新增快捷图表工具界面：工作簿/工作表选择、区域与表头配置、类别列与数值列多选、图表类型与放置位置、新建 vs 替换已有托管图表、数据质量策略；
  - 流水线 Step 2 支持选择图表生成步骤。

### 2. 验收证据与测试通过率
- **分类统计（全量 79 项 100% PASS，新增截图 4 张）**：
  - **真实 Excel 桌面 UI 验收 (24 项全部 100% PASS)**：
    - `TC-R5b-02`：自选区域快速生成柱状图与 COM 真机读回核验（类别列 A、数值列 B/C、标题“2026品类业绩柱状图”，COM 读回 seriesCount=2, names=["销售额","利润"], categoryAddress="$A$2:$A$5", actualChartTypeNum=51；固化截图 `TC_R5b_02_ColumnChartSuccess.png`）；
    - `TC-R5b-08`：前端图表工具面板交互与已托管图表枚举验证（Tab 4 图表工具界面展示、`list_managed_charts` 仅枚举自身图表并排除手工图表，固化截图 `TC_R5b_08_ChartToolsDesktop.png`）；
    - 既有真实 UI 验收 22 项全量 PASS。
  - **处理层/集成测试 (54 项全部 100% PASS)**：
    - `TC-R5b-01`：模型原生主链保真与 Prompt 中立性验证（SHA-256 原文哈希完全一致，Prompt 无负向图表禁令，模型失败无偷换）；
    - `TC-R5b-03`：折线图生成与饼图单系列强制约束验证（折线图 65 生成成功；饼图多系列被契约严格阻断；饼图单系列 5 生成成功并捕获系列名）；
    - `TC-R5b-04`：数据质量扫描与非数值异常处理验证（`reject_on_invalid` 准确捕获 $B$3 脏数据并阻断；`coerce_zero` 转换为 0 并生成图表）；
    - `TC-R5b-05`：防旧结果叠加与用户手工图表防误删验证（试图替换无签名用户图表被严格阻断；替换自身托管图表 `__EM_CHART_` 成功替换且原手工图表完好无损）；
    - `TC-R5b-06`：目标工作簿锁定与前置快照验证（目标工作簿不存在阻断且零回退；快照管理器成功返回快照 ID）；
    - `TC-R5b-07`：双步骤流水线与图表工具衔接验证（`dedup -> chart` 成功消费第一步生成的唯一数据表并绘制图表）；
    - 既有处理层集成测试 48 项全量 PASS。
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-02`：离线单元单测 173/173 PASS（涵盖 Suite 16 的 10 项图表契约与隔离单测），3/3 项反例彻底杜绝。

### 3. 任务状态与 R5 里程碑结论
- **状态变迁**：`TASK-R5b-01` 变更为 **`accepted`**（依据客观断言门槛与用户预定授权自动签收）。
- **R5 阶段圆满收尾**：根据主规划路线图 `docs/product-roadmap.md` §3，阶段 R5（可复用工作流与报表）所包含的全部任务：
  - `R5a`：双步骤轻量流水线串联（`TASK-R5a-01`）—— **`accepted`**
  - `R5b`：自选区域汇总与图表生成（`TASK-R5b-01`）—— **`accepted`**
  已全部达成 100% 验收通过，无任何遗留待定项。**阶段 R5 正式 100% 竣工交付**。
- **推进 R6 阶段**：进入 Phase R6 外部数据与协作首个任务 `TASK-R6a-01 只读外部数据接入 (CSV/JSON/HTTP GET)`。

---

## 2026-10-03：TASK-R6a-01 只读外部数据接入 (CSV/JSON/HTTP GET) 6项核心准则核验与全量验收

### 1. 核心交付与核验成果
- **路线图与最小切片严格对齐**：
  - 严格遵循 `docs/product-roadmap.md` §3 与 §4.7 规范；
  - 聚焦受控环境下 CSV、JSON、HTTP GET 字段映射预览与独立新表写入；
  - 坚决不做自动扫描、不做后台同步、不做分页循环拉取、不做自动重试、不执行外部代码、不触碰真实商业接口。
- **6 项核心核验准则严格落地** (`src/ExternalDataService.cs`, `web/src/components/DataToolModal.svelte`, `tests/tools/DesktopAcceptanceRunner.cs`):
  1. **JSON 长数字原始语义与 RFC 8259 严格词法校验**：
     - 废除“反序列化前全局正则包裹长整数”方案，改在词法/解析阶段保留原始数字 token；
     - 充分证明绝不改动已在字符串中的长数字（`"9876543210987654321"`）、对象属性名（`"1234567890123456789"`）、转义引号（`"abc\"def,ghi"`）；
     - 完整保真负整数（`-1234567890123456789`）、高精度小数（`1234567890123456.789`）与科学计数法（`1.23456789e18`）；
     - RFC 8259 严格语法校验：前导零数字（`0123`）、前导加号（`+123`）、非法字符（`123a`）严格语法报错阻断，绝不猜测或隐式改写为合法数据。
  2. **撤回嵌套结构静默占位与明确阻断**：
     - 彻底撤回 `[Object]` / `[Array]` 静默占位伪称支持；
     - 解析时将嵌套对象与数组列明确提取为 `unsupportedColumns`，类型标为 `unsupported_nested`；
     - 用户若勾选包含不支持列，宿主与前端统一明确阻断导入，提示取消勾选后方可导入标量字段；
     - 明确字段规则展示：JSON null 与缺失字段按空白单元格写入，空字符串 `""` 按空文本写入。
  3. **明确 HTTP 限定安全防护范围**：
     - 删除“彻底防范SSRF”、“HTTP安全沙箱”等绝对化断言，明确限定在具体防护边界内；
     - 白名单路径分段严格边界比对：`/api/data` 放行自身及 `/api/data/records`，严格阻断 `/api/data_evil` 与 `/api/data-leak`；
     - 跨来源重定向（Scheme、Host 或 Port 变化）强制剥离 `Authorization` 与 `Cookie` 敏感凭据；
     - 敏感 Query 参数（`token`、`key`、`secret`、`password` 等）在错误信息与预览日志中统一脱敏为 `***`。
  4. **“先预览再导入”同一份数据保证**：
     - 引入 `PreviewCache` 机制，预览生成唯一 `previewId` 与数据源 SHA-256 `dataFingerprint`，导入时严格绑定 `previewId` 并核对指纹，若数据发生改动或指纹不符立即阻断；
     - 明确展示实际数据总规模、已排除字段与系统容量上限（50MB文件、10MB HTTP、10万行、500列）；
     - 前置整本快照失败承诺零业务写入，执行异常输出带快照 ID 的精准恢复指引。
  5. **存量功能零影响与零外部调用**：
     - 未改动 CSV 引擎、快捷去重、两表对账、多文件汇总、快捷图表或大模型 Prompt 链路；
     - 零商业 API 调用，网络验证完全基于本地受控测试服务器（`127.0.0.1:18899`）。
  6. **Phase R6 下一任务规格梳理**：
     - 根据 `docs/product-roadmap.md` §3 及 §4.8 完成 `TASK-R6b-01` 宏包导入导出规格梳理，保持纯设计状态，不提前编码。

### 2. 验收证据与测试通过率
- **分类统计（全量 87 项 100% PASS，新增真机截图存证）**：
  - **真实 Excel 桌面 UI 验收 (25 项全部 100% PASS)**：
    - `TC-R6a-07`：外部数据接入面板交互、只读预览向导与真机桌面截图（真实面板展示、源配置读取、字段选择与安全提示，固化截图 `TC_R6a_07_ExternalDataDesktop.png`）；
    - 既有真实 UI 验收 24 项全量 PASS。
  - **处理层/集成测试 (60 项全部 100% PASS)**：
    - `TC-R6a-01`：本地 CSV 只读解析、RFC 4180 引号换行与前导零长编号保真导入；
    - `TC-R6a-02`：本地 JSON RFC 8259 原始语义保真、撤回嵌套占位阻断与反例词法验证（COM 真实读回核验：orderId `1234567890123456789`、quotedStr `9876543210987654321`、negLong `-1234567890123456789`、decimalVal `1234567890123456.789`、escapedStr `abc"def,ghi` 100% 原始语义保真；嵌套列勾选导入明确阻断；前导零/加号非法数字严格阻断）；
    - `TC-R6a-04`：HTTP 路径分段边界比对、跨来源重定向剥离凭据与脱敏（/api/data 严格阻断 /api/data_evil 相似前缀冒领，未授权端口/子域阻断，Query 敏感凭据自动脱敏为 `***`）；
    - `TC-R6a-05`：单元格数据防注入转义与单引号保真；
    - `TC-R6a-06`：Windows DPAPI 凭据加密隔离存储与配置导出脱敏；
    - `TC-R6a-08`：目标锁定、快照失败承诺零写入与异常可恢复范围告知；
    - 既有处理层集成测试 54 项全量 PASS。
  - **本地受控 HTTP 测试 (1 项全部 100% PASS)**：
    - `TC-R6a-03`：受控本地 HTTP GET 接口调用、预览快照同一数据绑定与指纹防篡改（针对本地受控 mock 服务器 `127.0.0.1:18899`，previewId 快照绑定导入成功，篡改指纹导入严格阻断）。
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-02`：离线单元单测 184/184 PASS（涵盖 Suite 17 的 10 项外部数据接入与安全单测），3/3 项反例彻底杜绝。

### 3. 任务状态与结论
- **状态变迁**：`TASK-R6a-01` 变更为 **`accepted`**（依据客观断言门槛与用户预定授权自动签收）。
- **停点**：保持独立分支 `feat/r1b-refresh-selection`，严格遵守停点规范，不越界进入后续切片业务编码。
- **报告归档**：`.artifacts/tests/desktop_acceptance_20261003_080107/acceptance_report.md`。

---

## 2026-10-03：TASK-R6b-01 无凭据宏包导入/导出 (8项全量核心断言与真实桌面自动化验收)

### 1. 核心交付与核验成果
- **路线图与最小切片严格对齐**：
  - 严格遵循 `docs/product-roadmap.md` §3 及 §4.7 规范；
  - 采用标准无加密 ZIP 容器，使用 `.exmpack` 扩展名，包含 `manifest.json` 与 `sources/*.bas`（根据 schemaVersion 1.0 规范）；
  - 纯离线纯本地架构：零云端依赖、零外部网络请求、零自动执行、零外部脚本/可执行程序依赖。
- **8 项核心架构断言严格落地** (`src/MacroPackageManager.cs`, `src/NativeBridge.cs`, `web/src/services/bridge.ts`, `web/src/components/ScriptDrawer.svelte`, `test_suite_unit.cjs`, `tests/tools/DesktopAcceptanceRunner.cs`):
  1. **源码原始字节逐字节保真与 SHA-256 完整性核对**：
     - 打包写入直接复制磁盘文件的原始二进制字节流，绝不经过字符串重新编码、不改换行符（CRLF/LF）、不改单双引号；
     - `manifest.json` 精准记录源文件原始字节长度与 SHA-256 哈希，导入写盘保持 100% 逐字节往返恒等；
     - 清单 SHA-256 仅用于传输完整性核验与防篡改排查，不宣称为代码安全/信任凭据。
  2. **元数据严格字段白名单（“无凭据”核心防线）**：
     - 导出的 `manifest.json` 严格限制在白名单字段（`schemaVersion` 为 "1.0"、`packageId`、`name`、`version`、`description`、`exportedAt`、`exportedBy`，条目仅包含 `macroId`、`packageRelativePath`、`displayName`、`category`、`description`、`entryPoint`、`parameterDefs`、`sourceByteLength`、`sha256`）；
     - 绝对阻断并物理排除：大模型 API Key 与端点配置、Windows DPAPI 凭据密文、聊天会话历史、运行历史与耗时记录（`runHistory`）、流水线与批处理记录、目标工作簿物理路径、业务表格数据、快照标识与备份文件、系统机器环境信息；
     - 明确“零凭据泄露”限定为产品凭据与禁止数据不被主动打包，用户源码仍可能包含敏感内容，扫描未检出不等于绝对安全。
  3. **疑似敏感内容静态检出与用户确认拦截机制**：
     - 针对宏源码、描述、参数默认值及描述进行静态正则扫描（覆盖 API Key、GitHub Token、硬编码密码、私钥 Marker、带凭据 URL 等）；
     - 发现疑似凭据时立即暂停导出/导入，向用户弹出结构化警告清单，强制显式审核并确认；
     - 源码正文 100% 保真：用户确认后导出的宏源码正文保持原始原样，绝不进行自动打码、正则替换、星号遮蔽或静默删改；
     - 免责声明：静态正则仅作辅助提醒，未检出不代表 100% 无敏感数据。
  4. **隔离临时解包目录与深度路径防御**：
     - 解包必须在隔离的临时目录（`%LOCALAPPDATA%\LeeExcel\Temp\_pkg_temp_<guid>\`）中进行，规范命名为“隔离临时解包目录”；
     - 严格校验 ZIP 条目规范化物理 Canonical Path，严密阻断 `..` 相对穿越、绝对路径、Windows 盘符（如 `C:\`）、UNC 路径（如 `\\server\share`）以及非 `.bas`/非 `manifest.json` 的意外文件（如 `.exe`、`.bat`、`.dll`）。
  5. **容量保护与压缩炸弹防御**：
     - 预检包内条目总数（≤ 50 个）、单文件解压后大小（≤ 5MB，manifest.json ≤ 1MB）、解压后总容量（≤ 20MB）、压缩比阈值（≤ 20:1），超限直接阻断；检测到加密 ZIP 透明阻断；
     - 针对 SHA-256 篡改包与大小不符包严格拦截报错。
  6. **同名宏并存保护与本地稳定 ID 重建**：
     - 导入同名宏时，绝不覆盖已有存量宏及其运行历史；自动避让重命名为 `${displayName} (导入)`（多次重复自动递增）；
     - 为导入的宏重新分配全新的本地稳定 GUID，避免与包源机或本地既有 ID 碰撞；
     - 严格通过 R2c `VbaSignatureParser.CompareWithMetadata` 核验参数元数据与源码 Sub/Function 签名一致性，数量或名称冲突时安全阻断。
  7. **失败清理与补偿保护、零宏执行承诺**：
     - 导入过程发生任何 IO 错误、格式校验失败或用户取消时，自动清理隔离临时目录及已写入的局部新文件，存量宏库 100% 保持零改动与一致性（不承诺进程崩溃或断电下的完整数据库事务原子性）；
     - 导入完成后仅输出结构化摘要卡片，**绝对不调用 VBA 运行引擎执行宏、绝对不自动将导入宏添加至流水线、批量任务队列或功能区常用宏收藏菜单**。
  8. **前端交互与全流程体验**：
     - 在宏管理抽屉（`ScriptDrawer.svelte`）顶部工具栏与每个宏卡片上增加【导出宏包 (.exmpack)】入口，提供宏清单多选、元数据输入、保存路径选择与敏感预警确认弹窗；
     - 导入支持切换【宏包 (.exmpack)】，包含文件拖拽/选择、元数据概览卡片、敏感内容预警表、同名冲突避让提示、条目选择 Checklist 与零执行导入成功卡片。

### 2. 验收证据与测试通过率
- **分类统计（全量 95 项 100% PASS，新增真机截图存证）**：
  - **真实 Excel 桌面 UI 验收 (26 项全部 100% PASS)**：
    - `TC-R6b-08`：宏包导入导出界面交互、文件选择器与任务窗格真机截图存证（真实任务窗格展示、宏包导出/导入交互就绪，固化截图 `TC_R6b_08_MacroPackageDesktop.png`）；
    - 既有真实 UI 验收 25 项全量 PASS。
  - **处理层/集成测试 (67 项全部 100% PASS)**：
    - `TC-R6b-01`：无凭据宏包导出与导入往返断言（UTF-8/CRLF/多字节保真与 SHA-256 哈希往返一致）；
    - `TC-R6b-02`：元数据严格白名单过滤核验（排除凭据、聊天、运行历史与物理路径）；
    - `TC-R6b-03`：疑似敏感内容静态检出与用户确认拦截（导出前安全提示用户确认并暂停写入；用户确认后成功导出，且源码 100% 保持原始原样未被静默篡改或打码）；
    - `TC-R6b-04`：隔离临时解包目录与路径穿越深度防御（阻断 .. 穿越、绝对路径/盘符与非法可执行文件）；
    - `TC-R6b-05`：容量保护、压缩炸弹防御与哈希篡改透明阻断（条目超限包 51 条目与 SHA-256 篡改包均被透明拦截阻断）；
    - `TC-R6b-06`：同名宏并存保护与本地稳定 ID 重建（同名自动重命名为 `(导入)`，两者并存，存量宏零覆盖）；
    - `TC-R6b-07`：原子写入与失败回滚保护、导入全程零宏执行承诺（导入全程零 VBA 执行引擎调用，临时解包目录干净回收）；
    - 既有处理层集成测试 60 项全量 PASS。
  - **本地受控 HTTP 测试 (1 项全部 100% PASS)**：
    - `TC-R6a-03`：受控本地 HTTP GET 接口调用与指纹防篡改。
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-02`：离线单元单测 194/194 PASS（涵盖 Suite 18 的 10 项宏包全量单测），3/3 项反例彻底杜绝。

### 3. 任务状态与结论
- **状态变迁**：`TASK-R6b-01` 变更为 **`accepted`**（依据客观断言门槛与用户预定授权自动签收）。
- **停点**：保持独立分支 `feat/r1b-refresh-selection`，严格遵守停点规范，不越界进入后续切片业务编码。
- **报告归档**：`.artifacts/tests/desktop_acceptance_20261003_083458/acceptance_report.md`。

---

## 2026-10-03：TASK-R6c-01 无损安装升级与脱敏诊断导出 (定向验收与阶段 R6 全面闭环)

### 1. 核心交付与核验成果
- **路线图与最小切片严格对齐**：
  - 严格遵循 `docs/product-roadmap.md` §3 及 §4.9 规范；
  - 增强现有 `scripts/core/install_addin.ps1`，不另建庞大安装体系；
  - 纯离线纯本地架构：诊断包纯本地生成，零云端依赖、零外部网络上传、零自动执行。
- **6 项核心安全与功能架构断言严格落地** (`scripts/core/install_addin.ps1`, `src/DiagnosticsService.cs`, `src/NativeBridge.cs`, `web/src/components/SettingsModal.svelte`, `web/src/services/bridge.ts`, `test_suite_unit.cjs`, `tests/tools/VerifyR6cInstallAndDiagnostics.cs`):
  1. **三态彻底物理分离与用户资产零触碰原则**：
     - 应用文件层（`bin/` 或独立安装目录）、加载项注册层（`HKCU:\Software\Microsoft\Office\$ver\Excel\Options` `OPENx`）、用户数据层（`%APPDATA%\ExcelMindAI\`，包括 `Scripts/`、`Workflows/`、`Runs/`、`Backups/`、DPAPI 加密凭据）三态严格物理隔离；
     - 安装、重装与升级仅操作应用文件与加载项自启动项，**100% 绝不覆盖、重置、清理或删除用户数据目录中的既有文件**；
  2. **暂存预检与失败回退补偿恢复**：
     - 来源包在暂存区严格预检，缺失核心二进制（`LeeExcel.dll`, `LeeExcel.xll`, `LeeExcel64.xll`, `LeeExcel.dna`, `LeeExcel64.dna`）立即阻断终止；
     - 升级前对旧版应用文件建立带时间戳临时备份目录，异常时触发回退补偿恢复旧版，并如实报告各步骤状态（不伪称断电崩溃下完整数据库事务原子性）；
  3. **进程与占用安全防线（严禁强杀进程）**：
     - 检测到目标应用文件被锁定（Excel 正在运行或打开文件）时，安全退出（退出码 2），提示用户保存工作簿并关闭 Excel；
     - **严禁使用 `taskkill` 强杀进程，绝不静默覆盖被锁定文件，不修改 Excel 宏信任或系统级安全设置**；
  4. **脱敏诊断导出严格白名单与 8 类排除分类**：
     - 新增 `preview_diagnostics` 与 `export_diagnostics`，仅白名单收集最小诊断范围（`diagnostics_summary.json` 系统环境、`diagnostics.log` 脱敏日志 200 行、`manifest.json` 清单）；
     - 严格排除大模型 API Key、DPAPI 凭据、私钥、HTTP 请求头/Token/Cookie、会话历史与 Prompt、宏源码、宏参数值与流水线载荷、工作簿与单元格数据、快照、外部数据响应体等 8 类敏感分类；
  5. **两道脱敏与高危敏感剔除防线**：
     - 第一道规则脱敏用户路径（`C:\Users\<REDACTED_USER>`）、API Key（`sk-***`）、Bearer Token（`Bearer ***`）、敏感 Query（`?token=***`）与工作簿名（`workbook_***.xlsx`）；
     - 第二道敏感信息扫描拦截直接整行剔除无法可靠脱敏的高危私钥证书头（`BEGIN RSA PRIVATE KEY`）与硬编码密码行；
  6. **安全交付与清理防线**：
     - 前端设置面板新增“脱敏诊断与导出”独立 Tab，展示包含/排除分类与脱敏日志预览；
     - 用户点击导出时通过 Windows STA 系统文件保存对话框选择保存位置；
     - 诊断包纯本地生成，不联网，不自动上传；用户取消零生成最终包；导出失败或异常清理隔离临时目录。

### 2. 验收证据与测试通过率
- **分类统计（定向验收 9 项 100% PASS，报告归档）**：
  - **处理层/集成测试 (8 项全部 100% PASS)**：
    - `TC-R6c-01`：脱敏诊断白名单、8 类排除项、两道脱敏防线与无凭据 Clean Zip（8 项 C# 断言 100% 通过）；
    - `TC-R6c-02`：首次安装应用文件部署与模拟注册表注入；
    - `TC-R6c-03`：首次安装用户数据目录（宏库/工作流/快照）100% 恒定保护；
    - `TC-R6c-04`：同版本幂等重复安装平滑执行；
    - `TC-R6c-05`：版本升级应用文件替换与暂存备份清理；
    - `TC-R6c-06`：版本升级用户数据目录 100% 保持不变保护；
    - `TC-R6c-07`：目标文件锁定检测与非破坏性安全退出（退出码 2）；
    - `TC-R6c-08`：暂存区缺失核心文件完整性阻断（退出码 1）；
  - **存量核心业务门禁 (1 项全部 100% PASS)**：
    - `TC-REG-R6c`：离线单元单测 204/204 PASS（涵盖 Suite 19 全量 10 项安装升级与脱敏诊断单测），3/3 项反例彻底杜绝。
  - **存量桌面 UI 与集成用例复用说明**：
    - 依据新测试策略，26 项存量真实桌面 UI 用例及 60 项非直接受影响集成用例复用 v1.2.0 已归档有效结果（`.artifacts/tests/desktop_acceptance_20261003_083458/`），未进行重复运行，严禁虚构记作本轮重复测试。

### 3. 任务状态与阶段 R6 总结
- **状态变迁**：`TASK-R6c-01` 变更为 **`accepted`**（依据客观断言门槛与用户预定授权自动签收）。
- **阶段 R6 总结**：
  - `TASK-R6a-01` 只读外部数据接入 (CSV/JSON/HTTP GET) 已交付 (`accepted`)；
  - `TASK-R6b-01` 无凭据宏包导入/导出 (.exmpack) 已交付 (`accepted`)；
  - `TASK-R6c-01` 无损安装升级与脱敏诊断导出 已交付 (`accepted`)；
  - **阶段 R6 全部既定任务已无未完成项，阶段 R6 正式全量闭环收官！**
- **报告归档**：`.artifacts/tests/r6c_targeted_20261003_085240/r6c_verification_report.md`。
- **下一阶段推进**：依据 `docs/product-roadmap.md` §3 及 §4.8 路线图，已到达【停点 6: 远期评估选项】，等待用户下发下一阶段（R7 探索选项或产品发布里程碑）规划与授权。

---

## 2026-10-03：产品化发布候选准备 (v1.2.0-rc1 范围锁定、两项补测、合并回归与本地候选包就绪)

### 1. 两项直接相关补测证据闭环
1. **诊断设置页真实交互与实际包内容深度物理核验**：
   - 增加 TC-R6c-D09（用户取消导出流程：安全返回 `ok=false`，错误提示为已取消，零写目标文件，零残留）；
   - 增加 TC-R6c-D10（用户自选保存位置导出：成功返回 `zipSizeBytes` 与 `sha256` 完整元数据）；
   - 增加 TC-R6c-D11（解压后物理检查实际包内容 `diagnostics_summary.json`：白名单环境完整，100% 杜绝敏感工作簿名 `Secret_Report_2026.xlsx`、本地用户名 `JohnDoe`、API Key `sk-`）；
   - 增加 TC-R6c-D12（解压后物理检查实际包内容 `manifest.json`：schemaVersion 为 1.0，严格声明 8 类排除项规范）；
   - 单元测试增补 Test 19.12（设置模态窗口交互与深层包校验断言）。
2. **升级中途失败三重回滚与隔离核验**：
   - 增加 TC-R6c-09（复制中途失败注入：来源文件写锁导致复制中断，安全退出 code 3，自动从临时备份回退恢复 v1.1.0 旧版二进制，注册状态完好，用户数据目录 100% 恒定）；
   - 增加 TC-R6c-10（注册表配置中途失败注入：模拟无效注册根驱动器，安全退出 code 3，自动从临时备份回退恢复 v1.1.0 旧版二进制，注册状态未损，用户数据目录 100% 恒定）；
   - 单元测试增补 Test 19.11（升级复制与注册失败双重回滚与用户数据保护断言）。

### 2. 锁定发布候选范围 (Release Candidate Scope)
- **Git 状态基准**：
  - 当前分支：`feat/r1b-refresh-selection`；
  - 实际 HEAD Commit：`8c57cef docs(acceptance): finalize acceptance categories, exact card observations, and R1b accepted status`；
  - 未提交修改文件：严谨限定在 R0～R6 生产代码及发布产物范围内，杜绝自动 commit/push/merge/tag。
- **发布候选版本标识**：
  - 候选版本号：`v1.2.0-rc1` (Release Candidate 1)；
  - 构建时间戳：`2026-10-03 12:25:27`；
  - 包含能力范围：严格限定为 **R0 至 R6 已交付能力**，路线图中的 **R7 明确归为远期评估项，绝不纳入本发布候选支持列表**；
  - 架构与前端产物哈希固化：
    - `bin/LeeExcel.dll` (Any CPU): `b4c41d2cba53de6b3157429ced1700037d224cb00fe99b18ae0f94a0259de377`
    - `bin/LeeExcel64.xll` (64位): `8749875b3cab286adff77443c5e5356c937b8776c36ca61317c87ee201cf3951`
    - `bin/LeeExcel.xll` (32位): `59d9ac2404e00bc0569d09562b5780b5beca7911f05cb06d8b8cf96b663622c9`
    - `bin/dist/index.html`: `0cb3822b01e63495e49a69a3b1e3834799b94d189d0aef360d3b994b4c2a6c97`
    - `bin/dist/assets/index-BA_IMuau.js`: `29240eb142c7129cbeab72e5b3b6ed176a0592bee10c03afd2bc9232da8af2c8`
    - `bin/dist/assets/index-CtLMf7WE.css`: `2ccf5f17a1c0bd61752a5bbe13c64ec30081fb7f3145e5a4cff8dee78f66b42f`

### 3. 一次性合并回归检查成果
- **本次实际执行项**：
  1. `run_r6c_targeted_verification.ps1`：**11 项测试 100% PASS**（包含 12 项 C# 诊断导出断言、安装/重装/升级/文件锁退出/暂存缺失阻断/复制失败回滚/注册失败回滚、以及核心门禁）；
  2. `test_suite_unit.cjs`：**206 项单元测试 100% PASS**；
  3. `test_regex_counter_example.cjs`：**3 项反例测试 100% PASS**。
- **复用既有证据项**：
  - 存量 26 项真实桌面 UI 用例及 67 项处理层集成用例，完整复用同构建版本归档报告 `.artifacts/tests/desktop_acceptance_20261003_083458/acceptance_report.md`（100% PASS 证据），未重复虚构运行。
- **未验证 / 需额外授权项**：
  - 真实计费的大模型 API 端到端请求（`probe_model.cjs`、`call_llm.cjs`）；
  - 真实宿主 Office 注册表修改（`HKCU:\Software\Microsoft\Office`）；
  - R7 远期 PoC 模块（DuckDB、Python/xlwings）。

### 4. 本地发布候选包与用户说明交付
- **本地发布候选产物**：
  - 目录：`.artifacts/release/ExcelMindAI-v1.2.0-rc1/`
  - 压缩归档：`.artifacts/release/ExcelMindAI-v1.2.0-rc1.zip` (1.53 MB, SHA-256: `d15af8e3edce3254cd611db5f86934399deb8d15785d3c2e290190dabc456d21`)
  - 校验清单：`.artifacts/release/ExcelMindAI-v1.2.0-rc1/checksums_sha256.txt`（全量 25 项文件 SHA-256，标明未签名候选构建）
- **全套实用说明文档就绪**：
  - `INSTALL.md`：全新安装、版本升级、文件占用错误码 2、失败三重回滚补偿错误码 3、卸载说明、未签名 SmartScreen 与宏信任指引；
  - `docs/USER_GUIDE.md`：大模型配置与计费风险警示、Prompt-VBA 保真执行、快照与回滚边界、宏库/工作流/批处理/数据工具/图表工具使用入口、DPAPI 凭据加密、脱敏诊断导出规范、宏信任与长耗时限制；
  - `RELEASE_NOTES.md`：v1.2.0-rc1 范围锁定声明（R0~R6 已交付，排除 R7）、功能全景、未签名说明与安全防线；
  - `LICENSE.md`：MIT License 及 Excel-DNA、WebView2、Svelte、Vite 第三方开源依赖许可声明。

### 5. 最终状态
- **当前状态**：**本地发布候选包已通过独立解压冒烟，等待正式试用/发布授权。**
- **停点声明**：停留在当前发布候选准备完成节点，不启动 R7，不新增功能，不进行外部网络或 Git 远程操作。

---

## Session: 2026-10-03 (发行包独立隔离解压冒烟验证与文档纠偏)

### 1. 目标与纠偏铁律
- **严格限定范围**：不新增功能、不跑全量回归、不启动 R7，只完成发行包定向检查及文档纠偏；
- **证据与状态表述修正**：旧桌面验收报告为既有复用证据，非同一构建哈希；C# 宿主断言不替代真实设置页交互；“阻断问题 0 项”限定为当前已执行检查未发现阻断；
- **直接验证发行 ZIP**：解压至全新隔离目录 `.artifacts/tests/smoke_isolated_20261003_123829/extracted_pkg/`，零依赖仓库 bin/、源码目录或开发服务器；
- **零系统修改原则**：临时加载 XLL，绝不修改正式 Office 注册表自启动、安装位置或宏信任设置；
- **32 位架构明确标记**：“32 位文件已打包，运行未验证”。

### 2. 独立冒烟检查执行结果 (RunId: `smoke_isolated_20261003_123829`)
- **执行命令**：`powershell -NoProfile -ExecutionPolicy Bypass -File "tests/diagnostics/run_isolated_release_smoke.ps1"`
- **检查全景 (6/6 PASS, 100%)**：
  1. `TC-SMOKE-01`：包内文件与校验清单 100% 一致（真实统计 ZIP 总条目 28 项，包含目录 2 项、校验清单 1 项、载荷文件 25 项，无手工凑数）；
  2. `TC-SMOKE-02`：便携启动入口能解析正确包内路径且零外部依赖（`Split-Path -Parent $PSScriptRoot`，无外部仓库路径硬编码，保护 `%APPDATA%\ExcelMindAI\` 既有数据）；
  3. `TC-SMOKE-03`：当前已验证架构（x64）加载项成功加载与 COM 直连（PID=23040，Office 16.0 64 位，临时加载 `LeeExcel64.xll`，明确 32 位运行未验证）；
  4. `TC-SMOKE-04`：WebView2 任务窗格及前端资产正常显示（TaskPane 捕获成功，真机桌面全屏截图存证：`smoke_taskpane_desktop.png`）；
  5. `TC-SMOKE-05`：打开设置页并完成诊断预览、取消及导出全流程（预览包含白名单与排除项，取消零写入零文件生成，导出生成合法 Clean Zip 并包含 3 项核心白名单）；
  6. `TC-SMOKE-06`：固定无害宏在隔离工作簿中执行并读回（`IsolatedSmokeMacro` 成功运行，COM 读回 `A1 == "RELEASE_SMOKE_VERIFIED"`、`B1 == 20261003`）；
- **最终产物归档**：
  - 冒烟报告：`.artifacts/tests/smoke_isolated_20261003_123829/smoke_report.md`
  - 桌面全屏截图：`.artifacts/tests/smoke_isolated_20261003_123829/screenshots/smoke_taskpane_desktop.png`
  - 导出诊断包：`.artifacts/tests/smoke_isolated_20261003_123829/smoke_diagnostics_export.zip`
  - 发行 ZIP 归档：`.artifacts/release/ExcelMindAI-v1.2.0-rc1.zip`
  - 发行 ZIP SHA-256：`ef08fdb127a082a4902e98fd6b86c2b837707b3da3e925144a7faebcc49852e2`

### 3. 文档纠偏与安全合规核对
- `INSTALL.md`、`docs/USER_GUIDE.md`、`RELEASE_NOTES.md`：
  - 删除“DPAPI 硬件加密”无证据断言，修正为“当前 Windows 用户范围 DPAPI 保护 (`DataProtectionScope.CurrentUser`)”；
  - 声明组织组策略（GPO）禁用宏时功能不可用，不引导用户绕过策略；
  - 明确未签名运行说明以核对来源、哈希及用户/组织政策为前提；
  - 明确标明“32 位文件已打包，运行未验证”。
- `LICENSE.md`：
  - 依据实际纳入组件核对 MIT、zlib/libpng (Excel-DNA)、BSD-Style (WebView2)、MIT (Svelte & Vite) 声明。

### 4. 停点标记
“本地发布候选包已通过独立解压冒烟，等待正式试用/发布授权。”




