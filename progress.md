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





