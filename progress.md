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











