# 项目规划与任务计划书 (Task Plan) : ExcelMind AI

> **主路线图引用**：本文件执行阶段计划严格遵循唯一主规划 [docs/product-roadmap.md](file:///c:/Users/35651/Desktop/Google/lee-excle/docs/product-roadmap.md)。  
> **当前锁定版本**：v1.2.0-rc1 (Release Candidate 1, Branch: `feat/r1b-refresh-selection`, Git HEAD: `8c57cef`；注：Git HEAD 仅为初始阶段提交，R1b-2 至 R6c 全部已实现能力及最终构建固化于当前未提交工作区，等待发布授权统一提交)  
> **当前状态**：**本地发布候选包已通过独立解压冒烟，等待正式试用/发布授权。**  
> **最近更新时间**：2026-10-03  

---

## 1. 真实仓库功能矩阵与状态划分（事实审计）

| 核心功能模块 | 技术实现与代码位置 | 真实状态划分 | 依据与验收证据 |
| :--- | :--- | :--- | :--- |
| **选区只读感知 (R1a)** | `src/SelectionContextService.cs`（COM 批量子区域读取）、`ChatInput.svelte`（预览卡与选项）、`llm.ts`（低信任 Prompt 组装） | **用户人工验收**<br>*(accepted)* | 用户前台真实 Excel 运行确认通过；Node 单测全绿 |
| **选区刷新与时效提示 (R1b-1)** | `src/SelectionContextService.cs`（定向读取原工作簿/表/地址）、`ChatInput.svelte`（采集时间、刷新原区域、并发防旧覆盖）、`bridge.ts` | **自动签收**<br>*(accepted)* | 真实桌面端自动化验收通过（4项真实UI+4项集成+1项门禁），详细记录卡片地址、样本值、时间戳及活动单元格；依据用户授权自动签收 |
| **双通道显式路由** | 前端 `ChatInput.svelte`、`App.svelte` 锁定 `requestMode` (CHAT vs AUTOMATION)；后端严格执行路由隔离 | **自动测试通过**<br>(用户人工体验良好) | `test_suite_unit.cjs` (PASS: 41/41)，跨通道隔离与异步切换锁定用例 100% 验证 |
| **代码提取与截断拦截** | `web/src/services/llm.ts` 废弃脆弱正则；未闭合代码块、缺少 `End Sub` 拦截；非代码回复不报错 | **自动测试通过** | `test_suite_unit.cjs` 覆盖多段围栏、截断拦截及正文中提及 Sub 场景；无代码纯文本标注 `no_code` |
| **源码 100% 保真** | `src/ScriptManager.cs` 计算 SHA256；不替换引号、不正则改写代码；`src/VbaRunner.cs` 原文与包装器分离 | **自动测试通过** | `test_suite_unit.cjs` 验证智能引号保真；`test_regex_counter_example.cjs` 验证反例杜绝 |
| **目标工作簿绑定** | `src/NativeBridge.cs` 锁定目标 `FullName`/`Name`，`VbaRunner.cs` 在执行前后校验身份，防多窗口串改 | **已实现 (源码级)**<br>*(真实多窗口待人工验证)* | 源码包含显式工作簿查找与绑定核验；尚未在多前台 Excel 窗口频繁切换下完成用户最终签收 |
| **COM 宏注入与执行** | `src/VbaRunner.cs` 注入标准模块 `VBComponents.Add(vbext_ct_StdModule)` 并调用 `app.Run` | **自动测试通过**<br>*(外部隔离实测)* | `tests/tools/VerifyRealExcelExecution.cs` 验证无害宏执行；真实生产复杂宏仍依赖宿主单线程调度 |
| **执行错误阶段展示** | `VbaRunner.cs` 细分 `failureStage`、`vbaErrNumber`、`comHResult`；`ExecutionCard.svelte` 阶段透明展示 | **已实现 (源码级)** | 错误阶段捕获已接入前端审计卡片；复杂 VBE 弹窗需 Windows API 探测辅助 |
| **执行前全本快照恢复** | `src/SnapshotManager.cs` 在 `%APPDATA%\ExcelMindAI\Backups\` 创建物理独立副本；提供救援副本恢复 | **已实现 (源码级)**<br>*(实际工作簿待人工验收)* | 代码实现完整，支持未保存/已保存工作簿备份与恢复前救援；用户尚未进行整套异常恢复签收 |
| **宏库多源导入与管理** | `src/ScriptManager.cs` 存储于 `%APPDATA%\ExcelMindAI\Scripts\`；支持文件/粘贴导入、重命名、导出 | **已实现** | `VerifyMacroFeature.cs` 编译验证；`.bas` 与 `.meta.json` 双向持久化成功；旧数据自动迁移可见 |
| **宏库检索、标签与运行历史 (R2a)** | `src/ScriptManager.cs`、`src/NativeBridge.cs`、`web/src/components/ScriptDrawer.svelte` | **自动签收**<br>*(accepted)* | 离线单测 10/10 PASS；桌面自动化验收 6 项处理层集成断言 100% 通过（标签持久化、三态与phase、代码哈希与快照存活、元数据无损更新、写失败容灾与界面告警、blocked 耗时保护）；满足门槛自动签收 |
| **Ribbon 动态菜单展示收藏宏 (R2b)** | `src/LeeExcelRibbon.cs`、`src/LeeExcelAddIn.cs`、`src/TaskPaneControl.cs`、`web/src/components/ScriptDrawer.svelte` | **自动签收**<br>*(accepted)* | 离线核心单测 83/83 PASS；真实桌面自动化验收全绿（5项真实UI+21项集成+1项门禁），覆盖收藏持久化、动态菜单与空状态、防串选防失效、执行前确认取消零执行、正确执行快照归档、固定引用保持与目标不存在阻断 |
| **批量宏任务引擎与可视化面板 (R4a/R4b)** | `src/BatchRunnerService.cs`、`src/NativeBridge.cs`、`web/src/components/BatchModal.svelte` | **自动签收**<br>*(accepted)* | 离线单测 143/143 PASS；真实桌面端自动化验收全绿（17 项真实桌面 UI + 39 项处理层集成 + 1 项存量门禁 100% PASS），STA 受控线程调度、单发轮询、遇错即停、重载仅恢复状态且不重跑已完整验证 |
| **双步骤任务流水线 (R5a)** | `src/WorkflowManager.cs`、`src/NativeBridge.cs`、`web/src/components/WorkflowModal.svelte` | **自动签收**<br>*(accepted)* | 真实桌面自动化验收全绿（8项 R5a 用例 100% PASS），覆盖定义持久化与自增版本、执行记录物理隔离、两步全链路（去重->对账）、步骤失败短路、活动工作簿切换隔离、宏哈希防漂移、不支持组合拦截与取消响应 |
| **自选区域图表生成与快捷工具隔离 (R5b)** | `src/ChartService.cs`、`src/NativeBridge.cs`、`web/src/components/DataToolModal.svelte`（Tab 4） | **自动签收**<br>*(accepted)* | 真实桌面自动化验收全绿（8项 R5b 用例 100% PASS，总计 24 UI + 54 集成 + 1 门禁全部通过），覆盖模型原生主链保真、柱状/折线/饼图契约、饼图单系列约束、脏数据拦截、防旧结果叠加与用户手工图表防误删、前置快照恢复范围告知、双步骤衔接 |
| **只读外部数据接入 (R6a)** | `src/ExternalDataService.cs`、`src/NativeBridge.cs`、`web/src/components/DataToolModal.svelte`（Tab 5） | **自动签收**<br>*(accepted)* | 真实桌面端自动化验收全绿（8项 R6a 用例 100% PASS，总计 25 UI + 60 集成 + 1 本地HTTP + 1 门禁全部通过），覆盖 CSV RFC 4180、JSON 路径抽取、19位长编号/前导零保真、URL白名单/SSRF/重定向防护、DPAPI凭据加密隔离、公式防逃逸、目标锁定与快照安全阻断 |
| **无凭据宏包导入/导出 (R6b)** | `src/MacroPackageManager.cs`、`src/NativeBridge.cs`、`web/src/components/ScriptDrawer.svelte` | **自动签收**<br>*(accepted)* | 真实桌面端自动化验收全绿（8项 R6b 用例 100% PASS，总计 26 UI + 67 集成 + 1 本地HTTP + 1 门禁全部通过），覆盖源码原始字节逐字节保真、SHA-256完整性核对、白名单元数据脱敏、敏感内容静态拦截与用户确认、隔离临时解包目录、路径穿越与炸弹防御、同名避让与全新稳定GUID、原子回滚与零宏执行承诺 |
| **无损安装升级与脱敏诊断 (R6c)** | `scripts/core/install_addin.ps1`、`src/DiagnosticsService.cs`、`src/NativeBridge.cs`、`web/src/components/SettingsModal.svelte` | **自动签收**<br>*(accepted)* | 离线核心门禁 204 单测 + 3 反例 100% PASS；定向验收 9 项（8 项集成/处理层 + 1 项存量门禁）100% 通过（报告：`.artifacts/tests/r6c_targeted_20261003_085240/`）；三态彻底物理分离保护用户数据、文件占用安全退出、升级备份回退补偿、环境与脱敏日志两道防线、纯本地无网络 Zip 导出 |
| **功能区 (Ribbon) 布局** | `src/LeeExcelAddIn.cs` 包含智能助手、宏工具（宏库、导入、收藏宏动态菜单）、设置（API设置）三大分组 | **用户人工验收** | 2026-10-02 真实 Excel 运行确认显示正常；UI Automation 探测取证完成并保留截图 |
| **冷启动无重复注册** | `ScriptManager.cs` 等内部辅助函数显式限定为 `internal`，避免 Excel-DNA 自动导出为 Excel 工作表函数 | **自动测试通过** | 真实 Excel 加载探测证据 `ribbon_automation_evidence.txt` 验证 Excel-DNA 注册无异常 |
| **大模型 API 本地配置** | `web/src/services/config.ts` 保存在 LocalStorage (`excelmind_ai_llm_config`)，兼容旧配置迁移 | **已实现** | 纯本地存储，兼容 DeepSeek/OpenAI/Claude/通义千问/Ollama 等；Key 不离机 |
| **会话历史持久化** | `web/src/App.svelte` 当前会话记录仅存在于 Svelte 运行时内存，刷新页面重置 | **已知限制** | 标记为系统已知限制，R1a 未越权实现持久化 |

---

## 2. 当前可用版本与最小回归清单 (v1.2.0 基线)

- **可用基线 Commit**：`231ae3a` (Release v1.2.0 自动构建工作流就绪)
- **最小回归验证清单 (必须 100% 保持)**：
  1. **对话不执行**：CHAT 模式下用户任何指令均不调用 COM 宏执行器，不触发快照创建。
  2. **操作写值**：AUTOMATION 模式下有效 VBA 成功在目标工作簿写值，写后读回 `WorkbookReadback` 准确捕获区域。
  3. **无代码回复不伪报宏失败**：AUTOMATION 模式下纯文本说明回复不触发执行，不伪报宏执行错误。
  4. **导入保存重开**：导入宏持久化到 `%APPDATA%\ExcelMindAI\Scripts\`，重开 Excel 宏库列表完整可见。
  5. **源码保真**：VBA 宏引号、中文、`Option Explicit`、模块属性绝不被静默正则改写。
  6. **运行错误展示**：宏执行失败如实汇报阶段（编译期/注入期/运行期）与错误描述，不报虚假成功。
  7. **快照恢复**：执行前在 `%APPDATA%\ExcelMindAI\Backups\` 物理备份；恢复前保存救援副本。
  8. **冷启动无重复注册**：内部 C# 函数不滥用 public static，Excel-DNA 冷启动无重复函数注册异常。

---

## 3. R0 治理与 R1–R4 实施任务执行记录表

| 任务 ID | 所属阶段 | 任务内容与用户价值 | 状态 | 交付文件 / 证据 | 停点判定 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `TASK-R0-01` | R0 | **真实功能状态审计与边界划分**<br>彻底梳理代码、测试与人工验收边界，区隔 Office.js 与桌面 VBA。 | `accepted` | `task_plan.md`<br>`findings.md`<br>`progress.md` | 事实边界核实完成 |
| `TASK-R0-02` | R0 | **可用版本基线固定与最小回归清单**<br>锁定 commit `231ae3a` (v1.2.0)，确立 8 项不可退化最小回归。 | `accepted` | `docs/product-roadmap.md`<br>`task_plan.md` | 8 项回归基准确立 |
| `TASK-R0-03` | R0 | **统一主规划路线图建立**<br>以 `docs/product-roadmap.md` 为唯一主规划，统领 R0–R7，消除多份计划冲突。 | `accepted` | `docs/product-roadmap.md` | 唯一规划文件生效 |
| `TASK-R0-04` | R0 | **R1a 最小切片接口与文件映射预备**<br>梳理选区只读服务接口契约与前端组件映射，本轮不编写代码。 | `accepted` | `docs/product-roadmap.md` §8<br>`task_plan.md` §4 | 接口已定义，等用户授权 |
| `TASK-R1a-01` | R1a | **C# 宿主选区只读上下文服务**<br>独立 `SelectionContextService`，批量读取 Value2/Formula/结构，严禁整列加载，绝不逐格跨 COM。 | `accepted` | `src/SelectionContextService.cs`<br>`src/NativeBridge.cs` | 用户人工实测通过，选区只读基本功能正常 |
| `TASK-R1a-02` | R1a | **前端“附加选区”卡片与只读预览**<br>在输入区增加附加按钮、紧凑预览卡、4项发送选项与透明查看弹窗。 | `accepted` | `web/src/components/ChatInput.svelte`<br>`web/src/App.svelte` | 用户人工实测通过，选区预览正常 |
| `TASK-R1a-03` | R1a | **选区数据结构化注入 Prompt**<br>低信任数据段注入，严格防范文本单元格指令提权；请求固化与跨工作簿拦截。 | `accepted` | `web/src/services/llm.ts`<br>`web/src/services/bridge.ts` | 用户人工实测通过，请求隔离正常 |
| `TASK-FIX-FOCUS-01` | 交互缺陷 | **修复 Excel 工作表与 WebView2 任务窗格焦点交接缺陷**<br>解决点击文本框输入后再点击工作表单元格，键盘输入仍进入插件的单向锁焦问题。 | `accepted` | `src/TaskPaneFocusHelper.cs`<br>`src/TaskPaneControl.cs`<br>`src/LeeExcelAddIn.cs`<br>`tests/tools/VerifyFocusHandover.cs` | 用户实测通过：插件输入后点击单元格，新输入正确进入单元格；缺陷正式关闭；其他未实测边界保留待验 |
| `TASK-FRAMEWORK-01` | 基础设施 | **真实 Excel 桌面端自动化验收框架**<br>独立 C# 执行器与 PowerShell 驱动入口，覆盖环境检查、样本准备、启动隔离 Excel、UIA 控制 Ribbon、COM 控制工作表、任务窗格 UI 交互、截图保存与安全清理。 | `accepted` | `scripts/run_desktop_acceptance.ps1`<br>`tests/tools/DesktopAcceptanceRunner.cs`<br>`tests/diagnostics/ProbeDesktopUia.cs` | 完整建立端到端桌面自动化能力，经真实测试验证 100% 成功，系统原有 Excel 实例绝对安全保护；按证明范围正式 accepted |
| `TASK-R1b-01` | R1b 切片 1 | **选区采集时间、快照提示与刷新原区域**<br>数据契约扩充 `capturedAt` 与 `attachmentId`；定向读取原工作簿/表/地址；并发防旧覆盖校验；刷新失败容灾保留旧数据；发送瞬间冻结附件。 | `accepted`<br>*(按授权自动签收)* | `src/SelectionContextService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/ChatInput.svelte`<br>`scripts/run_desktop_acceptance.ps1`<br>`.artifacts/tests/desktop_acceptance_*/` | 真实桌面端自动化验收通过（4项真实UI+4项集成+1项门禁），详细记录卡片观测地址、样本值、时间戳及活动单元格，满足预定证据门槛，正式自动签收进入 accepted |
| `TASK-FIX-HISTORY-01` | 主链缺陷修复 | **多轮上下文模型原始回复与界面摘要分离**<br>独立保存大模型原始完整回复 `rawContent`，第二轮使用真实回复；纠正历史窗口为最近 6 条消息；排除宿主包装器污染；增加超窗宏引用警示；核实 SaveCopyAs 未保存修改基线与保真证据字节。 | **多轮原始回复上下文缺失已修复，生产历史函数离线验证通过；App.svelte 完整发送流程采用仿真验证，当前版本真实 API 多轮端到端未验证。** | `web/src/services/conversationManager.ts`<br>`web/src/App.svelte`<br>`web/src/services/llm.ts`<br>`test_suite_unit.cjs`<br>`tests/tools/VerifySaveCopyAsDirtyState.cs` | 专项结论锁定并归档结束；不将 simulateFormalAppSend 描述为实际执行了生产 handleSend，不将测试自行组装 Payload 描述为已验证真实 HTTP 发包。 |
| `TASK-R1b-02` | R1b 切片 2 | **前序宏显式引用与字符量客观透明化**<br>1. 宏卡片提供主动【引用此宏】按钮，默认不选中，单次发送有效，发送后自动清空；<br>2. 来源保真，独立取原始正文，排除包装器与执行摘要；完整/不完整/未运行宏均支持且状态透明；<br>3. 超窗宏显式引用可进入 Prompt，已在历史中的宏提供查重说明；<br>4. 准确记录各消息及引用段长度，口径明确为 UTF-16 code units (JavaScript string.length)，绝不伪称 Token 数；<br>5. 发送前置门禁阻断时不无提示丢失已选引用。 | `accepted`<br>*(按授权自动签收)* | `web/src/services/conversationManager.ts`<br>`web/src/services/llm.ts`<br>`web/src/components/ExecutionCard.svelte`<br>`web/src/components/ChatInput.svelte`<br>`web/src/App.svelte`<br>`test_suite_unit.cjs` (Suite 7)<br>`tests/tools/DesktopAcceptanceRunner.cs` | 离线待发送组装拦截 8/8 PASS，桌面端自动化集成验收 (TC-R1b-08 ~ 13) 100% 通过；依据授权自动签收；当前版本真实 API 多轮端到端未验证。 |
| `TASK-R2a-01` | R2a 切片 1 | **宏库检索、标签与运行历史溯源**<br>1. 沿用现有 ScriptManager 与 .bas + .meta.json 存储，不引入外部数据库；<br>2. 标签体系：增删、展示与筛选；增删标签仅更新元数据，.bas 源码逐字节恒定；<br>3. 组合检索：支持名称、描述、分类、标签、代码五维联合过滤及胶囊交集；<br>4. 运行历史溯源：展示最近 10 次执行记录，三态区分准确（success、failed、blocked），记录执行阶段 phase、目标工作簿、代码哈希、宏耗时；<br>5. 快照存在性物理探测与降级原因记录（snapshotExists 与 snapshotReason）；<br>6. 旧格式元数据平滑兼容，单项损坏元数据容灾保护不影响整库且不覆盖原文件。 | `accepted`<br>*(按授权自动签收)* | `src/SimpleJson.cs`<br>`src/SnapshotManager.cs`<br>`src/ScriptManager.cs`<br>`src/NativeBridge.cs`<br>`web/src/services/bridge.ts`<br>`web/src/components/ScriptDrawer.svelte`<br>`test_suite_unit.cjs` (Suite 8)<br>`tests/tools/DesktopAcceptanceRunner.cs` | 离线存储测试 10/10 PASS；全套单元门禁 70/70 PASS + 3/3 反例；真实桌面自动化验收 (TC-R2a-01 ~ 06) 6 项处理层集成断言 100% 通过（标签读写通过 Bridge 集成调用证明，不称为真实 UI 点击验收；内存未保存修改如 A1 实测进入物理快照；纯内存新建工作簿跳过备份确认为当前产品实现限制而非 Excel 原生能力结论；保留快照失败阻断与显式目标不存在阻断，不回退活动工作簿）；依据预定授权自动签收。 |
| `TASK-R2b-01` | R2b | **Excel Ribbon 动态菜单展示收藏宏与运行前确认**<br>1. 宏库提供明确收藏/取消操作（⭐），持久化保存在独立 .meta.json，源码逐字节与标签、历史保持不变；<br>2. Excel Ribbon 增加动态菜单 `menuFavoriteMacros`，展示收藏宏列表；收藏、取消、重命名或删除后通过 InvalidateRibbon 实时刷新；空列表展示明确禁用空状态；<br>3. 菜单项关联稳定 safeId，不依赖显示名称或位置，彻底防范重名/重排串选；<br>4. 点击菜单先展示宏名称、实际入口、目标工作簿与源码查看入口，用户确认后才执行；取消零执行零快照；<br>5. 确认后复用快照、目标核对、执行、错误展示及运行记录链路；宏内固定的工作簿/工作表引用不自动替换并明确提示；<br>6. 显式目标不存在时阻断且不回退活动工作簿；不可行入口明确说明原因，不猜入口、不改源码。 | `accepted`<br>*(按授权自动签收)* | `src/ScriptManager.cs`<br>`src/NativeBridge.cs`<br>`src/LeeExcelAddIn.cs`<br>`src/TaskPaneControl.cs`<br>`web/src/services/bridge.ts`<br>`web/src/App.svelte`<br>`web/src/components/ScriptDrawer.svelte`<br>`test_suite_unit.cjs` (Suite 9)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6D) | 离线核心单测 83/83 PASS + 3/3 反例杜绝；真实 Excel 桌面端自动化验收通过（5 项真实桌面 UI 验收 100% PASS + 21 项处理层集成 100% PASS + 1 项存量门禁 100% PASS，报告位于 `.artifacts/tests/desktop_acceptance_*/`）；满足预定客观断言门槛，依据授权自动签收。 |
| `TASK-R2c-01` | R2c | **显式参数化宏契约与执行隔离**<br>1. 显式支持 7 种基础参数类型，禁止隐式类型，不支持签名透明阻断；<br>2. 参数与源码签名一致性检查，防代码逃逸安全转义；<br>3. 取消与提问零快照零执行；<br>4. 独立包装器隔离执行，源码哈希恒定逐字节不变。 | `accepted`<br>*(按授权自动签收)* | `src/VbaSignatureParser.cs`<br>`src/VbaRunner.cs`<br>`src/NativeBridge.cs`<br>`test_suite_unit.cjs` (Suite 10)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6E) | 离线核心单测 91/91 PASS + 3/3 反例；真实桌面端集成用例 5 项全绿；依据客观门槛自动签收。 |
| `TASK-R3a-01` | R3a | **确定性快捷去重工具**<br>1. 只读分析与严格统计恒等式；<br>2. 选区内重复行高亮与导出新表双模式；<br>3. 前导零与长文本保真，结构化元组比对杜绝分隔符碰撞；<br>4. 强制前置快照，失败或取消零业务写入。 | `accepted`<br>*(按授权自动签收)* | `src/DataToolsService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/DataToolModal.svelte`<br>`test_suite_unit.cjs` (Suite 11)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6F) | 离线核心单测 101/101 PASS；真实桌面端验收 5 项全绿；依据客观门槛自动签收。 |
| `TASK-R3b-01` | R3b | **两表主键差异对账工具**<br>1. 确定性值比较与结构化复合主键；<br>2. 19位长编号、前导零、公式防逃逸与单引号文本全保真；<br>3. 双向非对称单侧重复键核验与双端统计恒等式严密封闭；<br>4. 独立新表输出，源表 100% 零修改。 | `accepted`<br>*(按授权自动签收)* | `src/DataToolsService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/DataToolModal.svelte`<br>`test_suite_unit.cjs` (Suite 12)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6G) | 离线核心单测 122/122 PASS；真实 Excel 桌面端导出与 Value2 读回 6 项全绿；依据客观门槛自动签收。 |
| `TASK-R4a-01` | R4a | **批量宏任务队列与隔离执行**<br>1. 显式选择清单、宏代码、参数与独立输出目录固化队列；<br>2. 单任务、单实例、严格串行调度，无并发/常驻/自动重试；<br>3. 隔离工作副本临时执行，原文件只读保真(SHA-256恒定)，外部已有工作簿零触碰；<br>4. 输出同名冲突自动递增(_1)绝不覆盖已有产物；<br>5. 宏/文件哈希漂移阻断、失败阶段定位、遇错即停(Stop on Error)与任务边界安全取消；<br>6. 显式暴露固定引用与外部副作用风险提示，绝不宣称任意宏安全沙箱。 | `accepted`<br>*(按授权自动签收)* | `src/BatchRunnerService.cs`<br>`src/NativeBridge.cs`<br>`test_suite_unit.cjs` (Suite 13)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6H)<br>`.artifacts/tests/desktop_acceptance_20261002_190824/` | 离线核心门禁 139/139 PASS + 3/3 反例；真实 Excel 桌面端验收全绿（12 项真实 UI + 37 项集成 + 1 项门禁全部 100% 通过）；满足预定客观断言门槛，依据授权自动签收。 |
| `TASK-R4b-01` | R4b | **批量宏任务前端可视化面板与执行进度交互**<br>1. 4 步向导式文件选择与同名异径清晰展示，同路径自动去重提示；<br>2. 宏库、入口与参数复用绑定（各文件隔离副本独立绑定）；<br>3. 独立输出目录，防漂移严格预检与主动风险确认勾选（默认未勾选）；<br>4. 专用 STA 线程受控调度，异步立即返回，纯快照防并发单发轮询；<br>5. 6态统计卡片客观展示，遇错即停提示未全部成功，运行中边界取消；<br>6. 重载恢复仅读快照不重跑，一键定向打开输出目录。 | `accepted`<br>*(按授权自动签收)* | `src/BatchRunnerService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/BatchModal.svelte`<br>`web/src/services/bridge.ts`<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6I)<br>`.artifacts/tests/desktop_acceptance_20261002_194637/` | 离线核心门禁 143/143 PASS + 3/3 反例；真实桌面端自动化验收全绿（17 项真实桌面 UI + 39 项处理层集成 + 1 项存量门禁全部 100% 通过）；满足预定客观断言门槛，依据授权自动签收。 |
| `TASK-R4c-01` | R4c | **快捷数据工具：多文件列名对齐汇总工具**<br>1. 只读采样多工作簿，确定性精确同名列对齐（大小写敏感，不同列序自动对齐，缺失列留空）；<br>2. 空表头/重复表头/多列映射冲突阻断，输出同源路径冲突阻断；<br>3. 自动加权父目录消除同名异径歧义，来源元数据列注入与业务重名避让（`来源文件_元数据`）；<br>4. 关键数据逐字符保真（19位长编号/前导零/公式文本）；<br>5. 数据指纹漂移阻断，严格行数恒等式守恒（输出行数 = Σ源数据行数 + 1），安全输出独立新工作簿及真机 COM 读回核验。 | `accepted`<br>*(按授权自动签收)* | `src/DataToolsService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/DataToolModal.svelte`<br>`web/src/services/bridge.ts`<br>`test_suite_unit.cjs` (Suite 14)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6J)<br>`.artifacts/tests/desktop_acceptance_20261002_203118/` | 离线核心门禁 153/153 PASS + 3/3 反例；真实 Excel 桌面端全量验收 63 项 100% PASS（18 项真实 UI + 44 项处理层集成 + 1 项核心门禁）；依据客观证据与授权自动签收。 |
| `TASK-R5a-01` | R5a | **双步骤流水线轻量串联**<br>1. 严格两步骤线性串联，无 DAG/分支/循环；优先支持“去重输出→对账”与“已保存宏→已保存宏”；不支持外部新建工作簿的汇总与活动工作簿兜底；<br>2. 预检完整定义与两步配置，锁定唯一任务目标工作簿，强制前置任务级快照；<br>3. 失败停止后续步骤（Step 1 失败 Step 2 skipped，Step 2 失败现场保留），明确恢复范围；<br>4. 本地 JSON 版本化持久化（workflowId, definitionVersion, 每步参数/源码哈希/更新时间），修改递增版本，运行时固定版本；<br>5. 结构化 DTO 输出引用传递，受控 COM 调度与边界取消，长任务界面不卡死、防重入。 | `accepted`<br>*(按授权自动签收)* | `src/WorkflowManager.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/WorkflowModal.svelte`<br>`web/src/services/bridge.ts`<br>`test_suite_unit.cjs` (Suite 15)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6K)<br>`.artifacts/tests/desktop_acceptance_20261002_213703/` | 离线核心门禁 163/163 PASS + 3/3 反例；真实 Excel 桌面端全量验收 71 项 100% PASS（20 项真实 UI + 50 项处理层集成 + 1 项核心门禁）；依据客观证据与授权自动签收。 |
| `TASK-R5b-01` | R5b | **自选区域汇总与图表生成**<br>1. 保留模型原生标准 VBA 生成主链，保持源码 100% 保真与生成自由度，不设图表负向禁令，不将生成失败偷换为预设模板；<br>2. 快捷图表工具作为明确可选路径，第一切片支持柱状图/折线图/饼图，零第三方图表库；<br>3. 严格数据契约：类别列、数值系列、表头、数据行范围、图表类型、标题、放置位置、异常数据处理、饼图单系列约束；不猜测主键或汇总方式；<br>4. 防旧结果叠加与稳定标识（`__EM_CHART_<chartId>`），新建与替换明确区分，目标缺失或身份不符阻断，绝不删除用户已有图表；<br>5. 目标工作簿严格锁定，前置整本物理快照，失败如实报告与恢复入口；COM 读回系列/轴/数值绑定核验；<br>6. 与 R5a 衔接消费 StepOutputReference，保持单任务目标快照边界。 | `accepted`<br>*(按授权自动签收)* | `src/ChartService.cs`<br>`src/WorkflowManager.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/DataToolModal.svelte`<br>`web/src/components/WorkflowModal.svelte`<br>`test_suite_unit.cjs` (Suite 16)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6L)<br>`.artifacts/tests/desktop_acceptance_20261003_065534/` | 离线核心单测 173/173 PASS；真实 Excel 桌面端全量验收 79 项 100% PASS（24 项真实 UI + 54 项处理层集成 + 1 项核心门禁）；依据客观证据与授权自动签收。 |
| `TASK-R6a-01` | R6a | **只读外部数据接入 (CSV/JSON/HTTP GET)**<br>1. **路线图引用**：遵循 `docs/product-roadmap.md` §3 及 §4.7 规范，依赖 C# 安全客户端与包结构，验收停点为受控环境下 CSV/JSON/HTTP GET 字段映射预览与独立新表写入；<br>2. **输入范围**：用户显式选择的本地 CSV、本地 JSON，以及用户显式配置并确认的受控 HTTP GET；<br>3. **读取与表格契约**：明确编码、分隔符、引号与多行；JSON 词法级解析保留数字 token，合法负数/小数/科学计数保真，RFC 8259 语法校验严格阻断前导零/加号/非数字；嵌套对象/数组撤回静默占位并明确在 `unsupportedColumns` 阻断；空值/缺失明确规则；长编号、前导零、公式样文本保真（前置单引号防公式逃逸与精度丢失）；容量超限透明阻断；<br>4. **HTTP 安全边界**：严格限制 GET 方法；白名单路径分段精确比对（阻止 `/api/data_evil` 冒领）；跨来源重定向自动剥离 Authorization 与 Cookie 敏感凭据；敏感 Query 参数（token等）自动脱敏为 `***`；明确具体防护范围而非绝对安全沙箱；测试仅限本地受控服务器；<br>5. **凭据隔离**：配置与凭据分离，凭据采用 Windows DPAPI 加密存储，不存明文 JSON，不进 VBA/工作簿/模型/导出包，日志脱敏；<br>6. **写入与恢复**：“先预览再导入”消费同一份数据快照（`previewId` 与 SHA-256 `dataFingerprint` 绑定校验），锁定目标工作簿，强制前置整本快照，新建唯一工作表（绝不覆盖已有表），取消零写入，脱敏审计记录；<br>7. **不做项**：不自动扫描、不后台同步、不分页循环拉取、不自动重试、不执行外部代码、不触碰真实商业接口。 | `accepted`<br>*(按授权自动签收)* | `src/ExternalDataService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/DataToolModal.svelte`<br>`test_suite_unit.cjs` (Suite 17)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6M)<br>`.artifacts/tests/desktop_acceptance_20261003_080107/` | 离线核心门禁 184/184 PASS + 3/3 反例；真实 Excel 桌面端全量验收 87 项 100% PASS（25 项真实 UI + 60 项处理层集成 + 1 项受控本地 HTTP + 1 项核心门禁）；依据客观证据与授权自动签收。 |
| `TASK-R6b-01` | R6b | **无凭据宏包导入/导出**<br>1. **路线图引用**：遵循 `docs/product-roadmap.md` §3 及 §4.8 规范，标准 zip 格式（.exmpack）打包 manifest 元数据与 `sources/*.bas` 源码文件，严禁泄露 API Key、会话历史与用户工作簿业务数据；<br>2. **元数据白名单与纯源码保真**：manifest 仅收录白名单元数据字段（safeId、name、description、category、tags、parameters、sourceSha256、version），剥离本地绝对路径与历史运行环境；源码字节流 100% 逐字节保真打包，绝不改写、重写或脱敏源码内容；<br>3. **敏感信息扫描与确认**：导出前自动扫描高危模式（API Key、Authorization 凭据、私钥证书、密码硬编码等）；发现敏感信息阻断导出并弹出提示；仅当用户主动勾选确认后方可继续导出无损未脱敏源码；明确“零凭据泄露”限定为产品凭据与禁止数据不被主动打包，用户源码仍可能包含敏感内容，扫描未检出不等于绝对安全；<br>4. **解压防御机制**：规范化相对路径解析，严格阻断路径穿越（`..`、绝对路径、驱动器盘符、UNC 路径、反斜杠伪造）与非 .bas 伪造载荷；Zip Bomb 深度防御（单条目上限 5MB、总解压上限 20MB、条目数上限 50、解压膨胀比上限 20:1）；<br>5. **导入防冲与失败清理补偿**：只读预览解析 manifest 与源码哈希，再次扫描敏感信息并提示用户；名称冲突自动重命名（`宏名 (导入)`）并分配全新安全随机 safeId，避免覆盖存量宏；参数契约完整校验；写入过程采用隔离临时目录解压与校验，遇到任意异常立即执行失败清理与补偿，存量宏库零污染（不承诺进程崩溃或断电下的完整数据库事务原子性）；<br>6. **零自动执行防线**：导入完成后仅以静态条目形式存入宏库，绝不自动调用、执行或静默追加至收藏夹/工作流，确保用户绝对掌控权。 | `accepted`<br>*(按授权自动签收)* | `src/MacroPackageManager.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/ScriptDrawer.svelte`<br>`web/src/services/bridge.ts`<br>`test_suite_unit.cjs` (Suite 18)<br>`tests/tools/DesktopAcceptanceRunner.cs` (阶段 6N)<br>`.artifacts/tests/desktop_acceptance_20261003_083458/` | 离线核心门禁 194/194 PASS + 3/3 反例；真实 Excel 桌面端全量验收 95 项 100% PASS（26 项真实 UI + 67 项处理层集成 + 1 项受控本地 HTTP + 1 项核心门禁）；依据客观证据与授权自动签收。 |
| `TASK-R6c-01` | R6c | **无损安装升级与脱敏诊断导出**<br>1. **路线图引用**：遵循 `docs/product-roadmap.md` §3 及 §4.9 规范，三态严格物理分离（应用文件、注册表自启动项、用户数据），100% 保护用户宏库、工作流定义、运行记录、快照与凭据；<br>2. **暂存校验与补偿恢复**：暂存区校验核心文件完整性，升级暂存备份，失败自动执行补偿恢复，报告完成与未完成步骤，不伪称绝对原子性；<br>3. **进程与占用安全防线**：检测文件占用安全退出提示关闭 Excel，绝不使用 taskkill 强杀进程，绝不静默覆盖被锁定文件，不修改宏信任设置；<br>4. **脱敏诊断导出白名单规范**：严格收集系统环境、Excel 与 WebView2 状态及限定 200 行脱敏日志；8 类敏感分类默认完全排除；两道脱敏防线（正则替换工作簿/路径/Key/Token，敏感扫描第二道剔除高危私钥/密码行）；用户指定保存路径，纯本地 Zip 压缩，取消零残留，失败清理临时目录。 | `accepted`<br>*(按授权自动签收)* | `scripts/core/install_addin.ps1`<br>`src/DiagnosticsService.cs`<br>`src/NativeBridge.cs`<br>`web/src/components/SettingsModal.svelte`<br>`web/src/services/bridge.ts`<br>`test_suite_unit.cjs` (Suite 19)<br>`tests/tools/VerifyR6cInstallAndDiagnostics.cs`<br>`tests/diagnostics/run_r6c_targeted_verification.ps1` | 离线核心门禁 204/204 PASS + 3/3 反例；定向验收 9 项（8 项集成/处理层 + 1 项存量核心门禁）100% PASS，报告位于 `.artifacts/tests/r6c_targeted_20261003_085240/`；满足预定客观断言门槛，依据授权自动签收。 |

---

## 4. 阶段验收结论、自动化成果与自动签收规范

-7. **v1.2.0-rc1 发行包独立隔离解压冒烟与文档纠偏结论 (`accepted`，按授权确认)**：
   - **执行范围与隔离机制**：将最终发行包 `ExcelMindAI-v1.2.0-rc1.zip`（SHA-256: `ef08fdb127a082a4902e98fd6b86c2b837707b3da3e925144a7faebcc49852e2`）解压至全新的隔离目录 `.artifacts/tests/smoke_isolated_20261003_123829/extracted_pkg/`，从包内入口进行纯独立检查，零依赖仓库 `bin/`、源码或开发服务器；
   - **冒烟检查 6 项 100% PASS（报告：`.artifacts/tests/smoke_isolated_20261003_123829/smoke_report.md`）**：
     1. `TC-SMOKE-01`：包内文件与校验清单 100% 一致（真实统计 ZIP 总条目 28 项，包含目录 2 项、校验清单 1 项、载荷文件 25 项，无手工凑数）；
     2. `TC-SMOKE-02`：便携启动入口（`core/launch_portable.ps1` 与 `免安装启动.bat`）正确解析包内父目录路径，无任何外部仓库路径硬编码，保证用户数据 `%APPDATA%\ExcelMindAI\` 恒定保留；
     3. `TC-SMOKE-03`：当前已验证架构（x64）加载项通过测试 Excel 临时命令行参数成功挂载加载，COM 直连正常（不修改正式注册表自启动与信任中心）；
     4. `TC-SMOKE-04`：WebView2 任务窗格与前端资产正常显示，真机桌面截图固化存证（`smoke_taskpane_desktop.png`）；
     5. `TC-SMOKE-05`：设置页真实交互与诊断预览、取消（0 写入 0 产物）及导出全流程全绿；
     6. `TC-SMOKE-06`：固定无害宏在隔离测试工作簿中成功执行并通过 COM 真实读回（`A1 == "RELEASE_SMOKE_VERIFIED"`, `B1 == 20261003`）；
   - **既有证据复用与构建说明**：旧桌面验收报告（`.artifacts/tests/desktop_acceptance_20261003_083458/`）对应构建版本为 R0～R6b 代码基线（Commit `8c57cef`）。本轮改动仅为设置页脱敏诊断真机交互流程、文档表述纠偏及校验清单条目真实统计，核心业务未改，因此既有证据复用理由充分，非“同一构建版本”；
   - **架构运行状态客观区隔**：64 位已实测通过；32 位加载项文件（`LeeExcel.xll` 等）已打包，但当前测试机无 32 位 Excel，明确记录为“32 位文件已打包，运行未验证”；
   - **“阻断问题 0 项”限定范围**：限定为当前已执行的 6 项解压冒烟检查未发现阻断，不等同于真实安装、在线 API 和所有架构均已验收；
   - **停点保持**：“本地发布候选包已通过独立解压冒烟，等待正式试用/发布授权。”

-6. **R6c 无损安装升级与脱敏诊断导出验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界与 6 项核验成果锁定**：
     1. **三态彻底物理分离**：应用二进制（`bin/`、独立安装目录）、加载项注册表自启动项（`HKCU:\Software\Microsoft\Office\$ver\Excel\Options` `OPENx`）与用户数据目录（`%APPDATA%\ExcelMindAI\`，包括 `Scripts/` 宏库、`Workflows/` 工作流定义、`Runs/` 运行记录、`Backups/` 快照及 DPAPI 加密凭据）三态严格物理隔离；安装升级脚本只操作应用文件与注册表自启动键，100% 绝不覆盖、重置或删除用户既有数据目录；
     2. **暂存校验与失败补偿恢复**：安装升级包在暂存区严格预检核心文件完整性（`LeeExcel.dll`, `LeeExcel.xll`, `LeeExcel64.xll`, `LeeExcel.dna`, `LeeExcel64.dna`），缺失核心文件阻断安装；升级前对既有应用文件建立临时备份；复制或注册异常时触发回退补偿恢复旧版，并如实输出各步骤状态（不伪称断电或崩溃下具有数据库事务绝对原子性）；
     3. **进程与占用安全防线**：检测到目标应用文件被锁定（Excel 正在运行或打开文件）时，退出并明确提示用户保存并关闭 Excel（退出码 2），严禁使用 `taskkill` 强杀进程，绝不静默覆盖被锁定文件，不修改 Excel 宏信任或系统安全设置；
     4. **脱敏诊断包白名单规范**：新增 `preview_diagnostics` 与 `export_diagnostics`，仅白名单收集系统环境（OS、架构、CLR、Excel 与 WebView2 状态、脱敏工作簿名 `workbook_***.xlsx`、最近错误阶段）与脱敏日志（限定 200 行）；严格排除 API Key、会话历史、工作簿业务数据、宏代码、快照等 8 类敏感分类；
     5. **两道脱敏与剔除防线**：第一道规则脱敏用户路径（`C:\Users\<REDACTED_USER>`）、API Key（`sk-***`）、Bearer Token（`Bearer ***`）、敏感 Query（`?token=***`）与工作簿名；第二道高危敏感扫描直接剔除无法可靠脱敏的私钥证书头（`BEGIN RSA PRIVATE KEY`）与硬编码密码行；
     6. **本地安全交付与清理**：导出前前端展示包含/排除分类与脱敏日志预览，由用户通过系统文件对话框自选保存路径；诊断包纯本地生成，不联网，不自动上传，取消不生成包，失败清理隔离临时目录。
   - **定向验收框架执行通过（9 项用例 100% PASS，报告归档于 `.artifacts/tests/r6c_targeted_20261003_085240/r6c_verification_report.md`）**：
     - `TC-R6c-01` (处理层/集成测试)：脱敏诊断白名单、8 类排除项、两道脱敏防线与无凭据 Clean Zip（8 项 C# 断言 100% 通过）；
     - `TC-R6c-02` (处理层/集成测试)：首次安装应用文件部署与模拟注册表注入；
     - `TC-R6c-03` (处理层/集成测试)：首次安装用户数据目录（宏库/工作流/快照）100% 恒定保护；
     - `TC-R6c-04` (处理层/集成测试)：同版本幂等重复安装平滑执行；
     - `TC-R6c-05` (处理层/集成测试)：版本升级应用文件替换与暂存备份清理；
     - `TC-R6c-06` (处理层/集成测试)：版本升级用户数据目录 100% 保持不变保护；
     - `TC-R6c-07` (处理层/集成测试)：目标文件锁定检测与非破坏性安全退出（退出码 2）；
     - `TC-R6c-08` (处理层/集成测试)：暂存区缺失核心文件完整性阻断（退出码 1）；
     - `TC-REG-R6c` (存量核心业务门禁)：204 项单元测试（涵盖 Suite 19 全量 10 项安装升级与诊断用例）与 3 项反例测试 100% PASS；
   - 依据新测试策略，26 项存量真实桌面 UI 用例及 60 项非直接受影响集成用例复用 v1.2.0 已归档有效结果（`.artifacts/tests/desktop_acceptance_20261003_083458/`），满足预定客观断言门槛，正式依据授权自动签收为 `accepted`。

-5. **R6b 无凭据宏包导入/导出验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界与 6 项核验成果锁定**：
     1. **无凭据标准 Zip 包格式与元数据白名单**：创建并规范 `.exmpack` 标准压缩包（遵循 schemaVersion 1.0 规范，目录采用 sources/*.bas），根目录包含版本化 `manifest.json` 与 `sources/` 纯源码目录；严格使用元数据字段白名单（`id`, `name`, `description`, `category`, `tags`, `parameters`, `sourceSha256`, `version`, `exportedAt`），彻底剥离本地文件绝对路径、历史执行状态及用户工作簿业务数据；API Key、Authorization Header、DPAPI 加密凭据 100% 物理隔离，绝不随包导出；明确“零凭据泄露”限定为产品凭据与禁止数据不被主动打包，用户源码仍可能包含业务敏感内容，扫描未检出不等于绝对安全；
     2. **源码字节流 100% 逐字节保真与 SHA-256 强校验**：导出时不修改、不混淆、不格式化、不脱敏 `.bas` 源码文件，保持源码字节流 100% 逐字节保真；每个条目计算并固化 SHA-256 哈希；导入时先验哈希，若源码字节与 manifest 中的哈希不符（如遭遇篡改或传输损坏）严格阻断导入并提示；
     3. **敏感信息扫描与用户审查闭环**：内置敏感模式扫描器（覆盖大模型 API Key `sk-[A-Za-z0-9]{20,}`、Authorization Bearer、Private Key 证书头、数据库连接串与明文密码硬编码）；导出时发现敏感内容立即暂停导出（文件系统零生成文件）并向用户弹窗告警，唯有用户主动确认（`ignoreSensitiveWarnings: true`）方可继续导出原始未掩码源码；导入端同样提供敏感信息只读预检，表格清晰罗列敏感等级、规则名称与源码摘要，供用户审查；
     4. **解压安全防御（路径穿越与 Zip Bomb）**：严格进行规范化相对路径校验，阻断包含 `..`、根目录引导、驱动器盘符（`C:`）、UNC 路径（`\\`）、反斜杠目录伪造以及非 `.bas` 的危险 payload 条目；强制实施 Zip Bomb 四维容量限制（单条目解压上限 5MB、总解压上限 20MB、条目数上限 50、解压膨胀比上限 20:1），超限立即拒绝并清理隔离临时目录；
     5. **防重名冲突、契约校验与失败清理补偿**：导入时智能探测存量宏库重名情况，自动生成安全重命名名称（`宏名称 (导入)`）并分配全局唯一且安全编码的 safeId，绝不覆盖已有宏；严格验证参数契约合法性；导入采用隔离临时目录解包与预检，执行校验成功后才持久化入库，任何步骤失败立即执行清理与补偿恢复，宏库保持零污染与一致性（不承诺掉电/崩溃下的完整数据库事务原子性）；
     6. **零自动执行安全红线**：导入完成后仅将宏元数据与源码静态写入宏库，绝不自动调用 VBA 解释器执行、不静默将新宏追加至收藏夹、不自动嵌入工作流，给用户完整的静态审查与运行控制权；
     7. **下一阶段任务规划预备**：已根据 `docs/product-roadmap.md` §3 及 §4.9 完成 `TASK-R6c-01` 无损安装升级与脱敏诊断导出规格梳理，保持纯设计状态，不提前编码。
   - 真实 Excel 桌面端自动化验收框架全量执行通过（**阶段 6N 全部 8 项 R6b 用例 100% PASS，新增真机桌面截图 1 张**）：
     - `TC-R6b-01` (处理层/集成测试)：无凭据宏包导出、元数据白名单过滤、源码逐字节保真与 SHA-256 强校验（manifest 字段精简纯洁，源码字节哈希 100% 吻合）；
     - `TC-R6b-02` (处理层/集成测试)：只读预览解析、manifest 参数契约校验与重名智能检测（预览状态为 Success，正确解析参数与重名检测）；
     - `TC-R6b-03` (处理层/集成测试)：敏感凭据扫描阻断、用户显式确认导出与导入端只读预警（检测到 `sk-live-secretkey123456789012` 时首次导出严格中断零写文件，用户确认后生成包并在导入端正确罗列警告）；
     - `TC-R6b-04` (处理层/集成测试)：恶意路径穿越包拦截与非 .bas 伪造载荷防御（包含 `../evil.bas`、绝对路径及恶意扩展名条目被全面阻断）；
     - `TC-R6b-05` (处理层/集成测试)：Zip Bomb 恶意膨胀包与容量超限防御（高压缩比炸弹与超大容量条目被严格拦截）；
     - `TC-R6b-06` (处理层/集成测试)：重名冲突自动安全重命名与失败清理补偿保障（已存在重名宏自动重命名为 `(导入)`，中间失败时清理临时文件宏库无残留）；
     - `TC-R6b-07` (处理层/集成测试)：导入后零自动执行与收藏夹隔离安全防线（导入后验证 Excel COM 零运行状态，收藏夹列表零污染）；
     - `TC-R6b-08` (真实 Excel 桌面 UI 验收)：宏包导出/导入抽屉界面交互、包文件选择与真机桌面截图（真实 UI 弹窗与抽屉交互顺畅，真机截图固化存证：`TC_R6b_08_MacroPackageDesktop.png`）；
     - `TC-REG-02` (存量核心业务门禁)：离线 194 项单元测试（涵盖 Suite 18 全量 10 项宏包用例）与 3 项反例测试 100% PASS；
   - 全量 95 项用例（26 项真实桌面 UI + 67 项处理层集成 + 1 项受控本地 HTTP + 1 项核心门禁）100% PASS，满足预定客观断言门槛，正式依据授权自动签收为 `accepted`。

-4. **R6a 只读外部数据接入 (CSV/JSON/HTTP GET) 验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界与 6 项核验成果锁定**：
     1. **JSON 长数字原始语义与严格词法校验**：摒弃反序列化前全局正则包裹长整数方案，升级为词法/解析阶段保留原始数字 token。明确原始数值 token 均以文本形态（前置单引号）写入表格，保证字符与排版无损，不等同于 Excel 后续参与公式或数值计算时的精度保证；全面覆盖并验证负整数（`-1234567890123456789`）、高精度小数（`1234567890123456.789`）、科学计数法完整 token（`1.23456789e18`，COM 读回类型为文本且全文一致）保真，且绝对不污染字符串中的长数字（`"9876543210987654321"`）、对象属性名及转义字符（`"abc\"def,ghi"`）；根据 RFC 8259 规范，针对前导零数字（`0123`）、前导加号（`+123`）、非法字符（`123a`）严格语法报错阻断，绝不隐式猜测或改写；
     2. **撤回嵌套结构静默占位与空值规则说明**：彻底撤回 `[Object]` / `[Array]` 静默占位伪称支持。包含嵌套结构的列在预览时明确在 `unsupportedColumns` 中标记为 `unsupported_nested`，若用户勾选导入该列，宿主与前端统一明确阻断，要求取消勾选后方可导入标量字段；分别说明空值规则：JSON null 写入空白单元格，缺失字段写入空白单元格以对齐列结构，空字符串 `""` 写入空文本，三者在写入 Excel 单元格后客观均为空白或空文本形态，不等同于在 Excel 中完全可逆区分；
     3. **明确 HTTP 限定安全防护范围**：删除“彻底防范SSRF”、“HTTP安全沙箱”等绝对化断言，安全结论严格限定在实际已验证的受控边界内（协议、主机、端口、路径分段精确比对），不扩大为任意域名、DNS 解析地址或复杂重定向场景全面安全；跨来源重定向（Scheme、Host 或 Port 变化）强制剥离 Authorization 与 Cookie 敏感凭据；敏感 Query 参数（`token`、`key`、`secret`、`password` 等）在错误信息与预览日志中统一脱敏为 `***`；
     4. **“先预览再导入”同一份数据保证**：引入 `PreviewCache` 机制，预览生成唯一 `previewId` 与数据源 SHA-256 `dataFingerprint`，导入时严格绑定 `previewId` 并核对指纹，若数据发生改动或指纹不符立即阻断；明确展示实际数据总规模、已排除字段与系统容量上限（50MB文件、10MB HTTP、10万行、500列）；前置整本快照失败承诺零业务写入，执行异常输出带快照 ID 的精准恢复指引；
     5. **严格不做项与存量零影响**：未改动 CSV 解析核心规则，未触碰快捷图表、宏库、工作流或大模型主链路；零真实商业 API 调用，网络测试严格限制在本地受控测试服务器（`127.0.0.1:18899`）；
     6. **下一阶段任务规划预备**：已根据 `docs/product-roadmap.md` §3 及 §4.8 完成 `TASK-R6b-01` 宏包导入导出规格梳理，保持纯设计状态，不提前编码。
   - 真实 Excel 桌面端自动化验收框架全量执行通过（**阶段 6M 全部 8 项 R6a 用例 100% PASS，新增真机桌面截图 1 张**）：
     - `TC-R6a-01` (处理层/集成测试)：本地 CSV 只读解析、RFC 4180 引号换行与前导零长编号保真导入；
     - `TC-R6a-02` (处理层/集成测试)：本地 JSON RFC 8259 原始语义保真、撤回嵌套占位阻断与反例词法验证（COM 真实读回核验通过：orderId `1234567890123456789`、quotedStr `9876543210987654321`、negLong `-1234567890123456789`、decimalVal `1234567890123456.789`、escapedStr `abc"def,ghi` 100% 保真，前导零/加号非法数字严格阻断）；
     - `TC-R6a-03` (本地受控 HTTP 测试)：受控本地 HTTP GET、预览快照同一数据绑定与指纹防篡改（本地受控测试服务器验证通过：previewId `e80c8f00054a49c98cacddd7017c8b3c` 快照绑定导入，指纹变动严格阻断）；
     - `TC-R6a-04` (处理层/集成测试)：HTTP 路径分段边界比对、跨来源重定向剥离凭据与脱敏（/api/data 阻断 /api/data_evil，未授权端口/子域阻断，Query 敏感凭据自动脱敏为 `***`）；
     - `TC-R6a-05` (处理层/集成测试)：单元格数据防注入转义与单引号保真；
     - `TC-R6a-06` (处理层/集成测试)：Windows DPAPI 凭据加密隔离存储与配置导出脱敏；
     - `TC-R6a-07` (真实 Excel 桌面 UI 验收)：外部数据接入面板交互、只读预览向导与真机桌面截图（截图存证：`TC_R6a_07_ExternalDataDesktop.png`）；
     - `TC-R6a-08` (处理层/集成测试)：目标锁定、快照失败承诺零写入与异常可恢复范围告知；
     - `TC-REG-02` (存量核心业务门禁)：离线 184 项单元测试与 3 项反例测试 100% PASS；
   - 全量 87 项用例（25 项真实桌面 UI + 60 项处理层集成 + 1 项受控本地 HTTP + 1 项核心门禁）100% PASS，满足预定客观断言门槛，正式依据授权自动签收为 `accepted`。

-3. **R5b 自选区域汇总与图表生成验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界锁定**：
     1. **双轨架构严格分离**：坚守大模型原生标准 VBA 生成为主路径（用户自选区域/Prompt → LLM 原生自主编写无约束 VBA → 源码 100% 逐字节保真 → 目标工作簿执行 → 结构化读回与审计；零负向图表禁令，零模板代码注入，模型生成失败绝不静默替罪为预设图表）；快捷图表工具为用户主动选择的独立确定性 COM 路径（首切片严格限制为柱状图、折线图、饼图，零第三方图表库）；
     2. **严格数据契约**：明确指定类别列、数值系列、表头标志、行边界、标题、放置位置与异常数据策略；严禁猜测主键或汇总方式；饼图单系列强制约束；数据质量扫描支持 `reject_on_invalid`（报出精确行列单元格坐标阻断）与 `coerce_zero`（数值安全置零）；
     3. **防旧结果叠加与防误删防线**：托管图表统一采用 `__EM_CHART_<chartId>` 稳定命名，并在 Shape 的 `AlternativeText` 中注入包含 `\"generator\":\"ExcelMindAI\"` 的 JSON 元数据签名；新建与替换模式显式区分；替换目标缺失或未携带托管签名时强制安全阻断，100% 杜绝误删用户手工报表；
     4. **目标工作簿锁定与整本物理快照**：严格按照 `FullName`/`Name` 锁定目标工作簿，严禁回退 `ActiveWorkbook`；生成前强制由 `SnapshotManager` 创建整本物理副本，生成失败或回退时如实返回快照 ID 与恢复指引；
     5. **COM 结构化读回核验**：通过只读 COM 提取图表类型（字符名与数值枚举双向核对）、系列数量、系列名称、类别轴地址、数值系列地址，确保生成的物理图表与用户契约 100% 一致；
     6. **双步骤流水线衔接**：与 R5a 无缝对接，图表步骤可直接消费前序步骤产出的 `StepOutputReference`（如 `dedup -> chart`、`reconcile -> chart`、`saved_macro -> chart`），支持用户自定义绘图列；
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6L 全部 8 项 R5b 用例 100% PASS，新增截图证据 4 张**）：
     - `TC-R5b-01` (处理层/集成测试)：模型原生主链保真与 Prompt 中立性验证（SHA-256 原文哈希完全一致，Prompt 无负向图表禁令，模型失败无偷换）；
     - `TC-R5b-02` (真实 Excel 桌面 UI 验收)：自选区域快速生成柱状图与 COM 真机读回核验（类别列 A、数值列 B/C、标题“2026品类业绩柱状图”，COM 读回 seriesCount=2, names=["销售额","利润"], categoryAddress="$A$2:$A$5", actualChartTypeNum=51；固化截图 `TC_R5b_02_ColumnChartSuccess.png`）；
     - `TC-R5b-03` (处理层/集成测试)：折线图生成与饼图单系列强制约束验证（折线图 65 生成成功；饼图多系列被契约严格阻断；饼图单系列 5 生成成功并捕获系列名）；
     - `TC-R5b-04` (处理层/集成测试)：数据质量扫描与非数值异常处理验证（`reject_on_invalid` 准确捕获 $B$3 脏数据并阻断；`coerce_zero` 转换为 0 并生成图表）；
     - `TC-R5b-05` (处理层/集成测试)：防旧结果叠加与用户手工图表防误删验证（试图替换无签名用户图表被严格阻断；替换自身托管图表 `__EM_CHART_` 成功替换且原手工图表完好无损）；
     - `TC-R5b-06` (处理层/集成测试)：目标工作簿锁定与前置快照验证（目标工作簿不存在阻断且零回退；快照管理器成功返回快照 ID）；
     - `TC-R5b-07` (处理层/集成测试)：双步骤流水线与图表工具衔接验证（`dedup -> chart` 成功消费第一步生成的唯一数据表并绘制图表）；
     - `TC-R5b-08` (真实 Excel 桌面 UI 验收)：前端图表工具面板交互与已托管图表枚举验证（Tab 4 图表工具界面展示、`list_managed_charts` 仅枚举自身图表并排除手工图表，固化截图 `TC_R5b_08_ChartToolsDesktop.png`）；
     - `TC-REG-02` (存量核心业务门禁)：离线 173 项单元测试与 3 项反例测试 100% PASS；
   - 全量 79 项用例（24 项真实桌面 UI + 54 项处理层集成 + 1 项核心门禁）100% PASS，满足预定客观断言门槛，正式依据授权自动签收为 `accepted`。

-2. **R5a 双步骤流水线轻量串联验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界锁定**：
     1. 严格线性双步骤调度（白名单支持：`dedup -> reconcile`、`saved_macro -> saved_macro`、`dedup -> saved_macro`）；
     2. 明确排除多文件汇总（`consolidation` 涉及跨外部新工作簿生命周期与独立回滚范围，首切片严格拦截阻断，零活动工作簿回滚兜底）；
     3. 任务级快照强制绑定锁定工作簿（`FullName`/`Name` 严格查找，零 `ActiveWorkbook` 隐式回退，快照失败零步骤执行）；
     4. 精确故障处置与恢复承诺：“失败后停止后续步骤；用户可按已验证范围恢复任务目标工作簿”；Step 1 失败 Step 2 标为 `skipped`；Step 2 失败保留现场并返回快照 ID；
     5. 本地 JSON 版本化持久化：定义存储于 `%APPDATA%\ExcelMindAI\Workflows\`，执行记录独立存储于 `Runs\`；定义修改版本自增 (`definitionVersion + 1`)，执行固定版本；宏代码哈希 (`macroSha256`) 漂移时严格阻断并要求用户重新确认；
     6. 结构化输出引用 (`StepOutputReference`) 严密传递，第二步仅消费第一步确认成功的产物，无 `ActiveSheet` 猜测；步骤边界安全取消不强杀 Excel。
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6K 全部 8 项 R5a 用例 100% PASS，新增截图证据 2 张**）：
     - `TC-R5a-01` (处理层/集成测试)：工作流定义持久化、版本自增、定义重开与执行记录独立隔离（初版 v1，修改自增 v2，`list_workflows` 读回，`Runs/` 目录物理隔离）；
     - `TC-R5a-02` (真实 Excel 桌面 UI 验收)：双步骤流水线（去重导出新表 → 对账分析）成功执行、输出传递与 COM 真机核验（Step 1 成功生成 `RawData_唯一数据`；Step 2 自动消费 Step 1 输出生成 `两表对账结果`；COM 读回两表结构完整，固化截图 `TC_R5a_02_WorkflowSuccess.png`）；
     - `TC-R5a-03` (处理层/集成测试)：Step 1 失败短路机制（Step 1 指向不存在表失败后立即停止后续，Step 2 严格标记为 `skipped` 且零执行，`failedStepIndex=1`）；
     - `TC-R5a-04` (处理层/集成测试)：Step 2 失败现场保留（Step 1 结果完好保留现场，Step 2 失败后输出精确恢复承诺与任务快照 ID）；
     - `TC-R5a-05` (处理层/集成测试)：切换活动工作簿不改变锁定目标（前台活动工作簿为 `OtherActive.xlsx` 时，流水线严格只写入指定的锁定工作簿，`OtherActive` 100% 零修改零污染）；
     - `TC-R5a-06` (处理层/集成测试)：宏代码 SHA-256 防漂移阻断（修改宏代码或哈希不符时严格阻断执行，要求用户在工作流中重新确认）；
     - `TC-R5a-07` (处理层/集成测试)：不支持组合明确阻断（多文件汇总涉及跨工作簿外部生命周期，执行前直接拦截阻断，不作虚假单工作簿回滚承诺）；
     - `TC-R5a-08` (真实 Excel 桌面 UI 验收)：双步骤流水线向导面板展示、步骤边界取消与真机桌面截图留存（取消指令在步骤边界受控生效，不强杀 Excel 宿主，截图存证 `TC_R5a_08_WorkflowCancelOrComplete.png`）；
     - `TC-REG-02` (存量核心业务门禁)：离线 163 项单元测试与 3 项反例测试 100% PASS；
   - 全量 71 项用例（20 项真实桌面 UI + 50 项处理层集成 + 1 项核心门禁）100% PASS，满足预定客观断言门槛，正式依据授权自动签收为 `accepted`。

-1. **R4c 快捷数据工具·多文件列名对齐汇总验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界锁定**：纯本地确定性列对齐算法、只读分析预检、数据指纹校验、同名异径加权区分、元数据避让注入、关键数据逐字符保真、独立新工作簿输出与真机 COM 重新打开只读核验；前端 Tab 3 对齐预览与确认弹窗已编译生效；
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6J 全部 6 项 R4c 用例 100% PASS，新增截图证据 1 张**）：
     - `TC-R4c-01` (处理层/集成测试)：多文件不同列序同名列确定性精确对齐、缺失列填空与输出列序稳定；参数指定 `includeMetadataCols: false`，输出 4 个纯业务列 `[日期, 姓名, 金额, 部门]`，规模 4 行 × 4 列；
     - `TC-R4c-02` (处理层/集成测试)：关键数据逐字符保真输出（19位长编号、前导零、公式样文本及单引号）真机读回 100% 吻合；
     - `TC-R4c-03` (处理层/集成测试)：同名异径文件来源标识自动区分与元数据列名重名避让；
     - `TC-R4c-04` (处理层/集成测试)：异常表头与输出同源冲突严格阻断（空表头、重复表头、同源输出路径阻断）；
     - `TC-R4c-05` (处理层/集成测试)：严格行数恒等式守恒、数据指纹防篡改漂移阻断与宿主工作簿生命周期隔离；
     - `TC-R4c-06` (真实 Excel 桌面 UI 验收)：多文件列名对齐汇总面板只读分析、确定性对齐预览、确认弹窗与真机核验完成交互全链路；参数指定 `includeMetadataCols: true`，注入 2 个来源元数据列 `[来源文件, 来源工作表]` + 4 个业务列，输出总规模 4 行 × 6 列，生成截图并固化报告；
   - 离线门禁 153/153 单元测试与 3/3 反例测试全绿，全量 63 项用例全绿，依据授权自动签收为 `accepted`。

0. **R4b 批量宏任务前端可视化面板与执行进度交互验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界锁定**：完整交付批量任务可视化面板，覆盖向导配置、预检固化、主动风险确认、异步进度展示、六态卡片、运行中取消、重载纯状态恢复及定向打开输出目录；
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6I 全部 7 项 R4b 用例 100% PASS，新增截图 2 张**）：
     - `TC-R4b-01` (真实桌面 UI 验收)：面板异步启动与客观进度轮询正常，截图记录执行进行中与状态卡片；
     - `TC-R4b-02` (真实桌面 UI 验收)：同名不同路径文件完整保留，路径去重与提示准确；
     - `TC-R4b-03` (处理层/集成测试)：配置变更旧固化立即失效，重复启动严格阻断；
     - `TC-R4b-04` (真实桌面 UI 验收)：首个文件出错后遇错即停(`stopped_on_error`)，客观呈现未全部成功；
     - `TC-R4b-05` (真实桌面 UI 验收)：运行中点击取消即时登记，边界安全终止；
     - `TC-R4b-06` (真实桌面 UI 验收)：重载刷新仅恢复状态显示，不存在的任务显示错误，绝不自动重新执行；
     - `TC-R4b-07` (真实桌面 UI 验收)：普通宏执行通道与前台工作簿零回归，批量面板操作无破坏性副作用。
   - 客观断言与证据门槛全部达成，正式自动签收为 `accepted`。

1. **R4a 批量宏任务队列与隔离执行验收结论 (`accepted`，按授权自动签收)**：
   - **交付边界锁定**：底层批量引擎、原生桥接命令与隔离调度流水线已充分验证；用户可操作的文件选择、风险确认、进度与取消前端面板（R4b）尚未完成，不宣称批量产品全部交付；
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6H 全部 7 项 R4a 用例 100% PASS**）：
     - `TC-R4a-01` (处理层/集成测试 - 宿主集成)：批量串行执行 2 个文件均处理完成(success=2)，输出文件已读回(A1=BATCH_PROCESSED, B1=999)，本次测试原文件哈希完全一致保证只读零写入；虽产生截图但实际由桥接命令驱动，诚实归类为宿主集成测试；
     - `TC-R4a-02` (处理层/集成测试)：同名防覆盖自动递增序号验证通过，原有同名文件无修改保持不变，新文件安全递增保存为 `BatchSrc1_1.xlsx`；
     - `TC-R4a-03` (处理层/集成测试)：防版本漂移与静默篡改验证通过，源文件或宏代码哈希漂移后被严格阻断(`file_hash_mismatch`/`macro_hash_mismatch`)，绝不静默使用新版本；
     - `TC-R4a-04` (处理层/集成测试)：失败阶段精确定位(`runtime_error`)与遇错即停(Stop on Error)验证通过，首项运行时除零失败后立即停止后续调度(status=`stopped_on_error`)，第 2 项保持 `pending` 未执行，不自动重试，现场证据保留未被提前删除；
     - `TC-R4a-05` (处理层/集成测试)：任务边界安全取消策略验证通过，状态置为 `cancelled`，未处理项标记 `cancelled`，Excel 进程活性完好不强杀；
     - `TC-R4a-06` (处理层/集成测试)：用户已有外部工作簿生命周期隔离保护 100% 验证通过，外部已有工作簿在批量执行期间未被关闭、保存或篡改；
     - `TC-R4a-07` (处理层/集成测试)：固定引用与系统级副作用风险声明显式暴露验证通过，如实提示固定引用风险与非沙箱边界，不向用户作出虚假安全承诺；
     - `TC-REG-02` (存量核心业务门禁)：143/143 离线单元用例全部通过，3/3 反例确认杜绝。
   - 客观断言与证据门槛全部达成，正式自动签收为 `accepted`。
   - **交付边界锁定**：底层批量引擎、原生桥接命令与隔离调度流水线已充分验证；用户可操作的文件选择、风险确认、进度与取消前端面板（R4b）尚未完成，不宣称批量产品全部交付；
   - 真实 Excel 桌面自动化验收框架全量执行通过（**阶段 6H 全部 7 项 R4a 用例 100% PASS**）：
     - `TC-R4a-01` (处理层/集成测试 - 宿主集成)：批量串行执行 2 个文件均处理完成(success=2)，输出文件已读回(A1=BATCH_PROCESSED, B1=999)，本次测试原文件哈希完全一致保证只读零写入；虽产生截图但实际由桥接命令驱动，诚实归类为宿主集成测试；
     - `TC-R4a-02` (处理层/集成测试)：同名防覆盖自动递增序号验证通过，原有同名文件无修改保持不变，新文件安全递增保存为 `BatchSrc1_1.xlsx`；
     - `TC-R4a-03` (处理层/集成测试)：防版本漂移与静默篡改验证通过，源文件或宏代码哈希漂移后被严格阻断(`file_hash_mismatch`/`macro_hash_mismatch`)，绝不静默使用新版本；
     - `TC-R4a-04` (处理层/集成测试)：失败阶段精确定位(`runtime_error`)与遇错即停(Stop on Error)验证通过，首项运行时除零失败后立即停止后续调度(status=`stopped_on_error`)，第 2 项保持 `pending` 未执行，不自动重试，现场证据保留未被提前删除；
     - `TC-R4a-05` (处理层/集成测试)：任务边界安全取消策略验证通过，状态置为 `cancelled`，未处理项标记 `cancelled`，Excel 进程活性完好不强杀；
     - `TC-R4a-06` (处理层/集成测试)：用户已有外部工作簿生命周期隔离保护 100% 验证通过，外部已有工作簿在批量执行期间未被关闭、保存或篡改；
     - `TC-R4a-07` (处理层/集成测试)：固定引用与系统级副作用风险声明显式暴露验证通过，如实提示固定引用风险与非沙箱边界，不向用户作出虚假安全承诺；
     - `TC-REG-02` (存量核心业务门禁)：143/143 离线单元用例全部通过，3/3 反例确认杜绝。
   - 客观断言与证据门槛全部达成，正式自动签收为 `accepted`。

1. **R2b Excel Ribbon 动态菜单展示收藏宏与运行前确认验收结论 (`accepted`，按授权自动签收)**：
   - 真实 Excel 桌面端自动化验收框架全量执行通过（**5 项真实桌面 UI 验收 + 21 项处理层集成测试 + 1 项存量核心门禁全部 100% PASS**）：
     - `TC-R2b-01`：收藏与取消收藏持久化（添加/取消收藏仅更新 .meta.json，.bas 源码哈希严格恒定，历史记录与标签保持不变）；
     - `TC-R2b-02`：Ribbon 动态菜单展示收藏宏与空状态（XML 安全转义与稳定 Tag 标识绑定，空收藏夹显示禁用提示按钮）；
     - `TC-R2b-03`：重名、重命名、删除后 Ribbon 动态菜单不串选、不执行失效宏（Tag safeId 精准寻址，删除后菜单同步剔除）；
     - `TC-R2b-04`：点击菜单先展示宏名称、入口与目标工作簿确认，取消零执行零快照（取消操作 0 写入、0 快照）；
     - `TC-R2b-05`：用户确认后在目标工作簿执行正确宏，快照及运行历史关联本次调用（G3 写入，物理快照存在，历史记录正常归档）；
     - `TC-R2b-06`：显式目标不存在时阻断且宏中固定引用保持原样（Sheets(1) 原样执行，G4 写入，目标不存在阻断零回退）。
   - 存量核心门禁 83/83 PASS，3/3 正则反例杜绝；全部满足客观断言门槛，正式自动签收为 `accepted`。

2. **R2a 宏库检索、标签与运行历史溯源验收结论 (`accepted`，按授权自动签收)**：
   - 真实 Excel 桌面端自动化验收框架执行通过，6 项处理层集成断言全绿：
     - `TC-R2a-01`：真实宏库标签增删与组合检索（添加/移除标签仅更新 .meta.json，.bas 源码哈希 59e2603111ae 逐字节零改动，组合筛选跨五维检索生效）；
     - `TC-R2a-02`：实际运行入口执行固定宏并记录三态与执行阶段（成功记录为 success/execution、运行时失败记录为 failed/runtime_error、目标表不存在阻断为 blocked/before_run）；
     - `TC-R2a-03`：核对运行记录关联目标工作簿、代码哈希与本次快照真实性（目标表绑定当前表、源码哈希准确、快照 ID 真实生成且物理探测有效 snapshotExists=true）；
     - `TC-R2a-04`：重命名显示名称后运行历史与标签完好保留（displayName 更新，既有 tags、runHistory、createdAt 互不覆盖）；
     - `TC-R2a-05`：模拟元数据写入失败时宏不重复执行、结果不被改写且界面明确提示（宏成功执行写值且仅执行一次 ok=true，返回体向界面输出“⚠️ 运行记录保存失败”提示）；
     - `TC-R2a-06`：核对耗时口径：blocked 阻断记录 elapsedMs 严格为 0 且前端包含 `status !== 'blocked'` 保护，绝不向用户显示为实际运行耗时。
   - **文案纠偏与证据边界锁定**：
     - 原始测试实际证明：工作簿内存编辑状态（如未保存的 A1 修改）会进入物理快照副本，并非磁盘原文件最后一次落盘状态；
     - 纯内存新建工作簿跳过备份：明确为当前产品实现限制（SnapshotManager 要求工作簿具备有效磁盘物理路径），不扩大为 Excel 原生能力结论；
     - 标签读写等通过 Bridge 调用证明的行为归为处理层/集成测试，不混淆为真实界面点击验收；
     - 保持现有快照失败阻断、显式目标未匹配时拒绝执行的安全保护，绝不恢复活动工作簿兜底。
   - 存量核心门禁 70/70 PASS，3/3 反例杜绝；全部满足客观断言门槛，保持为 `accepted`。

1. **R1a 选区功能与焦点修复验收结论**：
   - 用户已完成前台实测：插件输入后点击单元格，键盘输入正确进入单元格，焦点交接正常。
   - `TASK-R1a-01` ~ `TASK-R1a-03` 及 `TASK-FIX-FOCUS-01` 全部验收通过（`accepted`），交互缺陷正式关闭。
2. **R1b 切片 1 验收结论 (`accepted`，按授权自动签收)**：
   - 真实 Excel 桌面端自动化验收框架执行通过，测试严格拆分为三类（**严禁合并统称为 9 项真实端到端**）：
     - **真实 Excel 桌面 UI 验收 (4项全部 PASS)**：
       - `TC-REG-01`：功能区 Ribbon 选项卡与 AI助手展开（`pass`，截图：`TC-REG-01_ribbon_activated.png`）；
       - `TC-R1b-01`：定向刷新原区域与活动单元格保持（`pass`，卡片地址 `$A$1:$C$5`，A2 样本由 1200.0 更新为 8888.88，capturedAt 更新，活动单元格严格保留在 `$D$10`，截图：`TC-R1b-01_refreshed_D10_kept.png`）；
       - `TC-R1b-02`：替换为当前选区（`pass`，卡片地址更新为 `$D$10:$E$12`，呈现新切片，capturedAt 重新计时，活动单元格 `$D$10`，截图：`TC-R1b-02_attachment_replaced.png`）；
       - `TC-R1b-07`：焦点平滑交接与草稿保持（`pass`，工作表单元格输入成功，输入框草稿完好保留，截图：`TC-R1b-07_focus_handover_verified.png`）。
     - **处理层/集成测试 (4项全部 PASS)**：
       - `TC-R1b-03`：刷新中禁止发送（按钮置灰与回车拦截生效）；
       - `TC-R1b-04`：刷新中移除附件后旧响应丢弃，卡片不复活；
       - `TC-R1b-05`：发送瞬间冻结附件，后续源数据变更历史记录不变；
       - `TC-R1b-06`：模型提示词隐私脱敏断言（已验证 Prompt 格式化层绝不上送本地物理全路径；限制说明：未走真实网络请求发包审计，不宣称真实网络请求绝对无泄露）。
     - **存量核心业务门禁 (1项全部 PASS)**：
       - `TC-REG-02`：双通道显式路由、源码保真、引号保真、反例杜绝等 42/42 PASS。
3. **用户明确授权的四分类自动签收规范**：
   - **类别 1：客观 UI、状态、文件和数据结果**：自动化断言与证据门槛（截图、只读 COM、JSON 读回）满足后，**按事先授权直接自动签收（`accepted`）**，不再要求用户手动点击；
   - **类别 2：商业大模型 API 端到端测试**：在用户明确授权调用次数、费用预算上限及脱敏范围后，**仍由测试框架自动化执行并记录网络请求凭据**，不归为必须人工点击；
   - **类别 3：固定复杂表格形态**：包含合并单元格、透视表、公式阵列等表格形态，只要能定义客观 COM/UIA 断言，**继续由测试框架自动化执行与自动签收**；
   - **类别 4：必须保留由用户签收的主观/特殊边界**：
     - 复杂界面视觉呈现排版的主观审美；
     - 未覆盖的第三方特殊中文输入法（如特定拼音/五笔输入法、悬浮候选框）在极深层输入交互时的体验。
4. **未覆盖的网络与特殊环境事实清单（绝不伪称全覆盖）**：
   - **网络发包层**：已验证 Prompt 格式化层不含本地绝对路径，但**未走真实外部商业 API 网络端点发包与抓包审计**，在此不宣称真实网络请求绝对无泄露，待授权真实 API 后补充端点抓包证据。
   - **第三方输入法环境**：已验证 Win32 焦点转移与原生键盘输入，尚未在搜狗/微信输入法复杂悬浮窗下进行长流式输出深度压测。


### 4.1 最小功能闭环 (R1a)
“用户在 Excel 中框选区域 → 任务窗格点击【附加当前选区】 → 前端弹出只读结构卡片（工作表、地址、行列数、表头候选、首行样本、选项勾选） → 用户确认后随自然语言发送 → 大模型基于真实表格结构生成针对性 VBA”。

### 4.2 接口契约定义
- **桥接动作 (Action)**：`get_selection_context`
- **请求负载 (Request)**：
  ```json
  { "action": "get_selection_context", "requestId": "req_sel_xxx", "params": { "sampleRowCount": 3 } }
  ```
- **响应负载 (Response)**：
  ```json
  {
    "ok": true,
    "action": "get_selection_context",
    "requestId": "req_sel_xxx",
    "data": {
      "workbookName": "当前工作簿.xlsx",
      "sheetName": "Sheet1",
      "address": "$A$1:$F$20",
      "rowCount": 20,
      "columnCount": 6,
      "isSingleArea": true,
      "headers": ["列1", "列2", "列3"],
      "sampleRows": [["数据A", "数据B", "数据C"]],
      "hasFormulas": false,
      "hasMergedCells": false,
      "isFilteredOrHidden": false,
      "truncated": false
    }
  }
  ```

### 4.3 预定修改文件映射
1. `src/NativeBridge.cs`：添加 `get_selection_context` 消息分发；
2. `src/VbaRunner.cs`：实现 `GetSelectionContext`，通过 COM 批量读取 `range.Value2`，严禁逐格 COM 遍历；多区域友好降级；
3. `web/src/services/bridge.ts`：导出 `getSelectionContext` 调用方法；
4. `web/src/components/ChatInput.svelte`：实现【附加当前选区】按钮与紧凑只读卡片；
5. `web/src/services/llm.ts`：实现安全的 Markdown 表格提示词格式化函数（防文本单元格指令注入）。

---

## 5. 历史归档与原始记录

### 5.1 【历史归档：2026-10-01 工作区首轮保守清理与测试规范】
- 针对 123 个临时文件完成物理备份（`.backup_cleanup_batch1_20261001/`）与安全清理；
- 16 个核心测试工具迁移至 `tests/tools/` 与 `tests/diagnostics/`；
- 9 个历史证据迁移至 `docs/history/evidence_202609/`；
- 建立安全清理入口 `scripts/clean_artifacts.ps1`，默认仅预览，需 `-Apply` 执行；
- 保留 11 个活跃保护脚本于 `scratch/`（含用户修改）。

### 5.2 【历史归档：2026-09-25 系统性质量修复专项】
- 移除写死 temperature，转为前端可选配置；
- 拆分 CHAT 与 AUTOMATION 显式双通道；
- 废弃脆弱的 `includes('Sub')` 与 D1:L9 偏置词；
- 引入写后读回机制 `WorkbookReadback`；
- 单元测试与系统集成测试 100% 通过。
