# 项目规划与任务计划书 (Task Plan) : ExcelMind AI

> **主路线图引用**：本文件执行阶段计划严格遵循唯一主规划 [docs/product-roadmap.md](file:///c:/Users/35651/Desktop/Google/lee-excle/docs/product-roadmap.md)。  
> **当前锁定版本**：v1.2.0 (Git Commit: `231ae3a`)  
> **当前执行切片**：**R1a：“用户显式附加选区 → 只读预览 → 按用户选择随本次请求发送”**（本轮严格仅限 R1a，不自动进入 R1b–R7）  
> **最近更新时间**：2026-10-02  

---

## 1. 真实仓库功能矩阵与状态划分（事实审计）

根据对现有代码（`src/`、`web/src/`、`scripts/`）、配置文件以及测试资产的实际审计，严格按以下 5 种状态进行区分，**绝不把历史 Office.js 概念混淆为当前 COM/VBA 执行能力**：

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
| **功能区 (Ribbon) 布局** | `src/LeeExcelAddIn.cs` 包含智能助手、宏工具（宏库、导入）、设置（API设置）三大分组 | **用户人工验收** | 2026-10-02 真实 Excel 运行确认显示正常；UI Automation 探测取证完成并保留截图 |
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

## 3. R0 治理与 R1a 实施任务执行记录表

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

---

## 4. 阶段验收结论、自动化成果与自动签收规范

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
