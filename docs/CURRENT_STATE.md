# ExcelMind AI 当前系统状态与基线入口 (CURRENT_STATE.md)

> **文档定位**：后续开发者与 AI 识别项目真实状态的唯一当前入口  
> **更新时间**：2026-10-07  
> **维护原则**：事实驱动、基线清晰、边界分明；旧报告只证明当时测试，不作为当前假设。

---

## 一、 实际发布基线 (Release Baseline)

| 属性 | 状态 / 事实 | 核对来源与说明 |
| :--- | :--- | :--- |
| **所属仓库** | `liyongnan1688/ExcelMind-AI` | GitHub 远端权威仓库 |
| **当前发布版本** | `v1.3.0-rc1` (Pre-release Candidate 1) | 真实发布版本 |
| **构建发布提交** | `b4402a9ead90d8e74c451ccb6683a22e20da3ae4` | `git rev-parse HEAD` / `origin/main` 确认<br>*(注：历史报告曾因笔误将全量哈希末尾误记为 `...0518be4b9`，已核实修正为真实提交 `b4402a9ead90d8e74c451ccb6683a22e20da3ae4`)* |
| **Git Tag** | `v1.3.0-rc1` | Annotated Tag（Tag 对象 `22dec826dc0b46f931cc1efac6bec805e78f3f06`），其解引用 Commit 准确指向发布提交 `b4402a9` |
| **GitHub Release** | 存在发布条目 `ExcelMind AI v1.3.0-rc1` (ID: `405746132`) | **实际核实与状态纠偏**：此前由工作流写死导致的 `prerelease=false` 状态已通过 GitHub REST API 成功更正，现状态明确为 **`prerelease: true`**（Pre-release Candidate 1），`make_latest` 为 `false`；Release 说明中已如实标注官方权威校验和与已知限制。遵循铁律：不移动 tag、不删除 Release、不覆盖既有资产。 |
| **用户下载校验权威基准 (Release ZIP)** | **SHA-256**:<br>`445ef8b481491bc99b31dcd286b1ac109ae6b519fafc51fc456aaa22687ff25e` | **实际核实**：直接从 GitHub Release 官方页面下载物理包 `ExcelMindAI-v1.3.0-rc1.zip` 计算得出，与 GitHub 资产 Digest 100% 一致。此哈希为**用户下载校验的唯一权威对象**。 |
| **本地构建记录资产 (Local Build ZIP)** | **SHA-256**:<br>`f544108d73ad8cfc60723b49b2306a225a93db82e89419065fe11991bce83515` | **物理位置**：`.artifacts/release/ExcelMindAI-v1.3.0-rc1.zip`<br>本地由提交 `b4402a9` 打包流水线生成，作为另一份构建记录留存，不能与 GitHub 下载文件并列写为同一个哈希。 |

### 两个 ZIP 文件差异与冒烟证据结论

1. **文件清单完全一致**：解压比对表明，两份 ZIP 内均包含完全一致的 26 个文件。
2. **二进制依赖完全一致**：包内 6 个核心依赖 DLL（`ExcelDna.Integration.dll`、`LeeExcel.xll`、`LeeExcel64.xll`、`Microsoft.Web.WebView2.Core.dll`、`Microsoft.Web.WebView2.WinForms.dll`、`WebView2Loader.dll`）SHA-256 100% 比对一致。
3. **文本文件差异根因**：所有脚本与 Markdown（`.bat`、`.dna`、`README.md`、前端 `dist/index.html`、`dist/assets/*.js`、`dist/assets/*.css`）在去除换行符差异（GitHub CI 为 CRLF，本地为 LF）后，**正文内容 100% 完全相同**。
4. **`LeeExcel.dll` 差异根因**：两份 DLL 文件大小均为 590,336 字节。二进制逐字节对比显示，仅 44 字节存在差异（0.0075%），具体为 PE 头部编译时间戳（`TimeDateStamp`）与 .NET 4.0 编译器为程序集自动生成的随机模块版本标识（MVID GUID）。
5. **两份包均已拥有独立真机冒烟验证证据（杜绝仅凭静态推断）**：
   - **本地构建包 (`f5441...`)**：在独立隔离目录运行 6 项解压冒烟全部通过（存证报告：`.artifacts/tests/smoke_isolated_20261007_203405/smoke_report.md`）；
   - **GitHub 官方下载包 (`445ef...`)**：在独立隔离目录运行 6 项解压冒烟全部通过（存证报告：`.artifacts/tests/smoke_github_release_20261007_214500/smoke_report.md`，测试 Excel PID=8672，包内 26 项清单核验 100% 通过，WebView2 任务窗格及 Ribbon 截图存证已固化，脱敏诊断导出全流程通过，固定无害宏读回 `A1 == 'RELEASE_SMOKE_VERIFIED'`、`B1 == '20261003'`，测试进程安全释放）；
   - 本地包与 CI 包不是逐字节相同，文档分别记录构建与实际运行证据，不称为逐字节一致；两者均已在真实 64 位 Excel 宿主上获得直接端到端运行验证证据。

---

## 二、 当前 main 与发布提交的关系

- **发布基线提交**：`b4402a9ead90d8e74c451ccb6683a22e20da3ae4`（包含功能收敛与 `RELEASE_NOTES.md`，构建产物即为 `v1.3.0-rc1`）。
- **文档整理与测试资产手册提交**：`eac0f0b`（`docs(chore): reorganize project documentation, test inventory and release verification [skip ci]`，包含全仓文档梳理、历史归档与测试索引）。
- **CI 发布工作流加固提交**：`4eee7b3`（`ci: harden release workflow publication guards and prerelease detection`，彻底修复 main push fallback v1.2.0、修复 prerelease 检测并增加防止覆盖已有 Release 的前置防护）。
- **收尾文档提交**：更新当前状态与公开包冒烟证据记录。
- **发布与整理关系铁律**：已发布的 `v1.3.0-rc1` 发行包由发布提交 `b4402a9` 构建；后续 main 上的文档与 CI 工作流加固提交**不属于该包的构建源码**，绝不宣称旧发行包由最新整理后的 main 构建。

---

## 三、 未提交或未合并修改 (工作区事实)

本轮工作区对以下非源码文件保持跟踪或记录：
1. `bin/LeeExcel.dll`：本地运行打包脚本重编译产生的 PE 时间戳与 MVID 变动，源码逻辑无变动。
2. `bin/dist/*`：本地前端构建生成物检出时的换行符差异，前端源码无变动。
3. `tests/diagnostics/run_isolated_release_smoke.ps1`：增加了 `-ZipPath` 参数化路径支持，并将默认目标包更新为 `v1.3.0-rc1`。

*注：遵守 AGENTS.md 规范，不盲目执行 `git reset` 或 `git clean` 丢弃工作区文件。*

---

## 四、 核心模块概览

```
ExcelMind-AI/
├── src/                                  # C# 宿主核心 (Excel-DNA + WinForms + COM)
│   ├── LeeExcelAddIn.cs                  # 加载项入口、生命周期控制、COM 注册
│   ├── LeeExcelRibbon.cs                 # 原生功能区定义、动态常用宏菜单、大按钮布局
│   ├── TaskPaneControl.cs                # WinForms 任务窗格承载容器、WebView2 注入
│   ├── NativeBridge.cs                   # 前端-后端 JSON-RPC 桥接调度与参数校验
│   ├── VbaRunner.cs                      # 100% 源码保真宏注入、执行包装器、写后读回
│   ├── ScriptManager.cs                  # 本地宏库持久化 (%APPDATA%\ExcelMindAI\Scripts\)
│   ├── SnapshotManager.cs                # 目标工作簿物理全量快照与一键回滚
│   ├── SelectionContextService.cs        # 结构化选区动态抽样与只读感知
│   ├── BatchRunnerService.cs             # 批量宏任务队列与副本隔离试运行
│   ├── WorkflowManager.cs                # 任务流水线串联与断点恢复
│   ├── ChartService.cs                   # 原生 Excel 图表生成与独立工具隔离
│   ├── ExternalDataService.cs            # 只读外部数据接入 (CSV/JSON/HTTP GET，URL 白名单)
│   ├── MacroPackageManager.cs           # 无凭据宏包导入导出与安全扫描
│   └── DiagnosticsService.cs             # 脱敏本地诊断包导出 (排除 8 类敏感项)
├── web/                                  # 前端工作台 (Svelte 5 + TypeScript + Vite)
│   ├── src/components/                   # UI 组件 (ChatInput, ExecutionCard, ScriptDrawer, Modals)
│   └── src/services/                     # 桥接、配置、Prompt 与宏协议处理
├── tests/                                # 长期测试与诊断资产
│   ├── tools/                            # 独立编译运行的 C# 测试源码
│   └── diagnostics/                      # 诊断脚本与解压冒烟验证脚本
└── docs/                                 # 项目文档中心
    ├── product-roadmap.md                # 唯一主规划 (R0~R6 已交付，R7 远期评估)
    └── history/                          # 历史规划、旧报告与历史证据归档目录
```

---

## 五、 有效规划与测试入口

### 1. 唯一主规划
- **路径**：[docs/product-roadmap.md](file:///c:/Users/35651/Desktop/Google/lee-excle/docs/product-roadmap.md)
- **范围**：R0～R6 均已完成并在 `v1.3.0-rc1` 交付；R7（DuckDB/Python）处于远期评估状态，非当前必做项。不另建第二套路线图。

### 2. 有效测试命令与按影响范围测试策略

依据 AGENTS.md 规范，日常开发与提交严格执行**按影响范围测试策略**，绝不因仅修改 MD 文档强制执行整套测试门禁：

1. **文档与辅助说明变更 (docs-only / chore)**：
   - 仅执行文档链接与静态审查；
   - **严禁强制重新构建应用或启动全套 Excel 桌面验收**；
   - 推送到 main 时提交信息包含 `[skip ci]`。
2. **核心代码与特定模块变更**：
   - 执行离线核心门禁（无需 Excel 进程，耗时 < 1 秒）：
     ```bash
     node test_suite_unit.cjs
     node test_regex_counter_example.cjs
     ```
   - 运行与修改模块直接相关的定向测试。
3. **涉及发行包与打包流程变更**：
   - 执行独立解压隔离冒烟验证（需真实 Excel 宿主，环境完全隔离）：
     ```powershell
     pwsh -File tests/diagnostics/run_isolated_release_smoke.ps1 -ZipPath ".artifacts/tmp/ci_download_verify/ExcelMindAI-v1.3.0-rc1_from_github.zip"
     ```
4. **全套桌面端自动化验收套件 (STA 完整验收，仅在重大版本发版前明确授权运行)**：
   - 源码：`tests/tools/DesktopAcceptanceRunner.cs`
   - 手册：参见 [tests/README.md](file:///c:/Users/35651/Desktop/Google/lee-excle/tests/README.md)。

---

## 六、 准确已知限制与安全边界

1. **大模型 API Key 存储机制**：
   - 模型 Key 保存在 WebView2 实例的 `LocalStorage` 中（纯本地直连，不上报任何第三方云端）；
   - **未采用 Windows DPAPI 保护**；历史报告中曾出现将 DPAPI 外推为模型 Key 保护的表述，属于错误表述，现已更正；
   - 外部数据源凭据在启用时独立采用 Windows DPAPI (`DataProtectionScope.CurrentUser`) 本地加密存储在 `%APPDATA%\ExcelMindAI\`，两者机制不同。
2. **运行环境验证范围**：
   - **已验证**：Windows 11 64 位 + 64 位 Microsoft Excel 桌面宿主已通过全量桌面自动化验收与解压冒烟。
   - **运行未验证**：32 位 Office 宿主及真实商业大模型 API 线上调用状态为“运行未验证”（代码与资产已就绪，但未在真实 32 位环境或连接真实付费账户跑回归）。
3. **快照回滚能力边界**：
   - 快照机制在宏执行前对目标工作簿进行整本物理备份，可完全还原目标工作簿的数据与工作表结构；
   - **工作副本不是任意 VBA 的安全沙箱**：快照无法撤销宏代码执行产生的外部系统级副作用（如外部文件删除、系统命令、网络请求或第三方数据库修改）。
4. **模型代码与快捷工具分离**：
   - 助手根据用户自然语言意图生成真实标准 VBA，宿主 100% 原文保真注入执行，不静默改写；
   - 内置快捷工具（按键去重、两表对账、多文件汇总、原生图表）走独立显式执行路径，与模型生成路径严格分离。
5. **未签名运行说明**：
   - 本候选包尚未导入商业 CA 代码签名证书；首次运行时用户应先核对下载来源与包内 SHA-256 校验和，遵守所在企业 IT 安全政策，由用户自主决定运行，本向导不默认引导用户绕过系统安全防护。
6. **功能区 (Ribbon) 当前终态**：
   - 顶部选项卡为 `ExcelMind AI`；
   - 第一分组标签为 `ExcelMind AI`；
   - 主按钮仅展示 32×32 品牌大图标，下方无任何文字；鼠标悬停展示 Tooltip“打开或收起 ExcelMind AI 工作台”；
   - 历史中出现的“打开”、双品牌文本或单字竖排方案均为已替代的历史版本方案。

---

## 七、 下一步与待清理项事实

1. **当前开发状态**：当前开发周期已彻底收敛，无新增功能开发，不启动 R7。
2. **CI 与发布加固现状**：
   - `.github/workflows/release.yml` 已完成加固，main/PR 普通推送仅触发 CI 构建检查，不发布 Release；
   - 彻底删除 `v1.2.0` fallback 逻辑；发布仅接受语义化 tag 或显式指定版本的手动触发；
   - 包含 `-rc` / `-beta` 后缀自动识别并标记 Pre-release；
   - 增加 Release 覆盖检测阻断防线；离线 9 项分支/PR 场景测试全部通过。
3. **174 个草稿及其他候选文件处置**：
   - 本轮严格遵照指示保持原样，不批准物理删除，也不继续扩大逐个审计；日常生成物清理按规约由 `scripts/clean_artifacts.ps1` 在授权下受控执行。
4. **后续待授权事项**：
   - 32 位 Office 真实环境适配与兼容性验证（需在具备 32 位 Office 的测试机上授权执行）；
   - 真实商业大模型 API 线上调用审计（需用户明确授权预算与测试 Key 后方可执行）。
