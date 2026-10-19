# ExcelMind AI 当前系统状态与基线入口 (CURRENT_STATE.md)

> **文档定位**：后续开发者与 AI 识别项目真实状态的唯一当前入口  
> **更新时间**：2026-10-19  
> **维护原则**：事实驱动、基线清晰、边界分明；旧报告只证明当时测试，不作为当前假设。

---

## 一、 实际发布基线 (Release Baseline)

| 属性 | 状态 / 事实 | 核对来源与说明 |
| :--- | :--- | :--- |
| **所属仓库** | `liyongnan1688/ExcelMind-AI` | GitHub 远端权威仓库 |
| **当前发布版本** | `v1.3.0` (Official Production Release) | 正式稳定发行版本 |
| **前序候选版本** | `v1.3.0-rc1` (Pre-release Candidate 1) | 历史候选发布（Tag `v1.3.0-rc1`，Commit `b4402a9`） |
| **Git Tag** | `v1.3.0` | 正式发行 Tag |
| **本地构建资产 (Local Release ZIP)** | **SHA-256**:<br>`2dd5be642a3b78ab67001e6f0206f2a80e579bc8de50cacd898206f915eeeb3e` | **物理位置**：`release/ExcelMindAI-v1.3.0.zip`<br>由审计整改后最新代码打包生成（1.57 MB） |

### 独立隔离解压冒烟与五大真机验收确证

1. **发行包解压冒烟验证（“能不能跑” - 6/6 PASS）**：
   - **存证报告**：`.artifacts/tests/smoke_isolated_20261019_150904/smoke_report.md`；
   - **动态标题**：`# ExcelMind AI v1.3.0 发行包独立隔离解压冒烟验证报告`（已彻底修复历史报告标题写死版本号的技术债）；
   - **结果**：清单核验（26项）、便携入口、64位加载、WebView2窗格展示与Ribbon截图、脱敏诊断导出全流程、固定无害宏读回全量通过，宿主进程安全退出。
2. **五大验收场景矩阵（“跑得对不对” - 10 项真机验证 100% PASS + 16 项模块/规范复核 15 PASS / 1 NOT TESTED）**：
   - **存证报告**：`.artifacts/acceptance/20261019/acceptance_summary_report.md`；
   - **场景 1 (Track A 只读工具真机验证 - 5/5 PASS)**：毫秒级全簿概览（实测 218 ms，杜绝修饰值）、跨表引用血缘分析、循环引用探测、空单元格受控、关闭容错全部通过；
   - **场景 2 (Track B 读回契约真机验证 - 5/5 PASS)**：字面匹配 `matched`、排版样式诚实标为 `unconfirmed`（杜绝虚假抽样通过徽章）、写脏数据报错与快照物理回滚、跨簿越权附带修改精准检出（`otherWorkbooksAffected = true`）并严格阻断存宏、目标错乱拦截全部通过；
   - **场景 3 (证据账本与压缩模块级复核 - 5/5 PASS)**：真实 Token 算子、阈值压缩保留成对锚点与 6 条活跃尾部、不可变账本生命周期固化、待确认任务防伪核查全部通过；
   - **场景 4 (UI 保存门禁 6 分支模块级复核 - 6/6 PASS)**：抽样匹配放行、待确认放行、报错拦截、预检失败拦截、越权拦截、目标错乱拦截全量与生产服务精准映射；
   - **场景 5 (部署兼容与缓存规范模块级复核 - 4 PASS / 1 NOT TESTED)**：便携启动支持中文与空格路径、静态 System Prompt 字节恒定从结构上保障前缀缓存（真实网关命中率待生产日志确证）、脱机宏库离线运行、升级无损；32 位 Office 宿主客观声明为未在 32 位机器运行回归（`NOT TESTED`）。

---

## 二、 当前 main 与发布提交的关系

- **发布基线提交**：`b4402a9ead90d8e74c451ccb6683a22e20da3ae4`（包含功能收敛与 `RELEASE_NOTES.md`，构建产物即为 `v1.3.0-rc1`）。
- **文档整理与测试资产手册提交**：`eac0f0b`（`docs(chore): reorganize project documentation, test inventory and release verification [skip ci]`，包含全仓文档梳理、历史归档与测试索引）。
- **CI 发布工作流加固提交**：`4eee7b3`（`ci: harden release workflow publication guards and prerelease detection`，彻底修复 main push fallback v1.2.0、修复 prerelease 检测并增加防止覆盖已有 Release 的前置防护）。
- **收尾文档提交**：更新当前状态与公开包冒烟证据记录。
- **发布与整理关系铁律**：已发布的 `v1.3.0-rc1` 发行包由发布提交 `b4402a9` 构建；后续 main 上的文档与 CI 工作流加固提交**不属于该包的构建源码**，绝不宣称旧发行包由最新整理后的 main 构建。

---

## 三、 审计整改专项落地事实 (feat/audit-rectification)

本轮针对 2026-10-19 审查报告（第一至第五轮）完成全面闭环整改，并通过全量离线核心门禁（**233/233 PASS**，100% 保持全绿）：
1. **降级分支接线完整链路断言 (解决第五轮审查 P2/接线盲区)**：
   - 增强 `loadRealTsService` 支持 `forceZeroDepFallback` 显式参数，增加 `Test 20.0b` 门禁断言；
   - 强行触发纯 Node 降级 else 分支加载并执行真实的 `macroSaveEligibility.ts` 纯函数，彻底覆盖从源码读取、语法剥离到沙箱注入执行的完整接线链路。
2. **纯文本兜底正则技术债显式标注 (解决第五轮审查 P2/架构清晰度)**：
   - 在 `compaction.ts` 中明确增加架构演进注释，将 `verifiedRegex` 散文兜底定性为“服务无元数据旧会话的过渡兼容技术债”，明确后续废弃收敛路径。
3. **彻底清除 `verified` 死状态（第四轮审查 P2 彻底闭环）**：
   - 真实底层 `verifyExecutionResult` 仅产出 `matched`、`unconfirmed`、`failed` 三态；
   - 彻底将 `bridge.ts`、`llm.ts`、`conversationManager.ts`、`App.svelte` 与 `ExecutionCard.svelte` 中的联合类型及判断收敛，移除虚假的 `verified` 成员，使类型契约 100% 反映生产真实。
4. **证据载荷真实性门控 (第三轮审查 P1-新 彻底闭环)**：
   - 彻底解决“管道已通但载荷失真”漏洞：在 `conversationManager.ts` 组装结构化元数据时，增加严格状态门控 `verificationStatus === 'matched' && !error`；
   - 仅当抽样核验真正一致且无运行异常时，才向证据账本写入 `区域 ${usedRangeAddress} 抽样核验一致`；对未确认的排版任务或失败报错任务，绝不虚构“验证通过”，确保不可变证据账本内容 100% 真实。
5. **抽离门禁为独立纯 TS 模块并测真身 (第三轮审查 P1-新 彻底闭环)**：
   - 新增独立纯函数服务模块 `web/src/services/macroSaveEligibility.ts`，导出 `checkMacroSaveEligibility`；
   - `ExecutionCard.svelte` 组件接入该纯函数计算保存资格；
   - `test_suite_unit.cjs` 通过动态转译直接加载该生产服务模块进行断言，生产与测试共享同一真实代码，彻底根治门禁逻辑的“测试抄写副本”弊病。
6. **架构描述诚实严谨化与工具校验门禁**：
   - 清理代码与文档中关于“In-Process Tool Loop”等夸大措辞，实事求是更正为“**Native COM Fast-Path Dispatch（原生轻量快速分派通道）**”；
   - `web/src/services/toolRegistry.ts` 作为前端工具唯一事实来源；C# `src/NativeBridge.cs` 增加坐标格式严格正则拦截；C# 加载项重新编译生成合规 `bin/LeeExcel.dll`。
7. **仓库根目录卫生与历史报告归档**：
   - 将根目录的五份审查报告（v1、v2、v3、v4、v5）全部归档至 `docs/history/reviews/`，严格遵循 `AGENTS.md` 规约保持根目录干净；
   - 前端工作台重新构建生成最终单文件静态包 `bin/dist/`。

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
