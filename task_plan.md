# 项目规划与任务计划书 (Task Plan) : ExcelMind AI

> **主路线图引用**：本文件执行阶段计划严格遵循唯一主规划 [docs/product-roadmap.md](docs/product-roadmap.md)。  
> **当前状态入口**：参见 [docs/CURRENT_STATE.md](docs/CURRENT_STATE.md)  
> **基准发布版本**：v1.3.0-rc1 (Git Commit: `b4402a9ead90d8e74c451ccb6683a22e20da3ae4`)  
> **当前任务**：发布后资产核对、项目文档与测试资产整理  
> **当前状态**：**已全部完成并就绪 (Completed & Ready to Push)**  
> **当前停点**：工作流加固、Release 元数据更正、官方包解压冒烟与文档收尾均已完成。  
> **下一步计划**：无未授权代码开发任务；R7 为远期评估项，当前不启动。  
> **授权边界**：不新增功能、不重新设计 UI、不修改已发布 tag、不覆盖 Release 资产、不重新发布已有版本、不启动真实 API。

---

## 1. 当前活动任务与执行状态

| 任务编号 | 任务内容 | 状态 | 验证与产出物 | 停点判定与授权边界 |
| :--- | :--- | :--- | :--- | :--- |
| `TASK-CHORE-01` | **已发布真实基线核对**<br>只读核查 GitHub 远端 main、tag、Release 状态及 Pre-release 标记 | `done` | `docs/CURRENT_STATE.md`<br>确认 tag 指向 `b4402a9`，记录初始状态 | 只读核对，不移动 tag |
| `TASK-CHORE-02` | **GitHub 资产与冒烟证据核对**<br>下载官方 Release ZIP，对比本地构建包，逐字节比对解压载荷哈希 | `done` | 两份包文件清单完全一致；确认换行符与编译器元数据差异 | 不覆盖已发布资产 |
| `TASK-CHORE-03` | **项目文档职责梳理与入口建立**<br>新增 `docs/CURRENT_STATE.md`，更新 `README.md`、`docs/product-roadmap.md` 等 | `done` | `docs/CURRENT_STATE.md`<br>`README.md`<br>`INSTALL.md`<br>`USER_GUIDE.md`<br>`RELEASE_NOTES.md` | 沿用现有体系，不建第二套路线图 |
| `TASK-CHORE-04` | **历史旧文档归档与纠偏**<br>建立保留/更新/归档/清理清单，归档旧规范至 `docs/history/` 并标明免责声明 | `done` | `docs/history/plans/`<br>`docs/history/specs/`<br>`findings.md` | 标注历史状态，不掩盖真实历史 |
| `TASK-CHORE-05` | **测试资产手册与隔离规范更新**<br>梳理测试命令、范围、环境要求、产物位置与安全隔离红线 | `done` | `tests/README.md`<br>`AGENTS.md` | 测试默认隔离，严禁污染正式数据 |
| `TASK-CI-01` | **发布工作流加固**<br>消除 main push fallback v1.2.0，修复 prerelease 判定，增加已有 Release 覆盖保护 | `done` | `.github/workflows/release.yml`<br>离线 9 项分支/PR 场景测试 100% PASS | 独立提交 `4eee7b3` |
| `TASK-REL-01` | **纠正现有 Release 元数据**<br>PATCH API 更新 Release `405746132`，prerelease 设为 true，补充限制说明 | `done` | GitHub API 响应验证 `prerelease: true`<br>不删除 Release、不移动 tag、不覆盖 ZIP | 零工作流副作用 |
| `TASK-SMOKE-01` | **GitHub 官方下载包最小冒烟**<br>针对 `445ef...` 真实包运行清单校验、64 位 Excel 独立挂载、Ribbon 截图与无害宏读回 | `done` | `.artifacts/tests/smoke_github_release_20261007_214500/`<br>6/6 PASS，实例安全退出 | 真实下载包真机存证 |
| `TASK-CHORE-06` | **按影响范围测试策略与安全提交**<br>更新 CURRENT_STATE，文档提交附带 [skip ci]，推送 main | `done` | 准备就绪，工作流具备多道防误发布保护 | 遵循既定规则推送 |

---

## 2. 已交付核心能力 (R0 ~ R6) 任务摘要与证据索引表

| 任务 ID | 所属阶段 | 核心交付内容与用户价值 | 验收状态 | 关键代码位置 | 历史验收证据与报告索引 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `TASK-R0-01` | R0 | **基线治理与代码保真**<br>统一主规划路线图，100% 源码保真，建立 8 项不可退化门禁。 | `accepted` | `src/VbaRunner.cs`<br>`src/ScriptManager.cs` | `test_suite_unit.cjs`<br>`test_regex_counter_example.cjs` |
| `TASK-R1a-01` | R1a | **选区只读感知**<br>COM 批量读取选区结构，前端展示只读卡片，严格防止提示词注入。 | `accepted` | `src/SelectionContextService.cs`<br>`web/src/components/ChatInput.svelte` | 用户人工实测通过；单元测试全绿 |
| `TASK-R1b-01` | R1b | **选区时效动态提示与前序宏显式引用**<br>选区采集时间展示、原区域定向刷新、前序宏卡片引用与字符量客观统计。 | `accepted` | `src/SelectionContextService.cs`<br>`web/src/services/conversationManager.ts` | `.artifacts/tests/desktop_acceptance_20261002_144323/`<br>真实桌面自动化验收通过 |
| `TASK-R2a-01` | R2a | **宏库检索、标签与运行历史溯源**<br>本地 `.bas`+`.meta.json` 存储，多维检索过滤，最近 10 次执行记录追溯。 | `accepted` | `src/ScriptManager.cs`<br>`web/src/components/ScriptDrawer.svelte` | `.artifacts/tests/desktop_acceptance_20261002_150757/`<br>真实桌面自动化验收通过 |
| `TASK-R2b-01` | R2b | **Ribbon 动态菜单展示收藏宏**<br>Excel 功能区动态渲染“常用宏”菜单，safeId 防串选，执行前二次确认。 | `accepted` | `src/LeeExcelRibbon.cs`<br>`src/LeeExcelAddIn.cs` | `.artifacts/tests/desktop_acceptance_20261002_155217/`<br>真实桌面自动化验收通过 |
| `TASK-R2c-01` | R2c | **显式参数化宏契约**<br>7 种基础参数类型强校验，自动渲染交互式输入表单，独立包装器隔离。 | `accepted` | `src/VbaSignatureParser.cs`<br>`src/VbaRunner.cs` | `.artifacts/tests/desktop_acceptance_20261002_161002/`<br>真实桌面自动化验收通过 |
| `TASK-R3a-01` | R3a | **确定性快捷去重工具**<br>本地哈希算法，多列复合键，前导零与长文本保真，强制前置快照。 | `accepted` | `src/DataToolsService.cs`<br>`web/src/components/DataToolModal.svelte` | `.artifacts/tests/desktop_acceptance_20261002_173553/`<br>真实桌面自动化验收通过 |
| `TASK-R3b-01` | R3b | **两表主键差异对账工具**<br>长文本/前导零保真，双端统计恒等式，独立新表输出，源表零篡改。 | `accepted` | `src/DataToolsService.cs`<br>`web/src/components/DataToolModal.svelte` | `.artifacts/tests/desktop_acceptance_20261002_184553/`<br>真实桌面自动化验收通过 |
| `TASK-R4a-01` | R4a | **批量宏任务队列与隔离执行**<br>串行 COM 调度，隔离工作副本试运行，防漂移预检，遇错即停与安全取消。 | `accepted` | `src/BatchRunnerService.cs`<br>`src/NativeBridge.cs` | `.artifacts/tests/desktop_acceptance_20261002_190824/`<br>真实桌面自动化验收通过 |
| `TASK-R4b-01` | R4b | **批量宏任务可视化面板**<br>STA 受控线程调度，纯快照单发轮询，6 态统计卡片，一键定位输出目录。 | `accepted` | `web/src/components/BatchModal.svelte`<br>`src/BatchRunnerService.cs` | `.artifacts/tests/desktop_acceptance_20261002_194637/`<br>真实桌面自动化验收通过 |
| `TASK-R4c-01` | R4c | **多文件列名对齐汇总工具**<br>多工作簿同名列智能对齐，加权消除同名异径歧义，行数守恒恒等式。 | `accepted` | `src/DataToolsService.cs`<br>`web/src/components/DataToolModal.svelte` | `.artifacts/tests/desktop_acceptance_20261002_203118/`<br>真实桌面自动化验收通过 |
| `TASK-R5a-01` | R5a | **双步骤流水线轻量串联**<br>清洗→对账等两步线性流水线，本地 JSON 版本化定义，全流程任务级快照。 | `accepted` | `src/WorkflowManager.cs`<br>`web/src/components/WorkflowModal.svelte` | `.artifacts/tests/desktop_acceptance_20261002_213703/`<br>真实桌面自动化验收通过 |
| `TASK-R5b-01` | R5b | **自选区域汇总与图表生成**<br>模型生成主链 100% 保真，快捷图表工具走独立显式路径，防旧结果叠加。 | `accepted` | `src/ChartService.cs`<br>`web/src/components/DataToolModal.svelte` | `.artifacts/tests/desktop_acceptance_20261003_065534/`<br>真实桌面自动化验收通过 |
| `TASK-R6a-01` | R6a | **只读外部数据接入**<br>CSV/JSON/HTTP GET 只读映射，URL 白名单与 SSRF 防护，DPAPI 凭据加密。 | `accepted` | `src/ExternalDataService.cs`<br>`web/src/components/DataToolModal.svelte` | `.artifacts/tests/desktop_acceptance_20261003_080107/`<br>真实桌面自动化验收通过 |
| `TASK-R6b-01` | R6b | **无凭据宏包导入/导出**<br>标准 `.exmpack` (Zip) 格式，白名单脱敏 manifest，源码逐字节保真，路径穿越与 Zip 炸弹防护。 | `accepted` | `src/MacroPackageManager.cs`<br>`web/src/components/ScriptDrawer.svelte` | `.artifacts/tests/desktop_acceptance_20261003_083458/`<br>真实桌面自动化验收通过 |
| `TASK-R6c-01` | R6c | **无损安装升级与脱敏诊断**<br>三态物理隔离保护用户数据，文件占用安全退出，两道脱敏防线本地导出诊断包。 | `accepted` | `scripts/core/install_addin.ps1`<br>`src/DiagnosticsService.cs` | `.artifacts/tests/r6c_targeted_20261003_085240/`<br>定向验收 9 项全部通过 |
| `TASK-UI-R6` | UI收敛 | **Ribbon 与任务窗格交互收敛**<br>Ribbon 品牌大图标（无按钮文本，Tooltip 说明），单模式入口与单发送按钮，320px/400px 响应式自适应。 | `accepted` | `src/LeeExcelRibbon.cs`<br>`web/src/components/ChatInput.svelte` | 用户人工视觉确认通过；独立解压冒烟验证通过 (`.artifacts/tests/smoke_isolated_20261007_203405/`) |

---

## 3. 核心离线回归验证清单 (必须 100% 保持)

每次变更必须执行并确保通过的 8 项核心防倒退原则：
1. **对话不执行**：CHAT 模式下用户任何指令均不调用 COM 宏执行器，不触发快照创建。
2. **操作写值**：AUTOMATION 模式下有效 VBA 成功在目标工作簿写值，写后读回 `WorkbookReadback` 准确捕获区域。
3. **无代码回复不伪报宏失败**：AUTOMATION 模式下纯文本说明回复不触发执行，不伪报宏执行错误。
4. **导入保存重开**：导入宏持久化到 `%APPDATA%\ExcelMindAI\Scripts\`，重开 Excel 宏库列表完整可见。
5. **源码保真**：VBA 宏引号、中文、`Option Explicit`、模块属性绝不被静默正则改写。
6. **运行错误展示**：宏执行失败如实汇报阶段（编译期/注入期/运行期）与错误描述，不报虚假成功。
7. **快照恢复**：执行前在 `%APPDATA%\ExcelMindAI\Backups\` 物理备份；恢复前保存救援副本。
8. **冷启动无重复注册**：内部 C# 函数不滥用 public static，Excel-DNA 冷启动无重复函数注册异常。
