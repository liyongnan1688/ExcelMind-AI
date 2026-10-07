# 执行进度日志 (Progress Log)

> **当前状态入口**：参见 [docs/CURRENT_STATE.md](docs/CURRENT_STATE.md)  
> **唯一主规划**：参见 [docs/product-roadmap.md](docs/product-roadmap.md)  
> **当前维护原则**：记录近期有效进度，较早详细流水按里程碑摘要归档。

---

## 近期有效进度

### Session: 2026-10-07 (v1.3.0-rc1 发布后资产核对与文档测试资产整理)
- [x] **发布基线核对**
  - 核查远端仓库 `liyongnan1688/ExcelMind-AI`，发布提交为 `b4402a9ead90d8e74c451ccb6683a22e20da3ae4`。
  - 核查 Tag `v1.3.0-rc1`（指向 `b4402a9`），Release 存在但未标 prerelease（记录为工作流写死 `$isPre=$false` 所致事实）。
  - 确认不修改 tag，不覆盖资产，不删除 Release，不重新发布。
- [x] **GitHub 资产与冒烟证据核对**
  - 下载官方 Release ZIP 并计算物理 SHA-256：`445ef8b481491bc99b31dcd286b1ac109ae6b519fafc51fc456aaa22687ff25e`，确认为用户下载校验权威基准。
  - 本地构建记录包 SHA-256 为 `f544108d73ad8cfc60723b49b2306a225a93db82e89419065fe11991bce83515`。
  - 解压逐字节比对两包 26 个文件：第三方及运行库二进制 100% 相同，文本文件差异为 CRLF vs LF 换行符，DLL 差异为 PE 时间戳与 MVID GUID（仅 44 字节）。两包运行载荷完全等价，冒烟证据有效复用。
- [x] **文档职责梳理与入口建立**
  - 建立统一入口文档 [docs/CURRENT_STATE.md](docs/CURRENT_STATE.md)。
  - 更新 [README.md](README.md) 链接及 Ribbon 当前布局描述。
  - 更新 [docs/product-roadmap.md](docs/product-roadmap.md)，锁定 R0~R6 已交付，R7 远期评估。
  - 更新 [INSTALL.md](INSTALL.md)、[docs/USER_GUIDE.md](docs/USER_GUIDE.md)、[RELEASE_NOTES.md](RELEASE_NOTES.md)。
  - 纠正模型 Key 存储机制说明（LocalStorage 保存，未采用 DPAPI）。
- [x] **历史记录归档**
  - 将 `docs/superpowers/` 早期打包方案移动归档至 `docs/history/plans/` 与 `docs/history/specs/`，添加历史免责声明。
  - 梳理 [findings.md](findings.md) 并标注历史状态。
- [x] **测试资产手册与隔离规范更新**
  - 更新 [tests/README.md](tests/README.md)，明确离线核心门禁、解压冒烟、桌面验收套件及环境/隔离要求。

### Session: 2026-10-07 (Ribbon 布局与前端交互收敛)
- [x] **Ribbon 终态布局确立**
  - 顶部选项卡：`ExcelMind AI`。
  - 第一分组：`ExcelMind AI`，32×32 品牌大图标，下方无按钮文字，Tooltip 提示“打开或收起 ExcelMind AI 工作台”。
  - 第二分组：【宏与数据】纵向堆叠展示（宏库、导入、常用宏动态菜单、数据工具）。
  - 第三分组：【任务与设置】展示批量处理、工作流与设置。
- [x] **前端任务窗格收敛**
  - 模式选择收敛至单一明确入口（操作 / 对话），单发送主按钮文案动态联动。
  - 320px/400px 响应式断言，窄屏自动展示数据工具选择器，删除宣传横幅与残留调试元素。
- [x] **真实环境验证**
  - 真实 Office 桌面截图存证，确认排版无单字逐行排列。
  - 独立隔离解压冒烟测试 6 项 100% PASS (`.artifacts/tests/smoke_isolated_20261007_203405/`)。

---

## 历史里程碑摘要归档 (2026-09-25 ~ 2026-10-03)

| 里程碑时间 | 阶段与核心成果 | 验收状态与证据 |
| :--- | :--- | :--- |
| **2026-10-03** | **R6c 无损安装升级、脱敏诊断与解压冒烟**<br>三态物理隔离保护用户数据，文件占用安全退出，两道脱敏防线，本地导出 Clean Zip；首次完成解压隔离冒烟。 | `accepted`<br>`.artifacts/tests/r6c_targeted_20261003_085240/`<br>`.artifacts/tests/smoke_isolated_20261003_123829/` |
| **2026-10-03** | **R6b 无凭据宏包导入导出**<br>标准 `.exmpack` (Zip) 打包 manifest 与 `.bas` 源码，白名单脱敏，逐字节保真，路径穿越与 Zip 炸弹防御。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261003_083458/` |
| **2026-10-03** | **R6a 只读外部数据接入**<br>CSV/JSON/HTTP GET 字段映射，URL 白名单与 SSRF 防护，配置与 DPAPI 凭据物理隔离，防公式逃逸。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261003_080107/` |
| **2026-10-03** | **R5b 自选区域汇总与原生图表**<br>模型原生 VBA 主链保真，快捷图表工具走独立路径，防旧结果叠加，整本前置快照。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261003_065534/` |
| **2026-10-02** | **R5a 双步骤任务流水线**<br>两步骤线性串联（去重→对账等），本地 JSON 版本化定义，步骤失败短路，任务级快照恢复。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261002_213703/` |
| **2026-10-02** | **R4 批量宏任务队列与多文件汇总**<br>STA 线程受控调度，隔离工作副本试运行，遇错即停与安全取消；多文件同名列精确对齐。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261002_194637/`<br>`.artifacts/tests/desktop_acceptance_20261002_203118/` |
| **2026-10-02** | **R3 确定性去重与两表对账工具**<br>本地哈希算法，多列复合键；19位长编号/前导零/公式文本全保真，双端统计恒等式守恒。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261002_173553/`<br>`.artifacts/tests/desktop_acceptance_20261002_184553/` |
| **2026-10-02** | **R2 宏库检索、Ribbon 动态菜单与参数化契约**<br>宏库多维检索与运行历史；功能区常用宏动态菜单与二次确认；显式参数化契约。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261002_155217/`<br>`.artifacts/tests/desktop_acceptance_20261002_161002/` |
| **2026-10-02** | **R1 选区只读感知与动态抽样**<br>COM 批量读取选区结构，低信任 Prompt 围栏防注入，前序宏显式引用与客观字符统计。 | `accepted`<br>`.artifacts/tests/desktop_acceptance_20261002_144323/` |
| **2026-10-01** | **工作区保守清理与跨机器打包建设**<br>清理历史垃圾产物，建立安全清理入口 `scripts/clean_artifacts.ps1`，建立跨版本打包流水线。 | `accepted`<br>`.backup_cleanup_batch1_20261001/`<br>`.backup_cleanup_batch2_20261001/` |
| **2026-09-25** | **系统性质量修复专项**<br>拆分 CHAT 与 AUTOMATION 双通道，100% 源码保真，建立写后读回 `WorkbookReadback`，锁定回归基线。 | `accepted`<br>`test_suite_unit.cjs` (PASS) |
