# ExcelMind AI 测试与诊断资产手册 (tests/README.md)

> **定位**：项目长期测试资产目录索引、有效命令、环境要求与隔离规范  
> **更新时间**：2026-10-07  
> **当前状态入口**：参见 [docs/CURRENT_STATE.md](../docs/CURRENT_STATE.md)  
> **质量准则**：证据驱动、测试隔离、安全无损；严禁以内存模拟冒充真实宿主验证，未验证范围明确如实声明。

---

## 一、 快速开始：有效测试命令与运行方式

### 1. 核心离线质量门禁 (必须 100% PASS)
> **环境要求**：Node.js 18+，**无需启动 Excel 进程**，零网络请求，零费用消耗。  
> **适用场景**：日常提交、CI/CD 构建及任何代码修改后的必跑门禁。

```bash
# 1. 核心单元与契约测试 (204+ 项用例)
node test_suite_unit.cjs

# 2. VBA 源码保真正则反例守卫测试 (3 项反例)
node test_regex_counter_example.cjs
```

- **测试范围**：
  - 意图识别与通道路由 (CHAT 问答 vs AUTOMATION 自动化)；
  - 结构化 VBA 代码提取与未闭合围栏拦截；
  - 源码 100% 保真与智能引号保护（绝不静默改写或删除 `Option Explicit`）；
  - 选区感知数据契约与低信任 Prompt 围栏防注入；
  - 宏库元数据读写、标签检索、安全 ID (safeId) 防串选；
  - 显式参数化宏契约校验与防代码逃逸转义；
  - 批量宏任务调度状态机与防漂移预检；
  - 双步骤流水线定义与单步失败短路保护；
  - 外部数据接入 URL 白名单、SSRF 与敏感凭据脱敏；
  - 无凭据宏包 Zip 结构、白名单 manifest、路径穿越与 Zip 炸弹防御；
  - 安装升级三态隔离与 8 类脱敏诊断白名单。

---

### 2. 独立发行包解压隔离冒烟验证 (Release Smoke)
> **环境要求**：Windows 10/11，64 位 Microsoft Excel 桌面版，.NET 4.0/4.6+。  
> **隔离机制**：解压至 `.artifacts/tests/<run-id>/extracted_pkg/`，临时命令行挂载加载项，**零修改正式注册表自启动项，零触碰真实用户数据目录**，测试退出即安全释放 COM。

```powershell
# 针对官方发布包或本地构建包执行解压隔离冒烟
pwsh -File tests/diagnostics/run_isolated_release_smoke.ps1 -ZipPath ".artifacts/release/ExcelMindAI-v1.3.0-rc1.zip"
```

- **检查项 (6 项 100% PASS)**：
  1. `TC-SMOKE-01`：ZIP 解压完整性与 26 项载荷文件存在性核对；
  2. `TC-SMOKE-02`：C# 核心 DLL 依赖图与反射解析可用性；
  3. `TC-SMOKE-03`：前端静态资源索引引用与 HTML/JS/CSS 完整性；
  4. `TC-SMOKE-04`：核心服务（设置、宏存储、快照管理器）无工作簿环境单例初始化；
  5. `TC-SMOKE-05`：真实 Excel 临时挂载加载 XLL，验证功能区 (Ribbon) 品牌大图标与分组正常展现并截图；
  6. `TC-SMOKE-06`：任务窗格加载与固定无害宏注入执行，COM 真实读回单元格数据校验。
- **输出位置**：`.artifacts/tests/smoke_isolated_<timestamp>/`（包含 `smoke_report.md` 及截图）。

---

### 3. 真实 Excel 桌面端全量自动化验收套件 (Desktop Acceptance Runner)
> **环境要求**：Windows 11 64 位，64 位 Microsoft Excel 桌面版。  
> **设计原理**：C# STA 独立执行器驱动真实 Excel 进程，通过 UI Automation 操作 Ribbon 与任务窗格，通过 COM 检查单元格与快照。

```powershell
# 运行桌面自动化验收套件
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run_desktop_acceptance.ps1
```

- **源码位置**：`tests/tools/DesktopAcceptanceRunner.cs`
- **输出位置**：`.artifacts/tests/desktop_acceptance_<timestamp>/`（包含详细 JSON 读回、断言报告与真机桌面截图）。

---

## 二、 测试工具源码清单 (`tests/tools/`)

| 工具源码文件 | 职能分类 | 具体用途 | 编译输出建议 | 副作用与环境要求 |
| :--- | :--- | :--- | :--- | :--- |
| **`VerifyIsolatedReleaseSmoke.cs`** | 发行包冒烟 | 解压发布 ZIP，执行 6 项独立隔离冒烟（反射/前端/无工作簿服务/真机加载/宏读回） | `.artifacts/tests/<id>/VerifyIsolatedReleaseSmoke.exe` | 启动独立 Excel 实例；退出时安全释放 |
| **`DesktopAcceptanceRunner.cs`** | 桌面端自动化验收 | R1~R6 核心功能全量真机桌面端自动化验收套件（STA 驱动，UIA+COM） | `.artifacts/tests/<id>/DesktopAcceptanceRunner.exe` | 启动独立 Excel 实例；生成临时工作簿与截图 |
| **`CaptureRibbonEvidence.cs`** | 界面取证 | 自动化挂载插件并捕获 Excel 顶部功能区 (Ribbon) 高清排版截图 | `.artifacts/tests/<id>/CaptureRibbonEvidence.exe` | 启动独立 Excel 实例；生成 PNG 截图 |
| **`VerifyRibbonLayout.cs`** | 界面取证 | 遍历功能区 UI Automation 元素树，核查按钮标签、大小与折行状态 | `.artifacts/tests/<id>/VerifyRibbonLayout.exe` | 检查前台 Excel 窗口；只读枚举 UIA 控件 |
| **`VerifySelectionContext.cs`** | 专项集成 | 专项验证选区只读抽样服务、大选区截断与 Value2 数组批量读取 | `.artifacts/tests/<id>/VerifySelectionContext.exe` | 启动独立 Excel 实例；只读读取选区 |
| **`VerifyScriptManagerR2a.cs`** | 专项集成 | 验证宏库本地持久化、多维标签检索与最近 10 次执行记录更新 | `.artifacts/tests/<id>/VerifyScriptManagerR2a.exe` | 纯文件系统操作；使用测试隔离目录 |
| **`VerifyMacroFeature.cs`** | 专项集成 | 验证 `.bas` 源码与 `.meta.json` 元数据持久化、导出与旧版本迁移 | `.artifacts/tests/<id>/VerifyMacroFeature.exe` | 纯文件系统操作；使用测试隔离目录 |
| **`VerifySaveCopyAsDirtyState.cs`** | 专项集成 | 验证 Excel `SaveCopyAs` 对未保存脏数据（如 A1 单元格修改）的备份行为 | `.artifacts/tests/<id>/VerifySaveCopyAsDirtyState.exe` | 启动独立 Excel 实例；生成测试工作簿 |
| **`VerifyScriptHistoryResolution.cs`** | 专项集成 | 验证宏历史记录 ID 关联、哈希校验与歧义排除机制 | `.artifacts/tests/<id>/VerifyScriptHistoryResolution.exe` | 纯文件系统操作；无 COM 副作用 |
| **`VerifyR6cInstallAndDiagnostics.cs`** | 专项集成 | 验证无损安装部署、三态物理隔离与 8 类脱敏诊断包本地导出 | `.artifacts/tests/<id>/VerifyR6cInstallAndDiagnostics.exe` | 纯文件系统操作；使用隔离临时目录 |
| **`VerifyFocusHandover.cs`** | 缺陷排查 | 验证任务窗格与 Excel 工作表之间的 Win32 焦点平滑交接 | `.artifacts/tests/<id>/VerifyFocusHandover.exe` | 需要运行中 Excel 实例；Win32 焦点测试 |
| **`VerifyPayload.cs`** | 离线解析 | 离线解析检验模型下发的 Payload JSON，校验代码拆解与哈希匹配 | `.artifacts/tests/<id>/VerifyPayload.exe` | 纯离线 JSON 解析，零网络或 COM 副作用 |
| **`FindDialogButtons.cs`** | 历史排查 | 遍历查找 Excel 弹窗中的所有控件按钮句柄与文本 | `.artifacts/tests/<id>/FindDialogButtons.exe` | 枚举系统顶级窗口与子控件 |
| **`ReadDialogText.cs`** | 历史排查 | 读取特定 Excel 对话框的详细提示文字与控件标题 | `.artifacts/tests/<id>/ReadDialogText.exe` | 枚举系统顶级窗口与子控件 |
| **`MemoryDump.cs`** | 历史排查 | 扫描特定进程内存特征码（如 VBA 宏切片），提取调试证据片段 | `.artifacts/tests/<id>/MemoryDump.exe` | 读取目标进程内存并写文件 |

---

## 三、 诊断与排查脚本清单 (`tests/diagnostics/`)

| 脚本文件名 | 职能分类 | 运行命令 | 副作用与环境限制 |
| :--- | :--- | :--- | :--- |
| **`run_isolated_release_smoke.ps1`** | 发行包冒烟入口 | `pwsh -File tests/diagnostics/run_isolated_release_smoke.ps1 -ZipPath <path>` | 启动独立 Excel 实例，完全隔离，退出释放 |
| **`capture_ribbon_evidence.ps1`** | Ribbon 自动化截图 | `pwsh -File tests/diagnostics/capture_ribbon_evidence.ps1` | 启动测试 Excel 并截取 Ribbon 原生高清图 |
| **`run_r6c_targeted_verification.ps1`** | R6c 专项验收入口 | `pwsh -File tests/diagnostics/run_r6c_targeted_verification.ps1` | 运行 R6c 安装升级与脱敏诊断定向验收 |
| **`enum_excel_wins.ps1`** | 窗口诊断 | `powershell -NoProfile -File tests/diagnostics/enum_excel_wins.ps1` | 只读枚举 Excel 窗口句柄，无副作用 |
| **`get_windows.ps1`** | 窗口诊断 | `powershell -NoProfile -File tests/diagnostics/get_windows.ps1` | 只读枚举系统前台窗口标题与 PID，无副作用 |
| **`read_config.ps1`** | 配置诊断 | `powershell -NoProfile -File tests/diagnostics/read_config.ps1` | 读取本机 WebView2 本地配置，自动脱敏 Key |
| **`read_vbe_dialog.ps1`** | 弹窗排查 | `powershell -NoProfile -File tests/diagnostics/read_vbe_dialog.ps1 -WorkbookPath <xlsx> -VbaSnippetPath <vba>` | 需显式参数；启动独立 Excel 读取 VBE 对话框 |
| **`spy_window.ps1`** | 弹窗排查 | `powershell -NoProfile -File tests/diagnostics/spy_window.ps1` | 启动后台 Excel 监听 `#32770` 弹窗 |
| **`probe_dialog.ps1`** | 弹窗排查 | `powershell -NoProfile -File tests/diagnostics/probe_dialog.ps1` | 异步探测语法错误弹窗并关闭 |
| **`probe_model.cjs`** | **联网模型探针** | `node tests/diagnostics/probe_model.cjs` | ⚠️ **严禁在自动化测试中自动运行！直接消耗真实 API 账户 Token 与网络流量！** |

---

## 四、 活跃与历史资产分类边界

1. **受保护活跃驱动脚本 (`scratch/`)**：
   - `scratch/call_llm.cjs`、`scratch/test_compile_behavior.ps1` 及其 9 个强依赖驱动脚本处于活跃开发保护中，由 [AGENTS.md](../AGENTS.md) 规则 1 严格保护，**不得擅自改动或强行迁移**。
2. **历史一次性排查草稿**：
   - 历史形成的 174 个单点排查草稿已建立详尽清单归档于 [docs/history/pending_174_scripts_manifest.md](../docs/history/pending_174_scripts_manifest.md)。
   - 此类脚本已退役，**严禁混入日常默认测试命令中运行**。

---

## 五、 测试隔离安全铁律 (AGENTS.md 红线)

所有执行测试的工具和脚本必须无条件遵守以下安全隔离铁律：

1. **数据目录绝对隔离**：
   - 测试必须使用独立隔离目录（如 `.artifacts/tests/<run-id>/`）；
   - **严禁将测试输出路径指向用户的真实数据目录**（`%APPDATA%\ExcelMindAI\` 或旧目录 `%APPDATA%\LeeExcel\`）；
   - 严禁测试向正式注册表注入持久化自启动项。
2. **外部网络与计费限制**：
   - 严禁在自动化回归或无用户明确授权的前提下自动运行真实商业大模型 API（如 `probe_model.cjs`、`call_llm.cjs`）；
   - 单元测试与端到端测试默认采用 Mock 或离线固定样本。
3. **Office 宿主与进程安全**：
   - 测试启动的 Excel 进程必须在 `try ... finally` 结构中确保 `Close(false)` 与 `Quit()` 释放 COM 对象；
   - **严禁使用 `taskkill /f /im excel.exe` 强杀用户的 Excel 进程**，严禁触碰或关闭用户本地已打开的工作簿。
4. **编译与生成物归档规范**：
   - 编译测试工具必须使用系统自带 64 位 `v4.0.30319/csc.exe` 并显式指定 `/out:".artifacts\tests\<run-id>\<name>.exe"`；
   - 严禁直接在 `tests/`、`bin/` 或项目根目录下生成测试可执行程序、临时工作簿或运行日志；
   - 严禁将临时运行日志或编译产物提交到 Git。
