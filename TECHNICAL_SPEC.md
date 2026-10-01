# Lee-Excel 技术规范与说明书 (Technical Specification)

> **版本**：v1.0 (基线治理后标准版本)  
> **文档性质**：系统技术规范、架构规约与开发者操作手册  
> **编写日期**：2026-10-01  
> **面向对象**：系统维护工程师、发布工程师与插件二次开发者  

---

## 1. 项目用途与整体架构

### 1.1 系统定位
Lee-Excel 是专为 Windows 平台 Microsoft Excel 桌面端开发的原生 AI 加载项。通过整合大语言模型强大的理解推理能力与本地 Office COM 自动化接口，提供侧边栏沉浸式交互，帮助用户高效完成数据加工、公式填充、格式美化与数据分析。

### 1.2 整体架构设计
系统采用分层松耦合架构：
1. **宿主层 (Excel Host)**：原生 Excel 进程，通过 COM 接口加载标准 XLL 二进制加载项并渲染 Ribbon 扩展菜单。
2. **桥接与容器层 (.NET Add-In Core)**：基于 Excel-DNA 框架构建 C# 类库，向宿主注入右侧原生任务窗格（CustomTaskPane），内部宿主 Microsoft Edge WebView2 控件。
3. **交互表示层 (Web TaskPane UI)**：Svelte 5 构建的 Office Fluent 风格单页应用（SPA），运行在 WebView2 的独立渲染进程中。
4. **调度网关 (Native Bridge)**：利用 WebView2 原生消息通道，通过 JSON 协议在 Web 前端与 C# 宿主之间进行异步双向调度。
5. **执行保护引擎 (VBA Runner & Snapshot Engine)**：负责宏代码提取标准化、多工作簿目标锁定、写后状态读回核验与数据快照防损撤回。

---

## 2. 实际技术栈与版本来源依据

| 技术组件 | 实际版本 / 规范 | 证据与来源依据 | 职责说明 |
| :--- | :--- | :--- | :--- |
| **开发语言 (后端)** | C# (C# 7.3 / .NET CLR 4.0) | `build_addin.ps1` 调用系统 .NET 4.0 csc 编译器 | 插件业务逻辑、COM 调度与快照引擎 |
| **运行时框架** | .NET Framework 4.6.2+ | `packages/` 中引用的 DLL 均位于 `lib/net462/` 目录 | 宿主程序集底层运行时 |
| **Excel-DNA** | 1.9.0 | `packages/ExcelDna.AddIn.1.9.0/` | XLL 加载项打包、COM Ribbon 注册 |
| **WebView2** | 1.0.4191.47 | `packages/Microsoft.Web.WebView2.1.0.4191.47/` | Office 任务窗格内嵌现代 Chromium 浏览器 |
| **前端框架** | Svelte 5 (`^5.20.0`) | `web/package.json` | 现代轻量化响应式 UI 渲染 |
| **构建与前端工具** | Vite 6 (`^6.1.0`) + TypeScript 5 (`^5.7.3`) | `web/package.json` | 前端模块化编译与打包 |
| **样式体系** | Vanilla CSS (Office Fluent 主题) | `web/src/styles/office-fluent.css` | 严格贴合 Office 经典商务体验 |

---

## 3. 清理后的目录结构与职责规约 (第二批长期治理架构)

```text
lee-excle/
├── .artifacts/                       # 【长期生成物】运行输出、构建物与临时调试隔离目录（.gitignore 忽略）
│   ├── tests/<run-id>/               # 每次测试执行独立输出目录（EXE、DLL、日志、临时工作簿、抓包）
│   └── tmp/<session-id>/             # 临时会话调试代码切片与中间产物
├── .backup_cleanup_batch1_20261001/  # 第一批清理文件物理备份与回滚目录（123个文件）
├── .backup_cleanup_batch2_20261001/  # 第二批清理文件物理备份、manifest与回滚脚本（66个文件）
├── bin/                              # 插件运行与正式交付目录（严禁存放测试EXE、临时工作簿或运行日志）
│   ├── LeeExcel64.xll                # 64 位 Excel 加载项原生入口（不可删除）
│   ├── LeeExcel.dll                  # C# 插件主程序集（不可删除）
│   ├── LeeExcel.dna / LeeExcel64.dna # Excel-DNA 配置文件（不可删除）
│   ├── ExcelDna.Integration.dll      # Excel-DNA 核心运行时
│   ├── Microsoft.Web.WebView2.*.dll  # WebView2 运行时程序集
│   ├── WebView2Loader.dll            # WebView2 原生加载器
│   └── dist/                         # 前端 Vite 生产构建产物（不可删除）
│       ├── index.html                # 前端主入口 HTML
│       └── assets/                   # 打包生成的 JS 与 CSS 资源
├── docs/                             # 项目规范与历史档案
│   ├── FILE_LIFECYCLE.md             # 文件生命周期、目录职责与清理规范
│   └── history/evidence_202609/      # 核心历史验收证据（9个关键日志与验证基线文件）
├── packages/                         # NuGet 离线依赖包目录（已跟踪，断网编译保障）
├── scratch/                          # 【当前临时实验区】仅保留当前活跃的 11 个高阶 E2E 驱动与被修改脚本
│   ├── call_llm.cjs                  # [受保护] 本地模型调用驱动（含未提交修改）
│   ├── test_compile_behavior.ps1     # [受保护] 编译行为探针（含未提交修改）
│   └── *.ps1 (9个强依赖驱动)         # [受保护] 强依赖 call_llm.cjs 的高阶验收套件
├── scripts/                          # 项目工程维护与安全清理脚本
│   └── clean_artifacts.ps1           # 安全生成物清理入口（默认预览，-Apply 执行，支持14天生命周期管理）
├── src/                              # C# 原生插件业务源码
│   ├── LeeExcelAddIn.cs              # 插件入口点，实现 IExcelAddIn 与 ExcelRibbon
│   ├── NativeBridge.cs               # 前后端通信 JSON 路由与分发网关
│   ├── ScriptManager.cs              # 本地宏脚本持久化管理器
│   ├── SimpleJson.cs                 # 独立零依赖的 JSON 序列化/反序列化器
│   ├── SnapshotManager.cs            # 工作簿快照与崩溃恢复管理器
│   ├── TaskPaneControl.cs            # CustomTaskPane 窗体控件与 WebView2 映射
│   └── VbaRunner.cs                  # VBA 宏提取注入、目标锁定与写后状态读回
├── tests/                            # 【长期测试资产】正式测试源码、诊断脚本与固定测试说明
│   ├── README.md                     # 测试工具编译、运行、依赖与副作用说明书
│   ├── tools/                        # 9 个保留的 C# 测试工具源码（显式编译输出至 .artifacts/）
│   └── diagnostics/                  # 7 个诊断脚本（窗口探针、COM分析、模型配置探测）
├── web/                              # 前端工程目录
│   ├── src/                          # Svelte 5 组件与服务
│   │   ├── components/               # 交互组件 (ChatInput, ExecutionCard 等)
│   │   ├── services/                 # 前端服务 (llm.ts, bridge.ts, config.ts)
│   │   └── styles/                   # 样式定义
│   ├── package.json / vite.config.ts # 前端构建配置
│   └── index.html                    # 前端模板入口
├── AGENTS.md                         # AI 协同开发规约（文件存储、API限制、编译规范）
├── build_addin.ps1                   # C# 插件编译与组织脚本
├── run_excel.ps1                     # Excel 插件挂载启动脚本
├── test_suite_unit.cjs               # 核心单元测试套件（27 项通过）
├── test_regex_counter_example.cjs    # 架构设计反例演示验证脚本
├── task_plan.md                      # 全局项目规划书（含历史质量修复记录）
├── TECHNICAL_SPEC.md                 # 项目技术规范说明书（本文档）
├── progress.md                       # 执行进度与治理日志
└── .gitignore                        # 限制性生成物与备份忽略配置文件
```

---

## 4. 核心模块职责与调用链路

### 4.1 核心 C# 类与职责
1. **`LeeExcelAddIn`** (`src/LeeExcelAddIn.cs`)：
   - 继承 `ExcelRibbon` 并实现 `IExcelAddIn`。
   - `AutoOpen()`：在 Excel 启动时自动触发，创建右侧 430px 的任务窗格并挂载 `WorkbookActivate`、`WorkbookOpen`、`NewWorkbook` 等事件通知。
   - 自定义 Ribbon：注入 `LeeExcelTab`，展示“AI 助手”功能组。
2. **`TaskPaneControl`** (`src/TaskPaneControl.cs`)：
   - Windows Forms 用户控件，承载 `Microsoft.Web.WebView2.WinForms.WebView2`。
   - 使用 `SetVirtualHostNameToFolderMapping` 将虚拟域名 `https://app.lee-excel/` 映射到本地 `bin/dist/` 物理目录，解决本地 `file://` 协议下的 ES Module CORS 限制。
   - 监听 `WebMessageReceived` 事件，将前端 JSON 消息转交给 `NativeBridge`。
3. **`NativeBridge`** (`src/NativeBridge.cs`)：
   - 核心调度网关。接收前端动作指令：
     - `get_workbook_info`：提取当前活跃/目标工作簿名、活动工作表与已用区域；
     - `execute_vba`：调度宏注入执行、目标工作簿锁定与写后状态读回；
     - `restore_snapshot`：触发指定快照恢复；
     - `list_scripts` / `save_script`：脚本库增查。
4. **`VbaRunner`** (`src/VbaRunner.cs`)：
   - 自动化核心。实现目标工作簿显式绑定（锁定指定 Workbook，防多窗口串改）；
   - 执行前创建快照；
   - 将动态代码注入目标工作簿的临时 VBComponent 模块，执行入口过程；
   - 执行后清除临时模块；
   - 采集 `WorkbookReadback` 数据（范围、行列数、抽样单元格文本、公式状态、边框、底色）并返回前端。
5. **`SnapshotManager`** (`src/SnapshotManager.cs`)：
   - 实现工作簿安全快照；对未保存的临时工作簿采用二进制与状态镜像保护；对已保存工作簿建立带时间戳的副本。

---

## 5. 前后端通信协议规范

前端与宿主采用标准 JSON 格式进行消息传递，通信机制基于 WebView2 原生双向通道。

### 5.1 前端发送格式 (Web -> Host)
```json
{
  "action": "execute_vba",
  "code": "Sub RunAutomation()\n    Dim ws As Worksheet\n    Set ws = Application.Workbooks(\"Book1.xlsx\").ActiveSheet\n    ws.Range(\"A1\").Value = \"Hello\"\nEnd Sub",
  "prompt": "在 A1 写入 Hello",
  "targetWorkbookName": "Book1.xlsx",
  "requestId": "req_1727740800_abc"
}
```

### 5.2 宿主响应格式 (Host -> Web)
```json
{
  "action": "execute_vba",
  "status": "success",
  "snapshotId": "snap_20261001_080000",
  "readback": {
    "targetWorkbookName": "Book1.xlsx",
    "targetVerified": true,
    "startCell": "A1",
    "usedRangeAddress": "$A$1",
    "rowCount": 1,
    "colCount": 1,
    "hasBorders": false,
    "hasInteriorColor": false,
    "sampleValues": ["Hello"],
    "hasFormulas": false
  },
  "message": "宏已运行完成"
}
```

---

## 6. 开发环境、依赖与构建命令

### 6.1 开发环境要求
- **操作系统**：Windows 10 / Windows 11 (x64)
- **.NET 框架**：.NET Framework 4.8 运行时及 `csc.exe` 编译器（内置于 `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\`）
- **Node.js**：Node.js v18+（推荐使用 pnpm 管理依赖）
- **办公软件**：Microsoft Excel 桌面版（已安装并在注册表注册）

### 6.2 真实构建与运行命令

#### 1. 编译 C# 插件核心
在项目根目录下执行：
```powershell
powershell -ExecutionPolicy Bypass -File .\build_addin.ps1
```
- **输出**：生成 `bin/LeeExcel.dll`，并将必要运行时从 `packages/` 复制到 `bin/`。

#### 2. 构建前端 Web 单页应用
进入 `web/` 目录执行：
```powershell
cd web
pnpm install
pnpm run build
cd ..
```
- **输出**：清空并重新生成 `bin/dist/`，包含 `index.html` 及静态资产。

#### 3. 启动 Excel 并挂载插件
在根目录下执行：
```powershell
powershell -ExecutionPolicy Bypass -File .\run_excel.ps1
```
- **行为**：从注册表读取 `excel.exe` 路径，传参挂载 `bin/LeeExcel64.xll` 启动 Excel。

---

## 7. 现有测试套件与安全规约

| 测试脚本 | 运行时 | 副作用与安全性 | 执行建议 |
| :--- | :--- | :--- | :--- |
| [`test_suite_unit.cjs`](file:///c:/Users/35651/Desktop/Google/lee-excle/test_suite_unit.cjs) | Node.js | **零副作用**（纯内存 mock，零网络，零修改） | **日常高频门禁** |
| [`test_regex_counter_example.cjs`](file:///c:/Users/35651/Desktop/Google/lee-excle/test_regex_counter_example.cjs) | Node.js | **零副作用**（算法反例演示） | **架构校验** |
| [`test_core.ps1`](file:///c:/Users/35651/Desktop/Google/lee-excle/test_core.ps1) | PowerShell + .NET | 向 `%APPDATA%` 写入测试脚本，不修改工作簿 | 离线核心类库回归 |
| [`test_system_suite.ps1`](file:///c:/Users/35651/Desktop/Google/lee-excle/test_system_suite.ps1) | Excel COM | 启动无头 Excel，在 `%TEMP%` 创建隔离文件 | 深度集成回归（需空闲环境） |
| [`test_real_llm_e2e.ps1`](file:///c:/Users/35651/Desktop/Google/lee-excle/test_real_llm_e2e.ps1) | Excel COM + API | **产生实际 API 费用与网络请求** | **受限测试，严禁自动执行** |
| [`tests/tools/`](file:///c:/Users/35651/Desktop/Google/lee-excle/tests/tools/) | C# (.NET 4.8) | 编译输出至 `.artifacts/tests/<run-id>/`，前置参数缺失直接退出；运行期维护 `.running` 与 `meta.json` | 专项功能与回归验证 |
| [`tests/diagnostics/`](file:///c:/Users/35651/Desktop/Google/lee-excle/tests/diagnostics/) | PowerShell / Node | 诊断脚本；参数缺失前置阻断；产物隔离至 `.artifacts/tests/` | 线上缺陷诊断定位 |

---

## 8. 常见故障与排查指引

1. **Excel 启动提示“无法加载加载项”**：
   - 检查 Excel 位数（32位 vs 64位）。本项目默认以 64 位 `LeeExcel64.xll` 运行；若为 32 位 Excel，需加载 `LeeExcel.dna` 对应编译产物。
   - 检查 `bin/` 目录下 `ExcelDna.Integration.dll` 与 `LeeExcel.dll` 是否齐全。
2. **任务窗格显示白屏或初始化引导页**：
   - 检查系统是否安装 Microsoft Edge WebView2 Evergreen Runtime；
   - 检查 `bin/dist/index.html` 是否存在；
   - 按 `F12` 打开 WebView2 DevTools 检查控制台网络与脚本报错。
3. **宏执行提示“目标工作簿身份核验失败”**：
   - 表明在宏注入执行期间，用户在前台切换了活动窗口，或者指定的目标工作簿已被关闭。系统目标绑定防御机制成功拦截了潜在串改。
4. **清理产物与恢复**：
   - 日常清理：执行 `powershell -File scripts/clean_artifacts.ps1`（仅预览）；执行物理清理需显式带 `-Apply`。
   - 历史回滚：第一批清理通过 `.backup_cleanup_batch1_20261001/rollback.ps1` 恢复；第二批清理通过 `.backup_cleanup_batch2_20261001/rollback.ps1` 恢复；测试工具路径修复前基线与备份保存在 `.backups/pre_path_fix_20261001/`。
