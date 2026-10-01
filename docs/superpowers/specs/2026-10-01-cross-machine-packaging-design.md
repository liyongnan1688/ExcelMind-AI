# Lee-Excel 跨机器分发与全版本 Excel 兼容设计规范

> **文档版本**：v1.0  
> **编写日期**：2026-10-01  
> **状态**：已批准待实施 (Approved)  
> **适用目标**：实现 Lee-Excel 插件在任意 Windows 目标机（新/老版本 Excel、32位/64位）开箱即用加载与全自动打包分发。

---

## 1. 背景与核心目标

### 1.1 背景
当前 Lee-Excel 加载项依赖于开发机环境，仅通过 `run_excel.ps1` 命令行调用挂载 64 位的 `LeeExcel64.xll`，并依赖开发机本地路径结构。目标电脑通常存在以下差异：
1. **Excel 架构不一致**：许多政企客户与老电脑安装的是 32 位 (x86) Excel，直接调用 64 位 XLL 将报加载错误；
2. **Excel 版本跨度大**：从 Office 2010、2013 到 2016、2019、2021 及 Microsoft 365，其注册表路径与默认行为各异；
3. **网络安全拦截 (Mark-of-the-Web)**：解压从网盘/微信/局域网下载的压缩包后，系统会给 DLL/XLL 附加安全锁定属性，导致 Office 拦截；
4. **系统依赖可能缺失**：特别是 Windows 10 早期版本或精简版系统缺少 Microsoft Edge WebView2 运行时。

### 1.2 核心目标
1. **全架构兼容**：一套发布包同时内建 32 位 (`LeeExcel.xll`) 与 64 位 (`LeeExcel64.xll`)，自适应不同架构的 Excel；
2. **多版本自动适配**：兼容 Office 2010 (14.0)、Office 2013 (15.0)、Office 2016/2019/2021/365 (16.0)，实现启动自动常驻加载；
3. **极简安装与便携体验**：提供“一键安装(常驻)”、“一键卸载”以及“免安装便携启动”，小白用户双击即用；
4. **依赖自动探测与修复**：安装前自动解锁文件锁定，并检测 WebView2，提供官方轻量 Bootstrapper 静默安装引导；
5. **一键全自动打包流水线**：开发端提供 `scripts/package_release.ps1`，单命令输出标准化分发压缩包。

---

## 2. 整体架构与发布包组织规范

### 2.1 发布包物理结构 (`LeeExcel_Release/`)

```text
LeeExcel_Release/
├── LeeExcel.xll                     # 32 位 Excel 原生加载项入口 (约 733KB)
├── LeeExcel.dna                     # 32 位 Excel-DNA 配置文件
├── LeeExcel64.xll                   # 64 位 Excel 原生加载项入口 (约 658KB)
├── LeeExcel64.dna                   # 64 位 Excel-DNA 配置文件
├── LeeExcel.dll                     # C# 插件主程序集 (Any CPU)
├── ExcelDna.Integration.dll         # Excel-DNA 核心交互库
├── Microsoft.Web.WebView2.Core.dll  # WebView2 托管核心
├── Microsoft.Web.WebView2.WinForms.dll
├── runtimes/                        # 双架构原生加载器目录
│   ├── win-x64/native/WebView2Loader.dll
│   └── win-x86/native/WebView2Loader.dll
├── dist/                            # Svelte 5 前端生产编译产物 (SPA)
│   ├── index.html
│   └── assets/
├── 安装插件(开启常驻).bat            # 桌面快捷操作入口：永久写入 Excel 启动项
├── 卸载插件.bat                     # 桌面快捷操作入口：安全注销并清理
├── 免安装启动.bat                   # 临时演示入口：自动判断位数直接启动，不写注册表
├── core/                            # 自动化脚本引擎
│   ├── install_addin.ps1            # 核心安装逻辑（环境检测、解锁、注册）
│   ├── uninstall_addin.ps1          # 核心卸载逻辑（反注册、清理注册表）
│   └── launch_portable.ps1          # 便携启动逻辑（探测架构并挂载对应 XLL）
└── README_使用说明.txt              # 跨机器使用指南与常见问题排查
```

---

## 3. 关键技术细节与兼容机制

### 3.1 32位与64位 Excel 架构判定
传统通过操作系统位数判断容易误判（64位 Windows 上安装 32位 Office 极其普遍）。
核心检测算法：
- 定位系统当前默认关联的 `excel.exe` 路径（从 `HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe` 读取）；
- 直接读取 `excel.exe` 二进制文件的 PE 文件头（Machine 字段）：
  - `0x014c`：32位 (x86) 进程，自动加载 `LeeExcel.xll`；
  - `0x8664`：64位 (x64) 进程，自动加载 `LeeExcel64.xll`。

### 3.2 WebView2Loader 架构动态自适应 (C# 宿主增强)
在 `src/TaskPaneControl.cs` 中，初始化 WebView2 之前：
- 检测当前运行期环境架构（`IntPtr.Size == 8` 为 64位，`4` 为 32位）；
- 查找当前插件基目录下的 `runtimes\win-x64\native` 或 `runtimes\win-x86\native`；
- 通过 `CoreWebView2Environment.SetLoaderDllFolderPath` 指定对应的目录；若同级目录已存在 loader 则平滑回退，确保任何机型绝对不会发生 `DllNotFoundException`。

### 3.3 新老版本 Excel 注册表注入规范
Excel 加载项采用微软官方标准 `OPEN` 启动参数项进行注册：
- 目标路径：`HKCU\Software\Microsoft\Office\<Version>\Excel\Options`
  - 扫描所有已安装的 Office 主版本：`14.0` (2010), `15.0` (2013), `16.0` (2016/2019/2021/365)；
- 键值命名规则：
  - 查找是否已存在 `OPEN`、`OPEN1`、`OPEN2`...
  - 若已有其他插件，顺延追加（例如新键为 `OPEN3`）；
  - 键值内容：`"/R \"<完整物理路径>\\LeeExcel[64].xll\""`。
- 卸载时重排序机制：
  - 精确匹配包含 `LeeExcel` 的项并删除；
  - 将剩余的 `OPEN*` 键重新紧凑编号（Excel 遇到断号将中止后续插件加载，重排可确保不破坏用户原本的其他加载项）。

### 3.4 外来文件安全解除 (Unblock-File)
针对目标机可能存在的 Mark-of-the-Web 阻止机制：
- 安装与启动脚本执行时，对插件目录下所有 `.dll`、`.xll`、`.html`、`.js` 执行 `Unblock-File`（或消除 `:Zone.Identifier:$DATA` 备用数据流），防止 Office 弹窗阻止。

### 3.5 WebView2 运行时自动化检测
- 查询注册表：
  - `HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}`
  - `HKCU\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}`
- 若均不存在 `pv` 版本值，控制台提示：
  “检测到当前系统尚未安装 Microsoft Edge WebView2 运行时，正在为您启动自动下载安装（约2MB）...”
  通过 PowerShell 从微软官方 CDN 下载 `MicrosoftEdgeWebview2Setup.exe` 并执行 `/silent /install`，实现真正无感就绪。

---

## 4. 自动化打包流水线 (`scripts/package_release.ps1`)

打包脚本执行顺序：
1. **前端编译**：检查 Node.js / pnpm 环境，执行 `pnpm run build`，编译生成生产级 `dist/`；
2. **C# 插件编译**：调用系统 `csc.exe` 编译 `LeeExcel.dll`（Any CPU）；
3. **收集离线依赖**：
   - 提取 `packages/ExcelDna.AddIn.1.9.0/` 的 32位 与 64位 XLL；
   - 提取 32位 与 64位 `WebView2Loader.dll` 放入 `runtimes/` 目录；
   - 拷贝 `LeeExcel.dna`、`LeeExcel64.dna`、`Microsoft.Web.WebView2.*.dll`、`ExcelDna.Integration.dll`；
4. **生成桌面与核心脚本**：
   - 生成 `.bat` 引导脚本（防中文乱码与高权限申请机制）；
   - 生成 `core/` 下的安装、卸载与便携启动 PowerShell 脚本；
   - 生成格式精美的 `README_使用说明.txt`；
5. **归档与压缩**：
   - 按照 AGENTS.md 规范，将打包生成物输出到 `.artifacts/release/LeeExcel_Release_v1.0.zip`；
   - 校验生成包完整性，打印输出文件清单与体积。

---

## 5. 质量保证与交付门禁

1. 离线核心测试套件 (`test_suite_unit.cjs` 及 `test_regex_counter_example.cjs`) 保持 100% 通过；
2. 架构安全检查：在 Windows 64位系统上验证 32位 与 64位双文件的有效性；
3. 严格遵守 AGENTS.md：禁止将 `.artifacts/release/` 下的压缩包或临时编译物提交到 git 仓库。
