# Lee-Excel 跨机器分发与全版本 Excel 兼容实施计划

> **历史资料，不代表当前版本**  
> **归档日期**：2026-10-01（对应基线 v1.1.0/v1.2.0）  
> **当前入口/替代文档**：[docs/product-roadmap.md](file:///c:/Users/35651/Desktop/Google/lee-excle/docs/product-roadmap.md)、[docs/CURRENT_STATE.md](file:///c:/Users/35651/Desktop/Google/lee-excle/docs/CURRENT_STATE.md)  
>
> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建 Lee-Excel 跨机器全自动打包流水线与分发套件，实现 32位/64位 Excel 与 Office 2010~365 全版本兼容，开箱即用。

**Architecture:** 双架构（x86/x64）加载项同包组织 + C# 宿主动态适配加载器 + PowerShell/Bat 智能环境探测与注册表注入引擎 + 一键打包流水线脚本。

**Tech Stack:** C# (.NET 4.6.2+ / WinForms / WebView2), Excel-DNA 1.9.0, Svelte 5, PowerShell 5.1+, Windows Registry, PE Header Parser.

**Spec:** `docs/superpowers/specs/2026-10-01-cross-machine-packaging-design.md`

## Global Constraints

- **产物隔离规范**：临时与构建输出仅允许输出至 `.artifacts/release/`，严禁污染 `bin/` 以外的根目录，严禁提交生成物到 Git；
- **测试门禁要求**：每次修改代码后必须通过现有离线门禁 `node test_suite_unit.cjs` 与 `node test_regex_counter_example.cjs`，保持 100% PASS；
- **Office 进程安全**：严禁在未确认情况下强行杀死用户正在使用的真实 Excel 进程；
- **路径解析**：所有脚本使用 `$PSScriptRoot` 动态解析相对路径，不假定当前工作目录。

---

### Task 1: C# 宿主增强：WebView2Loader 双架构自适应加载

**Files:**
- Modify: `src/TaskPaneControl.cs`
- Test: `test_suite_unit.cjs`, `build_addin.ps1`

**Interfaces:**
- Consumes: `ExcelDnaUtil.XllPath`, `Environment.Is64BitProcess` / `IntPtr.Size`
- Produces: 动态根据进程架构定位 `runtimes\win-x64\native` 或 `runtimes\win-x86\native` 目录下的 `WebView2Loader.dll` 并注入环境路径，确保 32 位与 64 位 Excel 均能安全加载 WebView2。

- [ ] **Step 1: 检查现有门禁基线**
Run: `node test_suite_unit.cjs`
Expected: 27/27 PASS

- [ ] **Step 2: 增强 TaskPaneControl.cs 中的 WebView2Loader 探测逻辑**
在 `src/TaskPaneControl.cs` 中的 `InitializeWebView()` 方法内，创建 `CoreWebView2Environment` 之前，加入对 `runtimes/win-x64/native` 与 `runtimes/win-x86/native` 的目录探测，通过 `CoreWebView2Environment.SetLoaderDllFolderPath` 指定对应的原生加载器路径。

```csharp
// 探测双架构 WebView2Loader.dll
try
{
    string archFolder = (IntPtr.Size == 8) ? "win-x64" : "win-x86";
    string nativeLoaderDir = Path.Combine(baseDir, "runtimes", archFolder, "native");
    if (Directory.Exists(nativeLoaderDir) && File.Exists(Path.Combine(nativeLoaderDir, "WebView2Loader.dll")))
    {
        CoreWebView2Environment.SetLoaderDllFolderPath(nativeLoaderDir);
    }
}
catch (Exception loaderEx)
{
    System.Diagnostics.Debug.WriteLine("SetLoaderDllFolderPath info: " + loaderEx.Message);
}
```

- [ ] **Step 3: 运行 C# 增量编译并验证**
Run: `powershell -ExecutionPolicy Bypass -File .\build_addin.ps1`
Expected: 编译成功: `bin\LeeExcel.dll`

- [ ] **Step 4: 运行离线门禁**
Run: `node test_suite_unit.cjs`
Expected: PASS

---

### Task 2: 编写发布包安装、卸载与便携启动核心引擎

**Files:**
- Create: `scripts/core/install_addin.ps1`
- Create: `scripts/core/uninstall_addin.ps1`
- Create: `scripts/core/launch_portable.ps1`
- Create: `scripts/core/bat_templates/install.bat`
- Create: `scripts/core/bat_templates/uninstall.bat`
- Create: `scripts/core/bat_templates/portable.bat`

**Interfaces:**
- Consumes: 注册表 `HKCU\Software\Microsoft\Office\<ver>\Excel\Options`, `HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`
- Produces: 
  - `install_addin.ps1`：环境检测、Unblock-File 解锁、PE 解析位数、注册表 OPEN 注册、WebView2 引导
  - `uninstall_addin.ps1`：安全反注册与注册表 OPEN 重新排序紧凑化
  - `launch_portable.ps1`：自动识别架构挂载对应 XLL 启动，不修改注册表

- [ ] **Step 1: 编写 `scripts/core/install_addin.ps1`**
实现功能：
1. 递归解除当前目录下所有文件的 Mark-of-the-Web (`Unblock-File`)；
2. 检测 WebView2 注册表项，缺失时提示并引导使用官方 Bootstrapper 静默安装；
3. 解析 `excel.exe` PE 头判定 32 位还是 64 位；
4. 扫描 Office 14.0/15.0/16.0 注册表，将对应位数的 XLL 路径写入 `Excel\Options\OPEN*`。

- [ ] **Step 2: 编写 `scripts/core/uninstall_addin.ps1`**
实现功能：
1. 扫描 Office 14.0/15.0/16.0 注册表下的 `OPEN*` 键值；
2. 识别并删除所有指向 `LeeExcel*.xll` 的条目；
3. 将剩余的 `OPEN*` 项按顺序紧凑重排，避免序号断号导致 Excel 中断加载其它插件。

- [ ] **Step 3: 编写 `scripts/core/launch_portable.ps1`**
实现功能：
1. 解锁当前目录文件；
2. 解析当前系统关联的 `excel.exe` 位数；
3. 直接调用 `excel.exe "<path_to_xll>"` 启动，不向注册表写入任何内容。

- [ ] **Step 4: 准备对应的一键引导 `.bat` 文件**
配置 `@echo off`，设置 UTF-8 编码并调用隐藏窗口的 PowerShell 脚本，确保用户在任何普通终端或双击时无乱码并顺畅执行。

---

### Task 3: 编写发布包使用说明与全自动打包流水线脚本

**Files:**
- Create: `scripts/package_release.ps1`
- Create: `scripts/core/README_template.txt`
- Modify: `build_addin.ps1` (同步支持输出 32 位与 64 位双架构基础文件)

**Interfaces:**
- Consumes: `web/package.json` (Vite build), `packages/ExcelDna.AddIn.1.9.0/`, `packages/Microsoft.Web.WebView2.1.0.4191.47/`
- Produces: `.artifacts/release/LeeExcel_Release_v1.0.zip`

- [ ] **Step 1: 编写 `scripts/core/README_template.txt`**
包含：插件简介、一键安装说明、便携启动说明、新老版本 Excel 兼容性说明、WebView2 离线处理指引、卸载说明。

- [ ] **Step 2: 编写 `scripts/package_release.ps1`**
实现流水线：
1. 执行前端生产打包 `pnpm --dir web run build`；
2. 编译 C# 主库为 `LeeExcel.dll`；
3. 创建临时打包暂存区 `.artifacts/tmp/pack_stage/`；
4. 组装 32 位 `LeeExcel.xll` 与 64 位 `LeeExcel64.xll`；
5. 复制 `runtimes/win-x64/native/WebView2Loader.dll` 与 `runtimes/win-x86/native/WebView2Loader.dll`；
6. 复制前端 `dist/` 与 `.dna`、托管 DLL；
7. 复制 `install.bat`、`uninstall.bat`、`portable.bat`、`core/` 脚本及 `README_使用说明.txt`；
8. 压缩为 `.artifacts/release/LeeExcel_Release_v1.0.zip`；
9. 清理临时暂存区，并对生成文件校验摘要与体积。

- [ ] **Step 3: 升级 `build_addin.ps1`**
在日常开发构建脚本中同时保留 32 位与 64 位双架构完整支持，确保本地开发测试也具备双架构基础。

---

### Task 4: 端到端打包验证与门禁核验

**Files:**
- Test: `test_suite_unit.cjs`, `test_regex_counter_example.cjs`
- Output: `.artifacts/release/LeeExcel_Release_v1.0.zip`

- [ ] **Step 1: 运行核心离线门禁**
Run: `node test_suite_unit.cjs`
Run: `node test_regex_counter_example.cjs`
Expected: 100% PASS

- [ ] **Step 2: 运行打包流水线**
Run: `powershell -ExecutionPolicy Bypass -File .\scripts\package_release.ps1`
Expected: 成功生成 `.artifacts/release/LeeExcel_Release_v1.0.zip`

- [ ] **Step 3: 校验压缩包内部结构完整性**
解包验证是否包含：
- `LeeExcel.xll` (32位)
- `LeeExcel64.xll` (64位)
- `LeeExcel.dll`
- `ExcelDna.Integration.dll`
- `Microsoft.Web.WebView2.*.dll`
- `runtimes/win-x64/native/WebView2Loader.dll`
- `runtimes/win-x86/native/WebView2Loader.dll`
- `dist/index.html` 及静态文件
- `安装插件(开启常驻).bat`、`卸载插件.bat`、`免安装启动.bat`
- `core/install_addin.ps1`、`core/uninstall_addin.ps1`、`core/launch_portable.ps1`
- `README_使用说明.txt`

- [ ] **Step 4: 验证规范合规性**
确保 `bin/` 与根目录未产生多余污染，临时产物已安全归档至 `.artifacts/`，符合 AGENTS.md 守则。
