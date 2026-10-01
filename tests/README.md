# Lee-Excel 测试与诊断资产手册 (tests/README.md)

本目录收纳 Lee-Excel 长期维护的测试工具源码（C#）及诊断脚本（PowerShell/Node）。

---

## 统一编译与运行准则

1. **编译器要求**：系统自带 64 位 .NET Framework 4.0/4.6+ 编译器：`C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`。
2. **生成物隔离禁令**：严禁编译生成到 `tests/` 或 `bin/` 目录！必须显式传入 `/out:".artifacts\tests\<run-id>\<name>.exe"`。
3. **输出隔离禁令**：工具运行生成的临时工作簿、抓包日志、内存转储必须输出至 `.artifacts/tests/<run-id>/`。

---

## 一、 测试工具集合 (`tests/tools/`)

| 工具文件名 | 具体用途 | 编译依赖 | 推荐编译命令 | 推荐运行方式 | 默认输出位置 | 副作用与环境限制 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`FindDialogButtons.cs`** | 遍历查找 Excel 弹窗中的所有控件按钮句柄与文本 | `user32.dll` (Win32) | `csc /nologo /target:exe /out:".artifacts\tests\run\FindDialogButtons.exe" tests\tools\FindDialogButtons.cs` | `.\FindDialogButtons.exe [vbaPath] [outDir]` | 控制台标准输出 + `.artifacts/tests/finddialogbuttons_<time>_<id>/` (生成 `meta.json` 与 `.running`) | 枚举系统顶级窗口与子控件；无输入时在 COM 调用前退出；默认回退到 `docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba` |
| **`ReadDialogText.cs`** | 读取特定 Excel 对话框的详细提示文字与控件标题 | `user32.dll` (Win32) | `csc /nologo /target:exe /out:".artifacts\tests\run\ReadDialogText.exe" tests\tools\ReadDialogText.cs` | `.\ReadDialogText.exe [vbaPath] [outDir]` | 控制台标准输出 + `.artifacts/tests/readdialogtext_<time>_<id>/` (落盘 `dialog_content.txt`、`meta.json`) | 读取前台窗口文本；无输入时在 COM 调用前退出；默认回退到 `docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba` |
| **`RunTaskPaneFrontEndRegression.cs`** | 驱动 Excel 任务窗格前端 WebView2 导航并进行端到端回归校验 | .NET 4.0+ / Excel Interop / WebView2 | `csc /nologo /target:exe /out:".artifacts\tests\run\RunTaskPaneFrontEndRegression.exe" /r:bin\LeeExcel.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll /r:packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll tests\tools\RunTaskPaneFrontEndRegression.cs` | `.\RunTaskPaneFrontEndRegression.exe [rawModelVba] [vbaCode] [outDir]` | `.artifacts/tests/runtaskpanefrontendregression_<time>_<id>/` (落盘 `Calendar_FrontEnd_Test.xlsx`、`meta.json`) | **需运行中 Excel 实例**；前置输入缺失直接退出；默认分别回退到 `docs/history/evidence_202609/` 的 Stage1 原文与 Stage2 提取源码 |
| **`RegressionTest.cs`** | 核心端到端回归：自动启动 Excel、注入宏代码、校验执行状态与隔离保护 | Excel Interop COM / LeeExcel.dll | `csc /nologo /target:exe /out:".artifacts\tests\run\RegressionTest.exe" /r:bin\LeeExcel.dll tests\tools\RegressionTest.cs` | `.\RegressionTest.exe [vbaPath] [outDir]` | 控制台标准输出 + `.artifacts/tests/regressiontest_<time>_<id>/` (生成 `meta.json`) | **启动 Excel COM 实例**；前置输入缺失直接退出；默认回退到 `docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba` |
| **`TestThreeConditions.cs`** | 验证三条件准则：①无语法弹窗 ②工作簿定向准确 ③宿主状态恢复 | Excel Interop COM | `csc /nologo /target:exe /out:".artifacts\tests\run\TestThreeConditions.exe" tests\tools\TestThreeConditions.cs` | `.\TestThreeConditions.exe [vbaPath] [outDir]` | 控制台标准输出 + `.artifacts/tests/testthreeconditions_<time>_<id>/` (生成 `meta.json`) | **启动 Excel COM 实例**；前置输入缺失直接退出；默认回退到 `docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba` |
| **`TestWbBinding.cs`** | 专项验证多工作簿并发下的 `Application.Run` 与 Workbook 传参绑定机制 | Excel Interop COM | `csc /nologo /target:exe /out:".artifacts\tests\run\TestWbBinding.exe" tests\tools\TestWbBinding.cs` | `.\TestWbBinding.exe [outputDir]` | 指定目录或 `.artifacts/tests/wb_binding_<time>/` | **启动 Excel COM 实例**；生成临时工作簿与日志 |
| **`VerifyCompleteDelivery.cs`** | 交付链路完整性校验工具：校验前置条件、阶段交付包及执行流 | Excel Interop COM / LeeExcel.dll | `csc /nologo /target:exe /out:".artifacts\tests\run\VerifyCompleteDelivery.exe" /r:bin\LeeExcel.dll tests\tools\VerifyCompleteDelivery.cs` | `.\VerifyCompleteDelivery.exe [vbaPath] [outDir]` | `.artifacts/tests/verifycompletedelivery_<time>_<id>/` (落盘 `IsolatedCalendarTest.xlsx`、`meta.json`) | **操作 Excel COM 实例**；前置输入缺失直接退出；彻底移除向 `scratch/` 写入行为；默认回退到 `docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba` |
| **`VerifyPayload.cs`** | 离线解析检验模型下发的 Payload JSON，校验 Stage1~3 代码拆解与哈希匹配 | `System.Web.Extensions.dll` | `csc /nologo /target:exe /r:System.Web.Extensions.dll /out:".artifacts\tests\run\VerifyPayload.exe" tests\tools\VerifyPayload.cs` | `.\VerifyPayload.exe <inputJson> [outputDir]` | 指定目录或 `.artifacts/tests/payload_verification/` | 离线纯文本解析，无网络或COM副作用 |
| **`MemoryDump.cs`** | 扫描特定进程内存特征码（如 VBA 宏切片、未捕获异常关键字），提取证据片段 | `kernel32.dll` (Win32) | `csc /nologo /target:exe /out:".artifacts\tests\run\MemoryDump.exe" tests\tools\MemoryDump.cs` | `.\MemoryDump.exe [outputDir]` | 指定目录或 `.artifacts/tests/memory_dumps/` | **调用 OpenProcess 与 ReadProcessMemory**；读取目标进程内存并写文件 |

---

## 二、 诊断与排查脚本 (`tests/diagnostics/`)

| 脚本文件名 | 具体用途 | 运行环境 | 运行命令 | 输出位置 | 副作用与环境限制 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`enum_excel_wins.ps1`** | 枚举当前 Windows 桌面所有 Excel 相关的主窗口与子窗口 HWND | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/enum_excel_wins.ps1` | 控制台输出 | 只读枚举，无副作用 |
| **`get_windows.ps1`** | 打印当前桌面活跃窗口的进程名、PID、类名与标题 | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/get_windows.ps1` | 控制台输出 | 只读枚举，无副作用 |
| **`spy_window.ps1`** | 注入语法错误宏并启动后台监听线程，排查 `#32770` 弹窗生成与关闭响应 | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/spy_window.ps1` | 控制台输出 | **启动后台 Excel COM 实例**；使用 Win32 消息关闭弹窗 |
| **`read_vbe_dialog.ps1`** | 触发 VBE 错误弹窗并读取对话框中的错误提示内容 | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/read_vbe_dialog.ps1 -WorkbookPath <xlsx> -VbaSnippetPath <vba>` | 控制台输出 + 指定/默认 `.artifacts/tests/` 目录 (`dialog_result.txt`、`meta.json`) | **强制显式参数**；参数缺失在任何 COM 操作前报错退出；彻底移除桌面通配扫描与废弃的历史样本硬编码 |
| **`probe_dialog.ps1`** | 异步触发不完整 If 语句，探测并自动销毁语法错误弹窗 | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/probe_dialog.ps1` | 控制台输出 | **启动后台 Excel COM 实例**并启动后台 Job |
| **`probe_model.cjs`** | **联网模型探针**：测试上游大模型 API 对 `max_tokens` 与 `reasoning_content` 的支持 | Node.js 16+ | `node tests/diagnostics/probe_model.cjs` | 控制台输出 | **联网模型探针：直接发起真实 HTTPS 请求并消耗 API Token 额度！严禁自动运行！** |
| **`read_config.ps1`** | 读取本机 WebView2 本地 LevelDB 存储的配置，脱敏输出 API Key 与 BaseUrl | PowerShell 5.1+ | `powershell -NoProfile -File tests/diagnostics/read_config.ps1` | 控制台脱敏输出 | 只读读取本地用户数据，自动脱敏敏感 Key，无网络请求 |
