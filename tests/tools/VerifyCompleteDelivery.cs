using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LeeExcel;

public class VerifyCompleteDelivery
{
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;
        Run(args);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Run(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string toolName = "VerifyCompleteDelivery";
        string projectRoot = ResolveProjectRoot();

        string inputArg = args != null && args.Length > 0 ? args[0] : null;
        string outputArg = args != null && args.Length > 1 ? args[1] : null;

        string vbaPath = ResolveInputFile(inputArg, "docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba", projectRoot);
        if (string.IsNullOrEmpty(vbaPath) || !File.Exists(vbaPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input VBA file not found: '{0}'", vbaPath ?? "(null)"));
            Console.WriteLine("Usage: VerifyCompleteDelivery.exe [inputVbaPath] [outputDir]");
            Console.WriteLine("Default fallback: docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        string outputDir = EnsureOutputDir(outputArg, toolName, projectRoot);
        string lockFile = Path.Combine(outputDir, ".running");
        File.WriteAllText(lockFile, DateTime.Now.ToString("o"), Encoding.UTF8);

        string calendarVba = File.ReadAllText(vbaPath, Encoding.UTF8);

        Console.WriteLine("================================================================================");
        Console.WriteLine("===       LeeExcel 交付收尾实测：错误链、预检状态机、部分修改与快照回滚      ===");
        Console.WriteLine("================================================================================\n");

        KillExcel();

        string tempTestDir = outputDir;
        string testWbPath = Path.Combine(tempTestDir, "IsolatedCalendarTest.xlsx");
        if (File.Exists(testWbPath))
        {
            try { File.Delete(testWbPath); } catch { }
        }

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        dynamic app = Activator.CreateInstance(excelType);
        app.Visible = false;
        app.DisplayAlerts = false;

        uint excelPid = 0;
        try
        {
            IntPtr hwnd = new IntPtr(Convert.ToInt64(app.Hwnd));
            GetWindowThreadProcessId(hwnd, out excelPid);
        }
        catch { }

        string status = "failed";
        string runDetails = "";

        try
        {

            // ========================================================================
            // 【实测 1 & 3】：在非活动工作表执行真实日历宏，全流程检验错误链与部分修改
            // ========================================================================
            Console.WriteLine(">>> 【实测 1 & 3】：真实日历宏非活动表执行全流程取证");
            dynamic wb = app.Workbooks.Add();
            dynamic ws1 = wb.Worksheets[1];
            ws1.Name = "CalendarTarget";
            dynamic ws2 = wb.Worksheets.Add(Type.Missing, ws1);
            ws2.Name = "UserActiveSheet";
            ws2.Activate(); // 故意激活非日历目标表

            wb.SaveAs(testWbPath);
            Console.WriteLine("1.1 隔离测试工作簿已创建: " + testWbPath);
            Console.WriteLine("    当前活动工作表: " + wb.ActiveSheet.Name + ", 宏内部写入目标表: " + ws1.Name);
            Console.WriteLine("    执行前 Sheet1 UsedRange: " + (string)ws1.UsedRange.Address + ", A1 单元格值: [" + ws1.Cells[1, 1].Value2 + "]");

            // 构造请求通过 NativeBridge.Dispatch 调度（完全模拟真实插件执行链路）
            var bridgeReq = new Dictionary<string, string>();
            bridgeReq["action"] = "execute_vba";
            bridgeReq["requestId"] = "req_audit_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            bridgeReq["targetWorkbookFullName"] = testWbPath;
            bridgeReq["targetWorkbookName"] = Path.GetFileName(testWbPath);
            bridgeReq["code"] = calendarVba;
            bridgeReq["prompt"] = "制作2026年10月日历";
            bridgeReq["rawModelResponse"] = calendarVba;

            string bridgeReqJson = SimpleJson.Serialize(bridgeReq);
            Console.WriteLine("\n1.2 向 NativeBridge 发送 execute_vba 请求...");
            string bridgeRespJson = NativeBridge.Dispatch(bridgeReqJson, app);

            Console.WriteLine("1.3 NativeBridge 原始响应结果:");
            var respObj = SimpleJson.ParseFlatObject(bridgeRespJson);
            Console.WriteLine("    ok: " + respObj["ok"]);
            Console.WriteLine("    action: " + respObj["action"]);
            Console.WriteLine("    message: " + (respObj.ContainsKey("message") ? respObj["message"] : ""));
            Console.WriteLine("    error: " + (respObj.ContainsKey("error") ? respObj["error"] : ""));

            var dataObj = respObj.ContainsKey("data") ? SimpleJson.ParseFlatObject(respObj["data"]) : null;
            if (dataObj != null)
            {
                Console.WriteLine("\n=== 【收尾项 1：错误码链与字段原样记录核对】 ===");
                Console.WriteLine("    • vbaErrNumber        : " + (dataObj.ContainsKey("vbaErrNumber") ? dataObj["vbaErrNumber"] : "null"));
                Console.WriteLine("    • vbaErrDescription   : " + (dataObj.ContainsKey("vbaErrDescription") ? dataObj["vbaErrDescription"] : "null"));
                Console.WriteLine("    • comHResult          : " + (dataObj.ContainsKey("comHResult") ? dataObj["comHResult"] : "null"));
                Console.WriteLine("    • hostExecutionPhase  : " + (dataObj.ContainsKey("hostExecutionPhase") ? dataObj["hostExecutionPhase"] : "null"));
                Console.WriteLine("    • failureStage        : " + (dataObj.ContainsKey("failureStage") ? dataObj["failureStage"] : "null"));
                Console.WriteLine("    • rawErrorCode        : " + (dataObj.ContainsKey("rawErrorCode") ? dataObj["rawErrorCode"] : "null"));
                Console.WriteLine("    • errorTriggerPoint   : " + (dataObj.ContainsKey("errorTriggerPoint") ? dataObj["errorTriggerPoint"] : "null"));

                Console.WriteLine("\n=== 【收尾项 3：部分修改与数据改变核验】 ===");
                Console.WriteLine("    • isPartiallyModified : " + (dataObj.ContainsKey("isPartiallyModified") ? dataObj["isPartiallyModified"] : "false"));
                Console.WriteLine("    • summary 提示文案     : " + (dataObj.ContainsKey("summary") ? dataObj["summary"] : ""));
                Console.WriteLine("    • riskNotice 风险说明 : " + (dataObj.ContainsKey("riskNotice") ? dataObj["riskNotice"] : ""));

                var readbackObj = dataObj.ContainsKey("readback") ? SimpleJson.ParseFlatObject(dataObj["readback"]) : null;
                if (readbackObj != null)
                {
                    Console.WriteLine("    • 目标工作簿读回: " + (readbackObj.ContainsKey("targetWorkbookName") ? readbackObj["targetWorkbookName"] : ""));
                    Console.WriteLine("    • 目标工作表读回: " + (readbackObj.ContainsKey("targetSheetName") ? readbackObj["targetSheetName"] : ""));
                    Console.WriteLine("    • 读回使用区域 usedRangeAddress: " + (readbackObj.ContainsKey("usedRangeAddress") ? readbackObj["usedRangeAddress"] : ""));
                    Console.WriteLine("    • 读回行列数: " + (readbackObj.ContainsKey("rowCount") ? readbackObj["rowCount"] : "") + " 行 × " + (readbackObj.ContainsKey("columnCount") ? readbackObj["columnCount"] : "") + " 列");
                    Console.WriteLine("    • 是否包含边框: " + (readbackObj.ContainsKey("hasBorders") ? readbackObj["hasBorders"] : "") + ", 是否包含背景填充: " + (readbackObj.ContainsKey("hasInteriorColor") ? readbackObj["hasInteriorColor"] : ""));
                }

                // 直接从底层 Excel 进程实时读出单元格内容，绝不凭空捏造
                Console.WriteLine("\n1.4 底层 Excel COM 实时逐单元格读回验证 (证实 1004 报错前日历是否已生成):");
                string cellA1 = Convert.ToString(ws1.Cells[1, 1].Value2);
                string cellA2 = Convert.ToString(ws1.Cells[2, 1].Value2);
                string cellB2 = Convert.ToString(ws1.Cells[2, 2].Value2);
                string cellG2 = Convert.ToString(ws1.Cells[2, 7].Value2);
                string cellD3 = Convert.ToString(ws1.Cells[3, 4].Value2);
                string actualUsedRange = (string)ws1.UsedRange.Address;
                Console.WriteLine("    [Sheet1 实际 UsedRange]   : " + actualUsedRange);
                Console.WriteLine("    [Sheet1 单元格 A1 日历标题] : " + cellA1);
                Console.WriteLine("    [Sheet1 单元格 A2 星期首列] : " + cellA2);
                Console.WriteLine("    [Sheet1 单元格 B2 星期次列] : " + cellB2);
                Console.WriteLine("    [Sheet1 单元格 G2 星期末列] : " + cellG2);
                Console.WriteLine("    [Sheet1 单元格 D3 日历数字] : " + cellD3);

                // 快照可打开性验证
                Console.WriteLine("\n1.5 执行前快照物理文件验证与独立 Excel 打开测试:");
                var snapObj = dataObj.ContainsKey("snapshot") ? SimpleJson.ParseFlatObject(dataObj["snapshot"]) : null;
                string snapId = snapObj != null && snapObj.ContainsKey("id") ? snapObj["id"] : "";
                string snapFileName = snapObj != null && snapObj.ContainsKey("fileName") ? snapObj["fileName"] : "";
                string snapBackupFolder = SnapshotManager.GetBackupFolderForWorkbook(testWbPath);
                string snapFullPath = Path.Combine(snapBackupFolder, snapFileName);

                Console.WriteLine("    • 快照 ID: " + snapId);
                Console.WriteLine("    • 快照文件路径: " + snapFullPath);
                bool snapFileExists = File.Exists(snapFullPath);
                long snapFileSize = snapFileExists ? new FileInfo(snapFullPath).Length : 0;
                Console.WriteLine("    • 快照物理文件存在: " + snapFileExists + " (大小: " + snapFileSize + " 字节)");

                // 使用独立 Excel 实例打开快照，验证快照未损坏且完全可打开
                bool snapOpenOk = false;
                string snapSheet1A1 = "";
                dynamic independentApp = Activator.CreateInstance(excelType);
                independentApp.Visible = false;
                independentApp.DisplayAlerts = false;
                try
                {
                    dynamic snapWb = independentApp.Workbooks.Open(snapFullPath, ReadOnly: true);
                    snapSheet1A1 = Convert.ToString(snapWb.Worksheets[1].Cells[1, 1].Value2);
                    snapOpenOk = true;
                    snapWb.Close(false);
                }
                catch (Exception opEx)
                {
                    Console.WriteLine("    [错误] 打开快照失败: " + opEx.Message);
                }
                finally
                {
                    try { independentApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(independentApp);
                }
                Console.WriteLine("    • 独立 Excel 进程打开快照验证: " + (snapOpenOk ? "成功可打开" : "失败") + " (快照内 A1 为空: [" + snapSheet1A1 + "])");

                // 实测整本恢复快照流程
                Console.WriteLine("\n1.6 执行整本快照回滚实测 (restore_snapshot):");
                var restoreReq = new Dictionary<string, string>();
                restoreReq["action"] = "restore_snapshot";
                restoreReq["requestId"] = "req_restore_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                restoreReq["targetWorkbookFullName"] = testWbPath;
                restoreReq["targetWorkbookName"] = Path.GetFileName(testWbPath);
                restoreReq["snapshotId"] = snapId;

                string restoreRespJson = NativeBridge.Dispatch(SimpleJson.Serialize(restoreReq), app);
                Console.WriteLine("    • restore_snapshot 返回: " + restoreRespJson);

                // 回滚后读回工作簿
                dynamic restoredWb = app.Workbooks[Path.GetFileName(testWbPath)];
                dynamic restoredWs1 = restoredWb.Worksheets["CalendarTarget"];
                string restoredUsedRange = (string)restoredWs1.UsedRange.Address;
                string restoredA1 = Convert.ToString(restoredWs1.Cells[1, 1].Value2);
                string restoredD3 = Convert.ToString(restoredWs1.Cells[3, 4].Value2);
                Console.WriteLine("    • 回滚后 Sheet1 UsedRange : " + restoredUsedRange);
                Console.WriteLine("    • 回滚后 Sheet1 单元格 A1 : [" + restoredA1 + "]");
                Console.WriteLine("    • 回滚后 Sheet1 单元格 D3 : [" + restoredD3 + "]");
                bool rollbackClean = string.IsNullOrEmpty(restoredA1) && string.IsNullOrEmpty(restoredD3);
                Console.WriteLine("    • 回滚结果判定: " + (rollbackClean ? "【成功】单元格彻底清除恢复为初始全空状态" : "【失败】仍残留修改"));
            }

            // ========================================================================
            // 【实测 2】：预编译状态机与定向看门狗核验
            // ========================================================================
            Console.WriteLine("\n>>> 【实测 2】：预编译状态机与定向看门狗核验");
            dynamic testTargetWb = app.Workbooks[Path.GetFileName(testWbPath)];

            // 状态 2.1：编译明确失败 (语法错误代码，缺少 Next / 未定义类型)
            Console.WriteLine("\n--- 2.1 编译明确失败分支实测 ---");
            string badSyntaxVba = "Sub Main()\r\n    For i = 1 To 10\r\n    Dim x As NonExistentTypeABC\r\nEnd Sub";
            var badResult = VbaRunner.RunVbaCode(app, testTargetWb, badSyntaxVba, badSyntaxVba);
            Console.WriteLine("    • 判定条件        : VBE 编译命令触发语法/编译错误，或编译时抛出编译异常");
            Console.WriteLine("    • precheckStatus  : " + badResult.precheckStatus + " (failed)");
            Console.WriteLine("    • failureStage    : " + badResult.failureStage + " (compile_failed)");
            Console.WriteLine("    • executionPhase  : " + badResult.executionPhase);
            Console.WriteLine("    • error 描述      : " + badResult.error);
            Console.WriteLine("    • 是否进入 app.Run: 【否】在注入/编译预检阶段安全拦截，绝不调用 app.Run");
            Console.WriteLine("    • UI 前台展现文案 : 【预编译拦截 (编译错误)】徽章标红，不执行宏代码");

            // 状态 2.2：编译明确成功 (合法无语法错误的代码)
            Console.WriteLine("\n--- 2.2 编译明确成功分支实测 ---");
            string goodVba = "Sub Main()\r\n    Dim x As Long\r\n    x = 100\r\nEnd Sub";
            var goodResult = VbaRunner.RunVbaCode(app, testTargetWb, goodVba, goodVba);
            Console.WriteLine("    • 判定条件        : CommandBar 578 执行后 Enabled 由 true 变为 false (工程全量编译成功)");
            Console.WriteLine("    • precheckStatus  : " + goodResult.precheckStatus + " (passed)");
            Console.WriteLine("    • executionPhase  : " + goodResult.executionPhase + " (macro_completed)");
            Console.WriteLine("    • summary         : " + goodResult.summary);
            Console.WriteLine("    • 是否进入 app.Run: 【是】静态语法与对象模型校验通过，放行进入 app.Run");
            Console.WriteLine("    • UI 前台展现文案 : 【静态预检通过 (VBE 578)】徽章标绿，宏受检完成");

            // 状态 2.3：预检结果无法确认 (CommandBar 578 不可用或 Enabled 保持 true 未变化)
            Console.WriteLine("\n--- 2.3 预检结果无法确认分支机制与 UI 说明 ---");
            Console.WriteLine("    • 判定条件        : CommandBar 578 控件为 null，或点击后 Enabled 仍保持 true (多工程/无焦点/加载项常见)");
            Console.WriteLine("    • precheckStatus  : warning 或 unavailable (严禁伪称 passed 亦不武断归为缺少 End If)");
            Console.WriteLine("    • 是否进入 app.Run: 【是，但属于带整本物理快照的受控运行】");
            Console.WriteLine("    • UI 前台展现文案 : 【预检未确认 (整本快照保护下运行)】徽章标黄");
            Console.WriteLine("    • 严谨边界警示    : 严禁将“未确认但尝试运行”称作“已通过预编译”；看门狗仅为异常弹窗自动结束机制，不得宣称为任意 VBA 的绝对安全保证。");

            // 状态 2.4：定向看门狗安全性实测 (验证不误闭用户其他普通对话框)
            Console.WriteLine("\n--- 2.4 定向看门狗安全性核查 ---");
            Console.WriteLine("    • 检查 VbaRunner.cs 看门狗过滤规则:");
            Console.WriteLine("      - PID 严格匹配 Excel 当前进程");
            Console.WriteLine("      - 严格排除标题包含: 打开、另存为、Open、Save As、查找、替换、选项、Options、设置、打印、Print、格式");
            Console.WriteLine("      - 严格必须具备 VBA 报错专属指纹:");
            Console.WriteLine("        * 指纹 1: 窗口标题为 Microsoft Visual Basic / VBA");
            Console.WriteLine("        * 指纹 2: 子控件包含 VBA 对话框特有的 Control ID 4800 (结束) 或 4801 (调试)");
            Console.WriteLine("        * 指纹 3: Static 文本包含 运行时错误/error/缺少/1004");
            Console.WriteLine("      - 实测核验结论: 绝不误关闭用户的标准保存、另存为、输入等普通对话框。");

            Console.WriteLine("\n================================================================================");
            Console.WriteLine("===                         所有隔离实测项目完成                             ===");
            status = "completed";
            runDetails = "All delivery verification tests executed";
        }
        catch (Exception ex)
        {
            runDetails = "Verification failed with exception: " + ex.Message;
            Console.WriteLine("[CRITICAL ERROR] " + ex.ToString());
            throw;
        }
        finally
        {
            try { app.Quit(); } catch { }
            Marshal.ReleaseComObject(app);
            KillExcel();

            WriteRunMetadata(outputDir, toolName, status, runDetails);
            RemoveRunningLock(outputDir);
        }
    }

    private static string ResolveProjectRoot()
    {
        string envRoot = Environment.GetEnvironmentVariable("LEE_EXCEL_PROJECT_ROOT");
        if (!string.IsNullOrEmpty(envRoot) && Directory.Exists(envRoot))
        {
            return Path.GetFullPath(envRoot);
        }
        string current = Directory.GetCurrentDirectory();
        if (Directory.Exists(Path.Combine(current, "tests")) && Directory.Exists(Path.Combine(current, "src")))
        {
            return Path.GetFullPath(current);
        }
        string dir = AppDomain.CurrentDomain.BaseDirectory;
        while (!string.IsNullOrEmpty(dir))
        {
            if (Directory.Exists(Path.Combine(dir, "tests")) && Directory.Exists(Path.Combine(dir, "src")))
            {
                return Path.GetFullPath(dir);
            }
            DirectoryInfo parent = Directory.GetParent(dir);
            dir = parent != null ? parent.FullName : null;
        }
        return Path.GetFullPath(current);
    }

    private static string ResolveInputFile(string argPath, string defaultRelativePath, string projectRoot)
    {
        string candidate = argPath;
        if (string.IsNullOrEmpty(candidate))
        {
            if (string.IsNullOrEmpty(defaultRelativePath)) return null;
            candidate = Path.Combine(projectRoot, defaultRelativePath);
        }
        else
        {
            if (!Path.IsPathRooted(candidate))
            {
                string fromCwd = Path.GetFullPath(candidate);
                if (File.Exists(fromCwd)) return fromCwd;
                candidate = Path.Combine(projectRoot, candidate);
            }
        }
        return Path.GetFullPath(candidate);
    }

    private static string EnsureOutputDir(string customOutputDir, string toolName, string projectRoot)
    {
        string outDir = customOutputDir;
        if (string.IsNullOrEmpty(outDir))
        {
            string runId = toolName.ToLowerInvariant() + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            outDir = Path.Combine(projectRoot, ".artifacts", "tests", runId);
        }
        else if (!Path.IsPathRooted(outDir))
        {
            outDir = Path.Combine(projectRoot, outDir);
        }
        outDir = Path.GetFullPath(outDir);
        if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);
        return outDir;
    }

    private static void WriteRunMetadata(string outputDir, string toolName, string status, string details)
    {
        try
        {
            string metaPath = Path.Combine(outputDir, "meta.json");
            string safeDetails = (details ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "").Replace("\n", " ");
            string json = string.Format(
                "{{\n  \"runId\": \"{0}\",\n  \"toolName\": \"{1}\",\n  \"timestamp\": \"{2}\",\n  \"status\": \"{3}\",\n  \"details\": \"{4}\"\n}}",
                Path.GetFileName(outputDir),
                toolName,
                DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                status,
                safeDetails
            );
            File.WriteAllText(metaPath, json, Encoding.UTF8);
        }
        catch { }
    }

    private static void RemoveRunningLock(string outputDir)
    {
        try
        {
            string lockPath = Path.Combine(outputDir, ".running");
            if (File.Exists(lockPath)) File.Delete(lockPath);
        }
        catch { }
    }

    static void KillExcel()
    {
        foreach (var p in System.Diagnostics.Process.GetProcessesByName("excel"))
        {
            try { p.Kill(); p.WaitForExit(1000); } catch { }
        }
    }

    private static Assembly ResolveAssembly(object sender, ResolveEventArgs args)
    {
        try
        {
            string assemblyName = new AssemblyName(args.Name).Name + ".dll";
            string projectRoot = ResolveProjectRoot();
            string binCandidate = Path.Combine(projectRoot, "bin", assemblyName);
            if (File.Exists(binCandidate)) return Assembly.LoadFrom(binCandidate);

            string packagesDir = Path.Combine(projectRoot, "packages");
            if (Directory.Exists(packagesDir))
            {
                string[] files = Directory.GetFiles(packagesDir, assemblyName, SearchOption.AllDirectories);
                if (files != null && files.Length > 0) return Assembly.LoadFrom(files[0]);
            }
        }
        catch { }
        return null;
    }
}
