using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public class TestThreeConditions
{
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_COMMAND = 0x0111;

    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string toolName = "TestThreeConditions";
        string projectRoot = ResolveProjectRoot();

        string inputArg = args != null && args.Length > 0 ? args[0] : null;
        string outputArg = args != null && args.Length > 1 ? args[1] : null;

        string vbaPath = ResolveInputFile(inputArg, "docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba", projectRoot);
        if (string.IsNullOrEmpty(vbaPath) || !File.Exists(vbaPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input VBA file not found: '{0}'", vbaPath ?? "(null)"));
            Console.WriteLine("Usage: TestThreeConditions.exe [inputVbaPath] [outputDir]");
            Console.WriteLine("Default fallback: docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        string outputDir = EnsureOutputDir(outputArg, toolName, projectRoot);
        string lockFile = Path.Combine(outputDir, ".running");
        File.WriteAllText(lockFile, DateTime.Now.ToString("o"), Encoding.UTF8);

        string vbaWithSelect = File.ReadAllText(vbaPath, Encoding.UTF8);
        string vbaWithoutSelect = vbaWithSelect.Replace("ws.Cells(1, 1).Select", "' [Removed ws.Cells(1, 1).Select]");

        KillExcelProcesses();

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        dynamic app = Activator.CreateInstance(excelType);
        app.Visible = false;
        app.DisplayAlerts = false;

        uint excelPid = 0;
        try
        {
            IntPtr excelHwnd = new IntPtr(Convert.ToInt64(app.Hwnd));
            GetWindowThreadProcessId(excelHwnd, out excelPid);
        }
        catch { }

        string status = "failed";
        string runDetails = "";

        try
        {
            Console.WriteLine("==========================================================================");
            Console.WriteLine("【对照实验 1】ws 为活动工作表 (ActiveSheet == ws)，包含 ws.Cells(1, 1).Select");
            Console.WriteLine("==========================================================================");
            dynamic wb1 = app.Workbooks.Add();
            dynamic ws1 = wb1.Worksheets[1];
            ws1.Activate();
            Console.WriteLine(string.Format("工作簿名: {0}, 活动表: {1}, 宏目标表: {2}", wb1.Name, wb1.ActiveSheet.Name, ws1.Name));

            TestDirectRun(app, wb1, vbaWithSelect, "对照 1", excelPid);
            wb1.Close(false);

            Console.WriteLine("\n==========================================================================");
            Console.WriteLine("【对照实验 2】ws 非活动工作表 (ActiveSheet != ws)，包含 ws.Cells(1, 1).Select");
            Console.WriteLine("==========================================================================");
            dynamic wb2 = app.Workbooks.Add();
            dynamic wsOther = wb2.Worksheets.Add(Type.Missing, wb2.Worksheets[wb2.Worksheets.Count]);
            wsOther.Name = "Sheet2_Active";
            wsOther.Activate();
            Console.WriteLine(string.Format("工作簿名: {0}, 当前活动表: {1}, 宏内目标表 Worksheets(1): {2}", wb2.Name, wb2.ActiveSheet.Name, wb2.Worksheets[1].Name));
            Console.WriteLine(string.Format("活动表是否为宏内目标表? {0}", (string)wb2.ActiveSheet.Name == (string)wb2.Worksheets[1].Name));

            TestDirectRun(app, wb2, vbaWithSelect, "对照 2", excelPid);
            wb2.Close(false);

            Console.WriteLine("\n==========================================================================");
            Console.WriteLine("【对照实验 3】删除 ws.Cells(1, 1).Select，无论活动表是谁");
            Console.WriteLine("==========================================================================");
            dynamic wb3 = app.Workbooks.Add();
            dynamic wsOther3 = wb3.Worksheets.Add();
            wsOther3.Name = "OtherActiveSheet3";
            wsOther3.Activate();
            Console.WriteLine(string.Format("工作簿名: {0}, 当前活动表: {1}, 宏内目标表 Worksheets(1): {2}", wb3.Name, wb3.ActiveSheet.Name, wb3.Worksheets[1].Name));

            TestDirectRun(app, wb3, vbaWithoutSelect, "对照 3", excelPid);
            wb3.Close(false);

            status = "completed";
            runDetails = "Three condition experiments completed";
        }
        catch (Exception ex)
        {
            runDetails = "Failed with exception: " + ex.Message;
            throw;
        }
        finally
        {
            try { app.Quit(); } catch { }
            Marshal.ReleaseComObject(app);
            KillExcelProcesses();

            WriteRunMetadata(outputDir, toolName, status, runDetails);
            RemoveRunningLock(outputDir);
        }
    }

    static void TestDirectRun(dynamic app, dynamic wb, string code, string label, uint excelPid)
    {
        string modName = "TestMod_" + Guid.NewGuid().ToString("N").Substring(0, 6);
        dynamic vbComp = wb.VBProject.VBComponents.Add(1);
        vbComp.Name = modName;
        vbComp.CodeModule.AddFromString(code);

        // 静态编译测试
        dynamic compileBtn = null;
        try
        {
            app.VBE.ActiveVBProject = wb.VBProject;
            vbComp.Activate();
            compileBtn = app.VBE.CommandBars.FindControl(Type.Missing, 578);
        }
        catch { }

        bool btnBefore = compileBtn != null && (bool)compileBtn.Enabled;
        if (compileBtn != null && btnBefore)
        {
            try { compileBtn.Execute(); } catch { }
        }
        bool btnAfter = compileBtn != null && (bool)compileBtn.Enabled;
        Console.WriteLine(string.Format("[{0}] VBE 静态预编译 ID 578: 执行前 Enabled={1}, 执行后 Enabled={2} (静态预编译通过={3})",
            label, btnBefore, btnAfter, !btnAfter));

        // 启动看门狗自动处理弹窗
        using (var cts = new CancellationTokenSource())
        {
            var watchdog = Task.Run(() =>
            {
                while (!cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        EnumWindows((hWnd, lParam) =>
                        {
                            try
                            {
                                uint pid;
                                GetWindowThreadProcessId(hWnd, out pid);
                                if (excelPid != 0 && pid == excelPid)
                                {
                                    var sbClass = new StringBuilder(256);
                                    GetClassName(hWnd, sbClass, sbClass.Capacity);
                                    string cls = sbClass.ToString();
                                    if (cls == "#32770" || cls == "bosa_sdm_XL9")
                                    {
                                        var sbTitle = new StringBuilder(256);
                                        GetWindowText(hWnd, sbTitle, sbTitle.Capacity);
                                        Console.WriteLine(string.Format("[{0}] 捕获到 Excel 模态弹窗: [{1}] 类={2}, 正在自动响应以提取底层错误...", label, sbTitle, cls));
                                        SendMessage(hWnd, WM_COMMAND, (IntPtr)1, IntPtr.Zero);
                                        SendMessage(hWnd, WM_COMMAND, (IntPtr)2, IntPtr.Zero);
                                        PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                    }
                                }
                            }
                            catch { }
                            return true;
                        }, IntPtr.Zero);
                    }
                    catch { }
                    Thread.Sleep(80);
                }
            });

            // 在 Task 中带超时调用 app.Run
            Exception capturedEx = null;
            var runTask = Task.Run(() =>
            {
                try
                {
                    app.Run(modName + ".Main");
                }
                catch (Exception ex)
                {
                    capturedEx = ex;
                }
            });

            bool completed = runTask.Wait(4000);
            cts.Cancel();
            try { watchdog.Wait(300); } catch { }

            if (!completed)
            {
                Console.WriteLine(string.Format("[{0}] 直接 app.Run 执行超时 (可能受模态框挂起)", label));
            }
            else if (capturedEx != null)
            {
                var comEx = capturedEx as COMException;
                if (comEx != null)
                {
                    Console.WriteLine(string.Format("[{0}] 直接 app.Run 抛出 COMException! 错误码=0x{1:X8} (十进制={2}), 描述={3}",
                        label, (uint)comEx.ErrorCode, comEx.ErrorCode, comEx.Message.Trim()));
                }
                else
                {
                    Console.WriteLine(string.Format("[{0}] 直接 app.Run 抛出通用异常: {1}", label, capturedEx.Message.Trim()));
                }
            }
            else
            {
                Console.WriteLine(string.Format("[{0}] 直接 app.Run 执行成功，无任何异常！", label));
            }

            try { wb.VBProject.VBComponents.Remove(vbComp); } catch { }
        }
    }

    static void KillExcelProcesses()
    {
        try
        {
            var procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (var p in procs)
            {
                try { p.Kill(); } catch { }
            }
        }
        catch { }
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
}
