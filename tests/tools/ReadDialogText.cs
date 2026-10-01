using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public class ReadDialogText
{
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    private const uint WM_COMMAND = 0x0111;

    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string toolName = "ReadDialogText";
        string projectRoot = ResolveProjectRoot();

        string inputArg = args != null && args.Length > 0 ? args[0] : null;
        string outputArg = args != null && args.Length > 1 ? args[1] : null;

        string vbaPath = ResolveInputFile(inputArg, "docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba", projectRoot);
        if (string.IsNullOrEmpty(vbaPath) || !File.Exists(vbaPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input VBA file not found: '{0}'", vbaPath ?? "(null)"));
            Console.WriteLine("Usage: ReadDialogText.exe [inputVbaPath] [outputDir]");
            Console.WriteLine("Default fallback: docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        string outputDir = EnsureOutputDir(outputArg, toolName, projectRoot);
        string lockFile = Path.Combine(outputDir, ".running");
        File.WriteAllText(lockFile, DateTime.Now.ToString("o"), Encoding.UTF8);

        string vbaCode = File.ReadAllText(vbaPath, Encoding.UTF8);
        string status = "failed";
        string runDetails = "";

        dynamic app = null;
        dynamic wb = null;

        try
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            app = Activator.CreateInstance(excelType);
            app.Visible = false;

            uint excelPid = 0;
            try
            {
                IntPtr excelHwnd = new IntPtr(Convert.ToInt64(app.Hwnd));
                GetWindowThreadProcessId(excelHwnd, out excelPid);
            }
            catch { }

            wb = app.Workbooks.Add();
            dynamic wsOther = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            wsOther.Name = "ActiveSheet2";
            wsOther.Activate();

            dynamic vbComp = wb.VBProject.VBComponents.Add(1);
            vbComp.Name = "ModTest";
            vbComp.CodeModule.AddFromString(vbaCode);

            // 看门狗：捕获弹窗并读取里面所有文字，然后点“结束”按钮 (ID 4844 / IDCANCEL)
            string capturedDialogContent = "";
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
                                uint pid;
                                GetWindowThreadProcessId(hWnd, out pid);
                                if (excelPid != 0 && pid == excelPid)
                                {
                                    StringBuilder sbClass = new StringBuilder(256);
                                    GetClassName(hWnd, sbClass, sbClass.Capacity);
                                    if (sbClass.ToString() == "#32770")
                                    {
                                        StringBuilder sbTitle = new StringBuilder(256);
                                        GetWindowText(hWnd, sbTitle, sbTitle.Capacity);
                                        var fullContent = new StringBuilder();
                                        fullContent.AppendLine("Dialog Title: " + sbTitle.ToString());

                                        // 枚举子控件
                                        EnumChildWindows(hWnd, (childHwnd, l) =>
                                        {
                                            StringBuilder childClass = new StringBuilder(256);
                                            StringBuilder childText = new StringBuilder(1024);
                                            GetClassName(childHwnd, childClass, childClass.Capacity);
                                            GetWindowText(childHwnd, childText, childText.Capacity);
                                            string t = childText.ToString();
                                            if (!string.IsNullOrEmpty(t))
                                            {
                                                fullContent.AppendLine("  Child [" + childClass + "]: " + t);
                                            }
                                            return true;
                                        }, IntPtr.Zero);

                                        capturedDialogContent = fullContent.ToString();
                                        Console.WriteLine("=== 实时截获到的 VBA 运行时错误弹窗内容 ===");
                                        Console.WriteLine(capturedDialogContent);

                                        // 点击 IDCANCEL / 结束
                                        SendMessage(hWnd, WM_COMMAND, (IntPtr)2, IntPtr.Zero);
                                    }
                                }
                                return true;
                            }, IntPtr.Zero);
                        }
                        catch { }
                        Thread.Sleep(80);
                    }
                });

                try
                {
                    app.Run("ModTest.Main");
                    runDetails = "app.Run returned without exception";
                }
                catch (Exception ex)
                {
                    Console.WriteLine("app.Run 捕获异常: " + ex.Message);
                    runDetails = "Caught expected exception: " + ex.Message;
                }

                cts.Cancel();
                try { watchdog.Wait(300); } catch { }
            }

            if (!string.IsNullOrEmpty(capturedDialogContent))
            {
                File.WriteAllText(Path.Combine(outputDir, "dialog_content.txt"), capturedDialogContent, Encoding.UTF8);
            }

            status = "completed";
        }
        catch (Exception ex)
        {
            runDetails = "Execution failed: " + ex.Message;
            throw;
        }
        finally
        {
            if (wb != null) { try { wb.Close(false); } catch { } }
            if (app != null) { try { app.Quit(); } catch { } }
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
}
