using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using LeeExcel;

public class RegressionTest
{
    public static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;
        Run(args);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Run(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string toolName = "RegressionTest";
        string projectRoot = ResolveProjectRoot();

        string inputArg = args != null && args.Length > 0 ? args[0] : null;
        string outputArg = args != null && args.Length > 1 ? args[1] : null;

        string vbaPath = ResolveInputFile(inputArg, "docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba", projectRoot);
        if (string.IsNullOrEmpty(vbaPath) || !File.Exists(vbaPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input VBA file not found: '{0}'", vbaPath ?? "(null)"));
            Console.WriteLine("Usage: RegressionTest.exe [inputVbaPath] [outputDir]");
            Console.WriteLine("Default fallback: docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        string outputDir = EnsureOutputDir(outputArg, toolName, projectRoot);
        string lockFile = Path.Combine(outputDir, ".running");
        File.WriteAllText(lockFile, DateTime.Now.ToString("o"), Encoding.UTF8);

        string calendarWithSelect = File.ReadAllText(vbaPath, Encoding.UTF8);

        Console.WriteLine("==========================================================================");
        Console.WriteLine("【回归验证】开始执行三项严密验收测试");
        Console.WriteLine("==========================================================================");

        KillExcelProcesses();

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        dynamic app = Activator.CreateInstance(excelType);
        app.Visible = false;
        app.DisplayAlerts = false;

        bool allPassed = true;
        string runDetails = "";

        try
        {
            // ---------------------------------------------------------
            // 验证 1：在指定单元格写值（AUTOMATION 正常通路）
            // ---------------------------------------------------------
            Console.WriteLine("\n[测试 1] AUTOMATION 通路：在指定单元格写值并读回");
            dynamic wb1 = app.Workbooks.Add();
            string normalVba = @"
Sub Main(targetWb As Workbook)
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets(1)
    ws.Range(""B2"").Value = ""回归测试项目""
    ws.Range(""C2"").Value = 9999
    ws.Range(""B2:C2"").Font.Bold = True
End Sub";

            var res1 = VbaRunner.RunVbaCode(app, wb1, normalVba, "指定单元格写入回归测试");
            string valB2 = Convert.ToString(wb1.Worksheets[1].Range["B2"].Value);
            string valC2 = Convert.ToString(wb1.Worksheets[1].Range["C2"].Value);
            bool test1Success = res1.success && (valB2 == "回归测试项目") && (valC2 == "9999");
            Console.WriteLine(string.Format("  - 结果: success={0}, phase={1}, B2='{2}', C2='{3}'",
                res1.success, res1.executionPhase, valB2, valC2));
            Console.WriteLine("  - [测试 1 状态]: " + (test1Success ? "PASS √" : "FAIL ✕"));
            if (!test1Success) allPassed = false;
            wb1.Close(false);

            // ---------------------------------------------------------
            // 验证 2：真实失败场景回归（包含 ws.Cells(1, 1).Select，活动表不一致）
            // 必须如实归类为 runtime_error (1004)，绝不能再报缺少 End If！
            // ---------------------------------------------------------
            Console.WriteLine("\n[测试 2] 错误如实分类：真实 1004 运行时错误场景（非活动工作表 Select）");
            dynamic wb2 = app.Workbooks.Add();
            dynamic wsOther2 = wb2.Worksheets.Add(Type.Missing, wb2.Worksheets[wb2.Worksheets.Count]);
            wsOther2.Name = "Sheet2_Active";
            wsOther2.Activate();

            var res2 = VbaRunner.RunVbaCode(app, wb2, calendarWithSelect, "真实 1004 日历宏回归测试");

            bool isRuntime1004 = !res2.success && (res2.executionPhase == "runtime_error") &&
                                 (res2.error.Contains("1004") || res2.error.Contains("Select") || (res2.rawErrorCode != null && res2.rawErrorCode.Contains("800A03EC")));
            bool noEndIfFalseClaim = !res2.error.Contains("缺少 End If") && !res2.summary.Contains("缺少 End If");

            Console.WriteLine(string.Format("  - 结果: success={0}, phase={1}, rawErrorCode='{2}', failureStage='{3}'",
                res2.success, res2.executionPhase, res2.rawErrorCode, res2.failureStage));
            Console.WriteLine(string.Format("  - 摘要 summary: {0}", res2.summary));
            Console.WriteLine(string.Format("  - 错误 error: {0}", res2.error.Trim()));
            Console.WriteLine(string.Format("  - 触发排查 triggerPoint: {0}", res2.errorTriggerPoint));
            Console.WriteLine(string.Format("  - 是否准确归类为 runtime_error 且包含 1004? {0}", isRuntime1004));
            Console.WriteLine(string.Format("  - 是否已彻底消除'缺少 End If'误报? {0}", noEndIfFalseClaim));

            bool test2Success = isRuntime1004 && noEndIfFalseClaim;
            Console.WriteLine("  - [测试 2 状态]: " + (test2Success ? "PASS √ (如实分类为运行时错误 1004，彻底根除缺少 End If 误报)" : "FAIL ✕"));
            if (!test2Success) allPassed = false;
            wb2.Close(false);

            // ---------------------------------------------------------
            // 验证 3：真正缺少 End If 的宏，如实归类为 compile_failed
            // ---------------------------------------------------------
            Console.WriteLine("\n[测试 3] 错误如实分类：真实语法缺少 End If 场景");
            dynamic wb3 = app.Workbooks.Add();
            string syntaxBadVba = @"
Sub Main(targetWb As Workbook)
    Dim x As Long
    x = 1
    If x = 1 Then
        x = 2
    ' 故意缺少 End If
End Sub";

            var res3 = VbaRunner.RunVbaCode(app, wb3, syntaxBadVba, "语法错误回归测试");
            bool isCompileFailed = !res3.success && (res3.executionPhase == "compile_failed" || res3.executionPhase == "syntax_failed");
            Console.WriteLine(string.Format("  - 结果: success={0}, phase={1}, failureStage='{2}', summary='{3}'",
                res3.success, res3.executionPhase, res3.failureStage, res3.summary));
            Console.WriteLine(string.Format("  - 错误 error: {0}", res3.error.Trim()));
            Console.WriteLine("  - [测试 3 状态]: " + (isCompileFailed ? "PASS √ (如实归类为编译/语法失败)" : "FAIL ✕"));
            if (!isCompileFailed) allPassed = false;
            wb3.Close(false);

            Console.WriteLine("\n==========================================================================");
            Console.WriteLine("【回归汇总】三项测试最终结论: " + (allPassed ? "全部通过 ALL PASSED √" : "存在未通过项 FAILED ✕"));
            Console.WriteLine("==========================================================================");

            runDetails = allPassed ? "All 3 tests passed" : "Some tests failed";
        }
        catch (Exception ex)
        {
            runDetails = "Regression failed with exception: " + ex.Message;
            throw;
        }
        finally
        {
            try { app.Quit(); } catch { }
            Marshal.ReleaseComObject(app);
            KillExcelProcesses();

            WriteRunMetadata(outputDir, toolName, allPassed ? "completed" : "failed", runDetails);
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
