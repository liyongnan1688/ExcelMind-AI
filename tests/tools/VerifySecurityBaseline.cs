using System;
using System.IO;
using System.Reflection;
using LeeExcel;

public class VerifySecurityBaseline
{
    private static int passCount = 0;
    private static int failCount = 0;

    public static int Main(string[] args)
    {
        try
        {
            return Run(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("[FATAL] " + ex.GetType().FullName + ": " + ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
            return 2;
        }
    }

    private static int Run(string[] args)
    {
        string projectRoot = args != null && args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", ".."));

        AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        CheckDna(projectRoot, "LeeExcel.dna");
        CheckDna(projectRoot, "LeeExcel64.dna");
        CheckSourceContains(projectRoot, "src\\TaskPaneControl.cs", "AreDevToolsEnabled = false");
        CheckSourceContains(projectRoot, "src\\TaskPaneControl.cs", "NavigationStarting += OnNavigationStarting");
        CheckSourceContains(projectRoot, "src\\TaskPaneControl.cs", "IsTrustedWebMessageSource(e.Source)");
        CheckSourceContains(projectRoot, "src\\TaskPaneControl.cs", "CoreWebView2HostResourceAccessKind.DenyCors");
        CheckSourceContains(projectRoot, "src\\VbaRunner.cs", "executionRisk.RiskLevel == VbaRiskLevel.L2 && !allowHighRisk");
        CheckSourceContains(projectRoot, "src\\NativeBridge.cs", "requireContextConfirm");
        CheckSourceContains(projectRoot, "src\\NativeBridge.cs", "requiresContextConfirm = requireContextConfirm");
        CheckSourceContains(projectRoot, "web\\src\\App.svelte", "requireUserConfirm");
        CheckSourceContains(projectRoot, "web\\src\\components\\ExecutionCard.svelte", "requireUserConfirm");
        CheckSourceContains(projectRoot, "web\\src\\services\\llm.ts", "escapeUntrustedSelectionText");
        CheckSourceContains(projectRoot, ".github\\workflows\\release.yml", "INPUT_VERSION: ${{ github.event.inputs.version }}");
        CheckSourceContains(projectRoot, ".github\\workflows\\release.yml", "RELEASE_TAG: ${{ needs.build.outputs.tag }}");
        CheckSourceNotContains(projectRoot, ".github\\workflows\\release.yml", "$inputVer = '${{");
        CheckSourceNotContains(projectRoot, ".github\\workflows\\release.yml", "$targetCommit = '${{");
        CheckNoExcelProcessKill(projectRoot, "tests\\tools\\RegressionTest.cs");
        CheckNoExcelProcessKill(projectRoot, "tests\\tools\\TestThreeConditions.cs");
        CheckNoExcelProcessKill(projectRoot, "tests\\tools\\RunTaskPaneFrontEndRegression.cs");
        CheckNoExcelProcessKill(projectRoot, "tests\\tools\\VerifyCompleteDelivery.cs");

        CheckRisk("Shell call", "Sub Main()\r\n    Shell \"cmd.exe /c whoami\", vbHide\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("Shell assignment", "Sub Main()\r\n    x = Shell(\"cmd.exe /c whoami\")\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("Call Shell", "Sub Main()\r\n    Call Shell(\"cmd.exe /c whoami\")\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("nested Shell", "Sub Main()\r\n    Debug.Print Shell(\"cmd.exe /c whoami\")\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("VBA.Shell", "Sub Main()\r\n    VBA.Shell \"cmd.exe /c whoami\"\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("nested CreateObject", "Sub Main()\r\n    Debug.Print CreateObject(\"WScript.Shell\").Run(\"calc\")\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("GetObject", "Sub Main()\r\n    Set x = GetObject(\"winmgmts:\")\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("Application.Run", "Sub Main()\r\n    Application.Run \"Danger\"\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("Application.Quit", "Sub Main()\r\n    Application.Quit\r\nEnd Sub", VbaRiskLevel.L2);
        CheckRisk("unknown expression", "Sub Main()\r\n    x = UnknownFunction()\r\nEnd Sub", VbaRiskLevel.L1);
        CheckRisk("read-only value", "Sub Main()\r\n    Dim x As Variant\r\n    x = Range(\"A1\").Value\r\nEnd Sub", VbaRiskLevel.L0);
        CheckRisk("read-only set", "Sub Main()\r\n    Dim ws As Worksheet\r\n    Set ws = ActiveSheet\r\nEnd Sub", VbaRiskLevel.L0);
        CheckRisk("string literal mentions Shell", "Sub Main()\r\n    MsgBox \"Shell is not executed\"\r\nEnd Sub", VbaRiskLevel.L0);
        CheckRisk("normal write", "Sub Main()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub", VbaRiskLevel.L1);

        var blocked = VbaRunner.RunVbaCode(null, null,
            "Sub Main()\r\n    x = Shell(\"cmd.exe /c whoami\")\r\nEnd Sub", "security baseline", null, null, false);
        Assert("unattended L2 execution is blocked", blocked != null && !blocked.success && blocked.failureStage == "risk_blocked");

        CheckExternalDataEscaping();

        Console.WriteLine("Security baseline: Pass = " + passCount + ", Fail = " + failCount);
        return failCount == 0 ? 0 : 1;
    }

    private static void CheckDna(string projectRoot, string fileName)
    {
        string path = Path.Combine(projectRoot, fileName);
        Assert(fileName + " exists", File.Exists(path));
        if (!File.Exists(path)) return;
        string text = File.ReadAllText(path);
        Assert(fileName + " uses ExplicitExports=true", text.IndexOf("ExplicitExports=\"true\"", StringComparison.OrdinalIgnoreCase) >= 0);
        Assert(fileName + " does not use ExplicitExports=false", text.IndexOf("ExplicitExports=\"false\"", StringComparison.OrdinalIgnoreCase) < 0);
    }

    private static void CheckSourceContains(string projectRoot, string relativePath, string expected)
    {
        string path = Path.Combine(projectRoot, relativePath);
        string text = File.Exists(path) ? File.ReadAllText(path) : "";
        Assert(relativePath + " contains " + expected, text.IndexOf(expected, StringComparison.Ordinal) >= 0);
    }

    private static void CheckSourceNotContains(string projectRoot, string relativePath, string unexpected)
    {
        string path = Path.Combine(projectRoot, relativePath);
        string text = File.Exists(path) ? File.ReadAllText(path) : "";
        Assert(relativePath + " does not contain " + unexpected, text.IndexOf(unexpected, StringComparison.Ordinal) < 0);
    }

    private static void CheckNoExcelProcessKill(string projectRoot, string relativePath)
    {
        string path = Path.Combine(projectRoot, relativePath);
        string text = File.Exists(path) ? File.ReadAllText(path) : "";
        Assert(relativePath + " does not kill Excel processes", text.IndexOf(".Kill()", StringComparison.Ordinal) < 0 && text.IndexOf("KillExcelProcesses", StringComparison.Ordinal) < 0 && text.IndexOf("KillExcel(", StringComparison.Ordinal) < 0);
    }
    private static void CheckExternalDataEscaping()
    {
        MethodInfo prepare = typeof(ExternalDataService).GetMethod(
            "PrepareCellValueForExcel",
            BindingFlags.NonPublic | BindingFlags.Static);
        if (prepare == null)
        {
            Assert("external data formula escaping method exists", false);
            return;
        }

        string[] samples = new string[]
        {
            "=cmd|'/c calc'!A0",
            "\t=cmd|'/c calc'!A0",
            "\r=cmd|'/c calc'!A0",
            "'alreadyQuoted",
            "42.5%",
            "000892",
            "1234567890123456789"
        };
        bool allEscaped = true;
        foreach (string sample in samples)
        {
            object actual = prepare.Invoke(null, new object[] { sample });
            if (!string.Equals(actual as string, "'" + sample, StringComparison.Ordinal))
            {
                allEscaped = false;
                break;
            }
        }
        Assert("external data cell/header escaping preserves formulas, control prefixes, quotes and IDs", allEscaped);
    }

    private static void CheckRisk(string label, string code, VbaRiskLevel expected)
    {
        var analysis = VbaRiskScanner.Analyze(code, "C:\\security-baseline\\book.xlsx");
        Assert(label + " => " + expected + " (actual " + analysis.RiskLevel + ")", analysis.RiskLevel == expected);
    }

    private static void Assert(string label, bool condition)
    {
        if (condition)
        {
            passCount++;
            Console.WriteLine("[PASS] " + label);
        }
        else
        {
            failCount++;
            Console.WriteLine("[FAIL] " + label);
        }
    }

    private static Assembly ResolveAssembly(object sender, ResolveEventArgs args)
    {
        try
        {
            string assemblyName = new AssemblyName(args.Name).Name + ".dll";
            string projectRoot = Environment.GetEnvironmentVariable("LEE_EXCEL_PROJECT_ROOT");
            if (string.IsNullOrEmpty(projectRoot)) return null;
            string binCandidate = Path.Combine(projectRoot, "bin", assemblyName);
            if (File.Exists(binCandidate)) return Assembly.LoadFrom(binCandidate);
        }
        catch { }
        return null;
    }
}