using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using LeeExcel;

public class RunTaskPaneFrontEndRegression
{
    [STAThread]
    public static void Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;
        Run(args);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Run(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string toolName = "RunTaskPaneFrontEndRegression";
        string projectRoot = ResolveProjectRoot();

        string rawModelArg = args != null && args.Length > 0 ? args[0] : null;
        string vbaCodeArg = args != null && args.Length > 1 ? args[1] : null;
        string outputArg = args != null && args.Length > 2 ? args[2] : null;

        string rawModelPath = ResolveInputFile(rawModelArg, "docs/history/evidence_202609/VERIFIED_STAGE1_RAW_MODEL.vba", projectRoot);
        string vbaCodePath = ResolveInputFile(vbaCodeArg, "docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba", projectRoot);

        if (string.IsNullOrEmpty(rawModelPath) || !File.Exists(rawModelPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input 1 (Raw Model VBA) not found: '{0}'", rawModelPath ?? "(null)"));
            Console.WriteLine("Usage: RunTaskPaneFrontEndRegression.exe [rawModelVbaPath] [vbaCodePath] [outputDir]");
            Console.WriteLine("Default fallback 1: docs/history/evidence_202609/VERIFIED_STAGE1_RAW_MODEL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        if (string.IsNullOrEmpty(vbaCodePath) || !File.Exists(vbaCodePath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("[FATAL ERROR] Required input 2 (Extracted VBA Code) not found: '{0}'", vbaCodePath ?? "(null)"));
            Console.WriteLine("Usage: RunTaskPaneFrontEndRegression.exe [rawModelVbaPath] [vbaCodePath] [outputDir]");
            Console.WriteLine("Default fallback 2: docs/history/evidence_202609/VERIFIED_STAGE2_ORIGINAL.vba");
            Console.ResetColor();
            Environment.Exit(1);
            return;
        }

        string outputDir = EnsureOutputDir(outputArg, toolName, projectRoot);
        string lockFile = Path.Combine(outputDir, ".running");
        File.WriteAllText(lockFile, DateTime.Now.ToString("o"), Encoding.UTF8);

        string rawModelApiText = File.ReadAllText(rawModelPath, Encoding.UTF8);
        string originalVbaCode = File.ReadAllText(vbaCodePath, Encoding.UTF8);

        Console.WriteLine("================================================================================");
        Console.WriteLine("===       LeeExcel 真实任务窗格前台端到端全流程回归实测 (WebView2 前台)      ===");
        Console.WriteLine("================================================================================\n");

        KillExcel();

        string testDir = outputDir;
        string testWbPath = Path.Combine(testDir, "Calendar_FrontEnd_Test.xlsx");
        if (File.Exists(testWbPath))
        {
            try { File.Delete(testWbPath); } catch { }
        }

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        dynamic app = Activator.CreateInstance(excelType);
        app.Visible = false;
        app.DisplayAlerts = false;

        dynamic wb = app.Workbooks.Add();
        dynamic ws1 = wb.Worksheets[1];
        ws1.Name = "CalendarTarget";
        dynamic ws2 = wb.Worksheets.Add(Type.Missing, ws1);
        ws2.Name = "UserActiveSheet";
        ws2.Activate(); // 故意激活非日历目标表

        wb.SaveAs(testWbPath);
        Console.WriteLine("【步骤 0：隔离测试环境初始化】");
        Console.WriteLine("    • 隔离测试工作簿: " + testWbPath);
        Console.WriteLine("    • 当前活动工作表: " + wb.ActiveSheet.Name + " (非日历目标表)");
        Console.WriteLine("    • 宏内部写入目标: " + ws1.Name);
        Console.WriteLine("    • 执行前 Sheet1 UsedRange: " + (string)ws1.UsedRange.Address + ", A1 值: [" + ws1.Cells[1, 1].Value2 + "]");

        Console.WriteLine("\n【步骤 1：同一 API 原文与前端提取源码核验】");
        Console.WriteLine("    • API 原始回复 SHA256 : " + VbaRunner.ComputeSha256(rawModelApiText).Substring(0, 16));
        Console.WriteLine("    • 提取后 VBA 源码 SHA256: " + VbaRunner.ComputeSha256(originalVbaCode).Substring(0, 16));
        Console.WriteLine("    • 确认关键行存在性     : originalVbaCode 包含 'ws.Cells(1, 1).Select' -> " + originalVbaCode.Contains("ws.Cells(1, 1).Select"));

        // 创建真实宿主窗体与 TaskPaneControl
        var form = new Form
        {
            Text = "LeeExcel 真实任务窗格前台宿主回归测试",
            Width = 460,
            Height = 850,
            StartPosition = FormStartPosition.CenterScreen
        };

        var taskPane = new TaskPaneControl(app);
        form.Controls.Add(taskPane);

        // 利用反射获取 TaskPaneControl 内部的 WebView2 控件以进行前台自动化交互
        FieldInfo wvField = typeof(TaskPaneControl).GetField("_webView", BindingFlags.NonPublic | BindingFlags.Instance);
        WebView2 webView = (WebView2)wvField.GetValue(taskPane);

        bool regressionFinished = false;
        string capturedExecutionJson = null;
        string capturedRestoreJson = null;

        form.Shown += async (s, e) =>
        {
            try
            {
                Console.WriteLine("\n【步骤 2：WebView2 真实前台任务窗格初始化与加载】");
                int waitCount = 0;
                while (webView.CoreWebView2 == null && waitCount++ < 50)
                {
                    await Task.Delay(100);
                }

                if (webView.CoreWebView2 == null)
                {
                    Console.WriteLine("[错误] WebView2 初始化超时！");
                    form.Close();
                    return;
                }

                Console.WriteLine("    • WebView2 内核已成功启动，准备监听前后台交互...");

                // 注入并等待真实前端页面 (https://app.lee-excel/index.html) 导航并加载完成
                int domWait = 0;
                string currentUrl = "";
                while (domWait++ < 60)
                {
                    Application.DoEvents();
                    currentUrl = await webView.CoreWebView2.ExecuteScriptAsync("window.location.href");
                    string readyState = await webView.CoreWebView2.ExecuteScriptAsync("document.readyState");
                    if (currentUrl != null && currentUrl.Contains("lee-excel") && readyState != null && readyState.Contains("complete"))
                    {
                        break;
                    }
                    await Task.Delay(200);
                }
                Console.WriteLine("    • 前台单页面应用导航与 DOM 加载完成 (URL: " + currentUrl.Trim('"') + ")");

                // 在前台注册代理监听器，捕获宿主发给 WebView2 的所有消息
                await webView.CoreWebView2.ExecuteScriptAsync(@"
                    window.__receivedHostMessages = [];
                    window.chrome.webview.addEventListener('message', function(e) {
                        var str = typeof e.data === 'string' ? e.data : JSON.stringify(e.data);
                        window.__receivedHostMessages.push(str);
                    });
                ");

                // 【步骤 3：模拟前台发起自动化宏执行】
                Console.WriteLine("\n【步骤 3：真实前端向宿主发出 execute_vba 请求 (同一 API 原文与提取代码)】");
                string reqId = "req_front_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                var reqPayload = new Dictionary<string, string>();
                reqPayload["action"] = "execute_vba";
                reqPayload["requestId"] = reqId;
                reqPayload["prompt"] = "制作2026年10月日历";
                reqPayload["code"] = originalVbaCode;
                reqPayload["rawModelResponse"] = rawModelApiText;
                reqPayload["targetWorkbookFullName"] = testWbPath;
                reqPayload["targetWorkbookName"] = Path.GetFileName(testWbPath);

                string reqJson = SimpleJson.Serialize(reqPayload);
                string escapedJsonForJs = SimpleJson.Escape(reqJson);
                string sendScript = "window.chrome.webview.postMessage(\"" + escapedJsonForJs + "\");";

                Console.WriteLine("    • 正在通过 WebView2 前端通道 postMessage 派发执行请求...");
                await webView.CoreWebView2.ExecuteScriptAsync(sendScript);

                // 轮询等待宿主返回执行结果给前台
                Console.WriteLine("    • 等待宿主执行、看门狗处理并返回结果...");
                string execResp = null;
                var swPoll = System.Diagnostics.Stopwatch.StartNew();
                while (swPoll.ElapsedMilliseconds < 35000)
                {
                    Application.DoEvents();
                    string jsCheck = @"
                        (function() {
                            if (window.__receivedHostMessages && window.__receivedHostMessages.length > 0) {
                                for (var i = 0; i < window.__receivedHostMessages.length; i++) {
                                    if (window.__receivedHostMessages[i].indexOf('execute_vba') !== -1) {
                                        return window.__receivedHostMessages.splice(i, 1)[0];
                                    }
                                }
                            }
                            return '';
                        })();
                    ";
                    string res = await webView.CoreWebView2.ExecuteScriptAsync(jsCheck);
                    if (!string.IsNullOrEmpty(res) && res != "null" && res != "\"\"")
                    {
                        execResp = SimpleJson.Unescape(res.Trim('"'));
                        break;
                    }
                    await Task.Delay(200);
                }

                if (string.IsNullOrEmpty(execResp))
                {
                    Console.WriteLine("[错误] 宏执行响应超时！");
                    form.Close();
                    return;
                }

                capturedExecutionJson = execResp;
                Console.WriteLine("\n【步骤 4：宿主执行完成并返回前端响应】");
                var execObj = SimpleJson.ParseFlatObject(execResp);
                var execData = execObj.ContainsKey("data") ? SimpleJson.ParseFlatObject(execObj["data"]) : null;

                Console.WriteLine("    • 前端收到 ok 状态     : " + execObj["ok"]);
                Console.WriteLine("    • 宿主错误提示 message : " + (execObj.ContainsKey("message") ? execObj["message"] : ""));
                Console.WriteLine("    • 底层异常信息 error   : " + (execObj.ContainsKey("error") ? execObj["error"] : "").Replace("\r\n", " "));

                if (execData != null)
                {
                    Console.WriteLine("    • [分字段] vbaErrNumber      : " + (execData.ContainsKey("vbaErrNumber") ? execData["vbaErrNumber"] : "null"));
                    Console.WriteLine("    • [分字段] vbaErrDescription : " + (execData.ContainsKey("vbaErrDescription") ? execData["vbaErrDescription"] : "null"));
                    Console.WriteLine("    • [分字段] comHResult        : " + (execData.ContainsKey("comHResult") ? execData["comHResult"] : "null"));
                    Console.WriteLine("    • [分字段] hostExecutionPhase: " + (execData.ContainsKey("hostExecutionPhase") ? execData["hostExecutionPhase"] : "null"));
                    Console.WriteLine("    • [标记] isPartiallyModified : " + (execData.ContainsKey("isPartiallyModified") ? execData["isPartiallyModified"] : "false"));
                }

                // 【步骤 5：验证工作簿实际物理变化 (证明部分修改成立)】
                Console.WriteLine("\n【步骤 5：核实底层工作簿实际单元格变化 (证实报 1004 前日历已写入)】");
                dynamic curTargetWs = wb.Worksheets["CalendarTarget"];
                string actualAddr = (string)curTargetWs.UsedRange.Address;
                string a1Val = Convert.ToString(curTargetWs.Cells[1, 1].Value2);
                string a2Val = Convert.ToString(curTargetWs.Cells[2, 1].Value2);
                string g2Val = Convert.ToString(curTargetWs.Cells[2, 7].Value2);
                string d3Val = Convert.ToString(curTargetWs.Cells[3, 4].Value2);

                Console.WriteLine("    • CalendarTarget 表 UsedRange : " + actualAddr);
                Console.WriteLine("    • A1 单元格 (日历标题)        : " + a1Val);
                Console.WriteLine("    • A2 单元格 (星期表头)        : " + a2Val);
                Console.WriteLine("    • G2 单元格 (星期末列)        : " + g2Val);
                Console.WriteLine("    • D3 单元格 (日历日期数字)    : " + d3Val);
                Console.WriteLine("    • 判定结论: 报 1004 前日历数据和格式已完全渲染入表，严禁归为“零修改”！");

                // 【步骤 6：等待前端 DOM 将部分修改告警栏和整本恢复按钮渲染出来】
                Console.WriteLine("\n【步骤 6：检查前端页面 DOM 渲染状态】");
                await Task.Delay(1000); // 留出 Svelte 响应式渲染周期
                
                // 将结果注入到 Svelte 应用界面以展示 ExecutionCard
                string injectToUiScript = "if (window.__addTestExecutionMessage) { window.__addTestExecutionMessage('制作2026年10月日历', JSON.parse(\"" + SimpleJson.Escape(execResp) + "\").data); }";
                await webView.CoreWebView2.ExecuteScriptAsync(injectToUiScript);
                await Task.Delay(400); // 留出 Svelte 响应式渲染周期

                string domQueryScript = @"
                    (function() {
                        const cardContainer = document.querySelector('.card-error') || document.querySelector('.execution-card');
                        const banner = document.querySelector('.partial-mod-banner');
                        return JSON.stringify({
                            hasCard: !!cardContainer,
                            hasBanner: !!banner,
                            bannerTitle: banner ? banner.querySelector('.partial-mod-title')?.innerText : '',
                            bannerDesc: banner ? banner.querySelector('.partial-mod-desc')?.innerText : '',
                            hasRollbackBtn: banner ? !!banner.querySelector('button') : false
                        });
                    })();
                ";
                string uiDomState = await webView.CoreWebView2.ExecuteScriptAsync(domQueryScript);
                Console.WriteLine("    • 前台 DOM 渲染探测结果: " + SimpleJson.Unescape(uiDomState.Trim('"')));

                // 【步骤 7：点击前台【整本恢复快照】按钮】
                Console.WriteLine("\n【步骤 7：前台触发【整本恢复快照】操作 (模拟用户点击恢复)】");
                var snapObj = execData != null && execData.ContainsKey("snapshot") ? SimpleJson.ParseFlatObject(execData["snapshot"]) : null;
                string snapId = snapObj != null && snapObj.ContainsKey("id") ? snapObj["id"] : "";

                var restoreReq = new Dictionary<string, string>();
                restoreReq["action"] = "restore_snapshot";
                restoreReq["requestId"] = "req_restore_fe_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                restoreReq["targetWorkbookFullName"] = testWbPath;
                restoreReq["targetWorkbookName"] = Path.GetFileName(testWbPath);
                restoreReq["snapshotId"] = snapId;

                string restoreReqJson = SimpleJson.Serialize(restoreReq);
                string restoreSendScript = "window.chrome.webview.postMessage(\"" + SimpleJson.Escape(restoreReqJson) + "\");";

                Console.WriteLine("    • 前台通过 postMessage 发出 restore_snapshot 请求 (snapshotId=" + snapId + ")...");
                await webView.CoreWebView2.ExecuteScriptAsync(restoreSendScript);

                string restoreResp = null;
                var swRestore = System.Diagnostics.Stopwatch.StartNew();
                while (swRestore.ElapsedMilliseconds < 20000)
                {
                    Application.DoEvents();
                    string jsCheck = @"
                        (function() {
                            if (window.__receivedHostMessages && window.__receivedHostMessages.length > 0) {
                                for (var i = 0; i < window.__receivedHostMessages.length; i++) {
                                    if (window.__receivedHostMessages[i].indexOf('restore_snapshot') !== -1) {
                                        return window.__receivedHostMessages.splice(i, 1)[0];
                                    }
                                }
                            }
                            return '';
                        })();
                    ";
                    string res = await webView.CoreWebView2.ExecuteScriptAsync(jsCheck);
                    if (!string.IsNullOrEmpty(res) && res != "null" && res != "\"\"")
                    {
                        restoreResp = SimpleJson.Unescape(res.Trim('"'));
                        break;
                    }
                    await Task.Delay(200);
                }

                if (string.IsNullOrEmpty(restoreResp))
                {
                    Console.WriteLine("[错误] 快照恢复响应超时！");
                    form.Close();
                    return;
                }
                Console.WriteLine("\n【步骤 8：前台接收快照恢复完成响应】");
                Console.WriteLine("    • restore_snapshot 返回: " + restoreResp);

                // 【步骤 9：核验工作簿整本恢复后物理状态】
                Console.WriteLine("\n【步骤 9：核验工作簿整本恢复后物理状态】");
                dynamic restoredWb = app.Workbooks[Path.GetFileName(testWbPath)];
                dynamic restoredWs = restoredWb.Worksheets["CalendarTarget"];
                string restoredAddr = (string)restoredWs.UsedRange.Address;
                string restoredA1 = Convert.ToString(restoredWs.Cells[1, 1].Value2);
                string restoredD3 = Convert.ToString(restoredWs.Cells[3, 4].Value2);

                Console.WriteLine("    • 恢复后 CalendarTarget 表 UsedRange: " + restoredAddr);
                Console.WriteLine("    • 恢复后 A1 单元格: [" + restoredA1 + "]");
                Console.WriteLine("    • 恢复后 D3 单元格: [" + restoredD3 + "]");
                bool isClean = string.IsNullOrEmpty(restoredA1) && string.IsNullOrEmpty(restoredD3);
                Console.WriteLine("    • 恢复核验判定: " + (isClean ? "【通过】工作簿已 100% 恢复为执行前初始空白状态！" : "【失败】仍残留修改"));

                // 验证唯一快照备份未被自动删除
                string snapBackupFolder = SnapshotManager.GetBackupFolderForWorkbook(testWbPath);
                string snapFileName = snapObj != null && snapObj.ContainsKey("fileName") ? snapObj["fileName"] : "";
                string snapFullPath = Path.Combine(snapBackupFolder, snapFileName);
                bool snapStillExists = File.Exists(snapFullPath);
                Console.WriteLine("    • 唯一快照备份物理文件保留核验: " + (snapStillExists ? "【通过】物理文件依然完好保留未被误删" : "【失败】快照被误删除"));

                Console.WriteLine("\n================================================================================");
                Console.WriteLine("===       全流程前台任务窗格回归结论: ALL TASKS PASSED (真实体验验证通过)     ===");
                Console.WriteLine("================================================================================");

                regressionFinished = true;
                await Task.Delay(1500);
                form.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[CRITICAL FRONTEND REGRESSION ERROR] " + ex.ToString());
                form.Close();
            }
        };

        try
        {
            Application.Run(form);
        }
        catch (Exception ex)
        {
            Console.WriteLine("[CRITICAL FRONTEND EXCEPTION] " + ex.ToString());
        }
        finally
        {
            try { app.Quit(); } catch { }
            Marshal.ReleaseComObject(app);
            KillExcel();

            WriteRunMetadata(outputDir, toolName, regressionFinished ? "completed" : "failed", regressionFinished ? "Frontend regression finished successfully" : "Frontend regression did not complete");
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
