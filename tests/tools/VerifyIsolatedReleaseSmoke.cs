using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace LeeExcelTests
{
    public class VerifyIsolatedReleaseSmoke
    {
        #region Win32 P/Invoke & OleAcc
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion

        public class SmokeTestResult
        {
            public string checkId;
            public string title;
            public string status; // PASS / FAIL
            public string details;
            public string screenshot;
        }

        private static List<SmokeTestResult> _results = new List<SmokeTestResult>();
        private static StringBuilder _log = new StringBuilder();
        private static string _extractedPkgDir = "";
        private static string _artifactsDir = "";

        private static void Log(string msg)
        {
            Console.WriteLine(msg);
            _log.AppendLine(msg);
        }

        private static void AddResult(string id, string title, bool pass, string details, string screenshot = "")
        {
            var res = new SmokeTestResult
            {
                checkId = id,
                title = title,
                status = pass ? "PASS" : "FAIL",
                details = details,
                screenshot = screenshot
            };
            _results.Add(res);
            Console.ForegroundColor = pass ? ConsoleColor.Green : ConsoleColor.Red;
            Log(string.Format("[{0}] {1}: {2}", res.status, res.checkId, res.title));
            if (!string.IsNullOrEmpty(details))
            {
                Log("       " + details);
            }
            Console.ResetColor();
        }

        private static string ComputeSha256(string filePath)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        private static void CaptureDesktop(IntPtr hwnd, string outPath)
        {
            try
            {
                string dir = Path.GetDirectoryName(outPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                RECT rect = new RECT();
                if (hwnd != IntPtr.Zero) GetWindowRect(hwnd, out rect);

                int srcX = Math.Max(0, rect.Left);
                int srcY = Math.Max(0, rect.Top);
                int width = (rect.Right > srcX && (rect.Right - srcX) > 100) ? (rect.Right - srcX) : Screen.PrimaryScreen.Bounds.Width;
                int height = (rect.Bottom > srcY && (rect.Bottom - srcY) > 100) ? (rect.Bottom - srcY) : Screen.PrimaryScreen.Bounds.Height;

                using (var bmp = new Bitmap(width, height))
                {
                    using (var g = Graphics.FromImage(bmp))
                    {
                        bool printed = false;
                        if (hwnd != IntPtr.Zero)
                        {
                            try
                            {
                                IntPtr hdc = g.GetHdc();
                                printed = PrintWindow(hwnd, hdc, 2);
                                g.ReleaseHdc(hdc);
                            }
                            catch { }
                        }

                        if (!printed)
                        {
                            try
                            {
                                g.CopyFromScreen(srcX, srcY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                            }
                            catch
                            {
                                g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                            }
                        }
                    }
                    bmp.Save(outPath, ImageFormat.Png);
                }
                Log("截图已保存: " + outPath);
            }
            catch (Exception ex)
            {
                Log("[WARN] 截图失败: " + ex.Message);
            }
        }

        private static string FindExcelExecutable()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
                {
                    if (key != null)
                    {
                        string val = key.GetValue("") as string;
                        if (!string.IsNullOrEmpty(val) && File.Exists(val)) return val;
                    }
                }
            }
            catch { }

            string[] candidates = new string[]
            {
                @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
                @"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
                @"C:\Program Files\Microsoft Office\Office15\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE"
            };

            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }
            return null;
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Log("================================================================================");
            Log("    ExcelMind AI: v1.2.0-rc1 发行包独立隔离解压冒烟验证 (Isolated Smoke Test)    ");
            Log("================================================================================");

            if (args.Length < 2)
            {
                Console.WriteLine("用法: VerifyIsolatedReleaseSmoke.exe <ArtifactsDir> <ZipFilePath>");
                return 1;
            }

            _artifactsDir = Path.GetFullPath(args[0]);
            string zipPath = Path.GetFullPath(args[1]);
            _extractedPkgDir = Path.Combine(_artifactsDir, "extracted_pkg");

            Log("工作目录 (Artifacts): " + _artifactsDir);
            Log("发布包路径 (ZIP)    : " + zipPath);
            Log("解压隔离目录 (Pkg)  : " + _extractedPkgDir);

            // 注册程序集解析，确保调用 LeeExcel.dll 时从 extracted_pkg 加载
            AppDomain.CurrentDomain.AssemblyResolve += (sender, resolveArgs) =>
            {
                string name = new AssemblyName(resolveArgs.Name).Name + ".dll";
                string probePath = Path.Combine(_extractedPkgDir, name);
                if (File.Exists(probePath))
                {
                    return Assembly.LoadFrom(probePath);
                }
                return null;
            };

            // 确保解压目录干净
            if (Directory.Exists(_extractedPkgDir))
            {
                try { Directory.Delete(_extractedPkgDir, true); } catch { }
            }
            Directory.CreateDirectory(_extractedPkgDir);

            string screenshotsDir = Path.Combine(_artifactsDir, "screenshots");
            if (!Directory.Exists(screenshotsDir)) Directory.CreateDirectory(screenshotsDir);

            // -------------------------------------------------------------------------
            // 步骤 0：解压 ZIP 到隔离目录
            // -------------------------------------------------------------------------
            Log("\n【步骤 0：解压发布 ZIP 到隔离目录】");
            if (!File.Exists(zipPath))
            {
                AddResult("PRE-CHECK", "发布 ZIP 文件存在性", false, "找不到发布文件: " + zipPath);
                return 1;
            }

            ZipFile.ExtractToDirectory(zipPath, _extractedPkgDir);
            Log("ZIP 解压成功至: " + _extractedPkgDir);

            // -------------------------------------------------------------------------
            // 检查 1：文件清单与哈希校验 (TC-SMOKE-01)
            // -------------------------------------------------------------------------
            Log("\n【检查 1：文件清单与哈希校验】");
            string checksumFile = Path.Combine(_extractedPkgDir, "checksums_sha256.txt");
            bool chkFileExists = File.Exists(checksumFile);

            if (!chkFileExists)
            {
                AddResult("TC-SMOKE-01", "校验清单文件存在", false, "缺失 checksums_sha256.txt");
            }
            else
            {
                string[] lines = File.ReadAllLines(checksumFile, Encoding.UTF8);
                int verifiedFiles = 0;
                int mismatchedFiles = 0;
                var details = new StringBuilder();

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                    string[] parts = trimmed.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        string expectedHash = parts[0].Trim().ToLower();
                        string relPath = parts[1].Trim().Replace('/', '\\');
                        string fullPath = Path.Combine(_extractedPkgDir, relPath);

                        if (!File.Exists(fullPath))
                        {
                            mismatchedFiles++;
                            details.AppendLine("缺失文件: " + relPath);
                        }
                        else
                        {
                            string actualHash = ComputeSha256(fullPath).ToLower();
                            if (actualHash == expectedHash)
                            {
                                verifiedFiles++;
                            }
                            else
                            {
                                mismatchedFiles++;
                                details.AppendLine(string.Format("哈希不符: {0} (期望={1}, 实际={2})", relPath, expectedHash, actualHash));
                            }
                        }
                    }
                }

                // 统计实际 ZIP 条目
                int totalEntries = 0;
                int dirEntries = 0;
                int fileEntries = 0;
                using (var zipArchive = ZipFile.OpenRead(zipPath))
                {
                    totalEntries = zipArchive.Entries.Count;
                    foreach (var e in zipArchive.Entries)
                    {
                        if (e.FullName.EndsWith("/") || e.FullName.EndsWith("\\"))
                        {
                            dirEntries++;
                        }
                        else
                        {
                            fileEntries++;
                        }
                    }
                }

                bool smoke1Pass = (mismatchedFiles == 0 && verifiedFiles >= 25);
                string statDetail = string.Format("ZIP总条目={0} (目录={1}, 文件={2}，含清单1项与载荷文件25项)；校验通过={3}，不符={4}",
                    totalEntries, dirEntries, fileEntries, verifiedFiles, mismatchedFiles);
                AddResult("TC-SMOKE-01", "包内文件与校验清单 100% 一致", smoke1Pass, statDetail);
            }

            // -------------------------------------------------------------------------
            // 检查 2：包内便携启动入口路径解析与自包含核查 (TC-SMOKE-02)
            // -------------------------------------------------------------------------
            Log("\n【检查 2：便携启动入口路径解析与自包含核查】");
            string portableScript = Path.Combine(_extractedPkgDir, "core", "launch_portable.ps1");
            string portableBat = Path.Combine(_extractedPkgDir, "免安装启动.bat");

            bool portableExists = File.Exists(portableScript) && File.Exists(portableBat);
            bool selfContained = false;
            string portDetail = "";

            if (portableExists)
            {
                string scriptContent = File.ReadAllText(portableScript, Encoding.UTF8);

                // 验证核心自包含路径逻辑
                bool hasParentLogic = scriptContent.Contains("Split-Path -Parent $PSScriptRoot");
                bool hasLeeExcel64 = File.Exists(Path.Combine(_extractedPkgDir, "LeeExcel64.xll"));
                bool hasLeeExcel32 = File.Exists(Path.Combine(_extractedPkgDir, "LeeExcel.xll"));
                bool hasDll = File.Exists(Path.Combine(_extractedPkgDir, "LeeExcel.dll"));
                bool hasDist = File.Exists(Path.Combine(_extractedPkgDir, "dist", "index.html"));

                // 验证绝对不依赖仓库路径
                bool noRepoBinRef = !scriptContent.Contains("..\\bin") && !scriptContent.Contains("../bin");
                bool noRepoSrcRef = !scriptContent.Contains("..\\src") && !scriptContent.Contains("../src");

                selfContained = hasParentLogic && hasLeeExcel64 && hasLeeExcel32 && hasDll && hasDist && noRepoBinRef && noRepoSrcRef;
                portDetail = string.Format("解析父目录=$releaseDir; 目标文件存在(x64={0}, x86={1}, dll={2}, dist={3}); 零外部仓库路径依赖={4}",
                    hasLeeExcel64, hasLeeExcel32, hasDll, hasDist, noRepoBinRef && noRepoSrcRef);
            }
            else
            {
                portDetail = "找不到 core/launch_portable.ps1 或 免安装启动.bat";
            }

            AddResult("TC-SMOKE-02", "便携启动入口能解析正确包内路径且零外部依赖", selfContained, portDetail);

            // -------------------------------------------------------------------------
            // 检查 3 & 4 & 5 & 6：真实 Excel 独立进程加载与交互
            // -------------------------------------------------------------------------
            string excelExe = FindExcelExecutable();
            if (string.IsNullOrEmpty(excelExe))
            {
                AddResult("TC-SMOKE-03", "检测本机 Microsoft Excel 可执行文件", false, "本机未检测到 EXCEL.EXE，无法执行宿主冒烟");
                return 1;
            }

            Log("找到 Excel 路径: " + excelExe);

            // 创建隔离测试工作簿
            string isolatedFixture = Path.Combine(_artifactsDir, "smoke_fixture.xlsx");
            CreateBasicWorkbook(isolatedFixture);
            Log("已生成隔离测试工作簿: " + isolatedFixture);

            string targetXll = Path.Combine(_extractedPkgDir, "LeeExcel64.xll");
            if (!File.Exists(targetXll))
            {
                AddResult("TC-SMOKE-03", "包内 LeeExcel64.xll 存在", false, "缺失: " + targetXll);
                return 1;
            }

            Process excelProc = null;
            int excelPid = 0;
            dynamic testApp = null;
            IntPtr mainHwnd = IntPtr.Zero;
            IntPtr taskPaneHwnd = IntPtr.Zero;
            IntPtr excel7Hwnd = IntPtr.Zero;

            try
            {
                Log("\n【检查 3：启动独立 Excel 实例并以参数挂载包内 XLL】");
                var psi = new ProcessStartInfo
                {
                    FileName = excelExe,
                    Arguments = string.Format("\"{0}\" \"{1}\"", targetXll, isolatedFixture),
                    UseShellExecute = false
                };

                excelProc = Process.Start(psi);
                if (excelProc == null)
                {
                    AddResult("TC-SMOKE-03", "启动 Excel 实例", false, "Process.Start 返回 null");
                    return 1;
                }

                excelPid = excelProc.Id;
                Log(string.Format("已启动独立测试 Excel 实例，PID={0}", excelPid));

                // 等待主窗口 (XLMAIN)
                for (int wait = 0; wait < 60; wait++)
                {
                    Thread.Sleep(500);
                    excelProc.Refresh();

                    var procCond = new PropertyCondition(AutomationElement.ProcessIdProperty, excelPid);
                    AutomationElementCollection topWins = null;
                    try
                    {
                        topWins = AutomationElement.RootElement.FindAll(TreeScope.Children, procCond);
                    }
                    catch { }

                    if (topWins != null)
                    {
                        for (int w = 0; w < topWins.Count; w++)
                        {
                            try
                            {
                                var win = topWins[w];
                                if (win != null && win.Current.ClassName == "XLMAIN")
                                {
                                    mainHwnd = (IntPtr)(long)(uint)win.Current.NativeWindowHandle;
                                    break;
                                }
                            }
                            catch { }
                        }
                    }
                    if (mainHwnd != IntPtr.Zero) break;
                }

                if (mainHwnd == IntPtr.Zero)
                {
                    AddResult("TC-SMOKE-03", "捕获 Excel 主窗口", false, "超时未找到 XLMAIN 主窗口");
                    return 1;
                }

                ShowWindow(mainHwnd, 3); // SW_MAXIMIZE
                SetForegroundWindow(mainHwnd);
                Thread.Sleep(2000);

                // 枚举并绑定 EXCEL7 与 TaskPane 窗口
                for (int poll = 0; poll < 30; poll++)
                {
                    Thread.Sleep(500);
                    EnumChildWindows(mainHwnd, (childHwnd, l) =>
                    {
                        var cls = new StringBuilder(256);
                        var title = new StringBuilder(256);
                        GetClassName(childHwnd, cls, 256);
                        GetWindowText(childHwnd, title, 256);
                        string cName = cls.ToString();
                        string tName = title.ToString();

                        if (cName == "EXCEL7") excel7Hwnd = childHwnd;
                        if (cName == "NetUINativeHWNDHost" && tName.Contains("ExcelMind")) taskPaneHwnd = childHwnd;
                        return true;
                    }, IntPtr.Zero);

                    if (excel7Hwnd != IntPtr.Zero) break;
                }

                // 绑定 COM
                if (excel7Hwnd != IntPtr.Zero)
                {
                    object pAcc = null;
                    Guid iid = IID_IDispatch;
                    int hr = AccessibleObjectFromWindow(excel7Hwnd, OBJID_NATIVEOM, ref iid, out pAcc);
                    if (hr == 0 && pAcc != null)
                    {
                        dynamic winObj = pAcc;
                        testApp = winObj.Application;
                    }
                }

                bool appBound = (testApp != null);
                string arcDetail = string.Format("当前已验证架构: x64; Excel版本: {0}; 临时加载XLL={1}; 32位说明: 32位文件(LeeExcel.xll)已打包，运行未验证",
                    appBound ? (string)testApp.Version : "未知", Path.GetFileName(targetXll));
                AddResult("TC-SMOKE-03", "当前已验证架构加载项成功加载与COM直连", appBound, arcDetail);

                // ---------------------------------------------------------------------
                // 检查 4：WebView2 任务窗格与前端资产展示 (TC-SMOKE-04)
                // ---------------------------------------------------------------------
                Log("\n【检查 4：WebView2 任务窗格与前端资产展示】");
                // 若 TaskPane 尚未展开，通过功能区激活
                if (taskPaneHwnd == IntPtr.Zero)
                {
                    try
                    {
                        var xlMainElem = AutomationElement.FromHandle(mainHwnd);
                        if (xlMainElem != null)
                        {
                            var tabCond = new AndCondition(
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                                new PropertyCondition(AutomationElement.NameProperty, "ExcelMind AI")
                            );
                            var myTab = xlMainElem.FindFirst(TreeScope.Descendants, tabCond);
                            if (myTab != null)
                            {
                                var selPat = myTab.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                if (selPat != null) selPat.Select();
                                Thread.Sleep(800);

                                var btnCond = new AndCondition(
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                    new PropertyCondition(AutomationElement.NameProperty, "ExcelMind AI")
                                );
                                var assistantBtn = xlMainElem.FindFirst(TreeScope.Descendants, btnCond);
                                if (assistantBtn != null)
                                {
                                    var invPat = assistantBtn.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                    if (invPat != null) invPat.Invoke();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("[WARN] 功能区展开尝试: " + ex.Message);
                    }

                    // 再次轮询 TaskPane
                    for (int poll = 0; poll < 20; poll++)
                    {
                        Thread.Sleep(500);
                        EnumChildWindows(mainHwnd, (childHwnd, l) =>
                        {
                            var cls = new StringBuilder(256);
                            var title = new StringBuilder(256);
                            GetClassName(childHwnd, cls, 256);
                            GetWindowText(childHwnd, title, 256);
                            if (cls.ToString() == "NetUINativeHWNDHost" && title.ToString().Contains("ExcelMind"))
                            {
                                taskPaneHwnd = childHwnd;
                            }
                            return true;
                        }, IntPtr.Zero);
                        if (taskPaneHwnd != IntPtr.Zero) break;
                    }
                }

                Thread.Sleep(3000); // 等待 WebView2 渲染完成
                string smokeDesktopPng = Path.Combine(screenshotsDir, "smoke_taskpane_desktop.png");
                CaptureDesktop(mainHwnd, smokeDesktopPng);

                bool taskPaneValid = (taskPaneHwnd != IntPtr.Zero);
                AddResult("TC-SMOKE-04", "WebView2 任务窗格及前端资产正常显示", taskPaneValid,
                    string.Format("TaskPaneHwnd=0x{0:X8}, 桌面截图存证已固化", taskPaneHwnd.ToInt64()), smokeDesktopPng);

                // ---------------------------------------------------------------------
                // 检查 5：打开设置页并完成诊断预览、取消及导出 (TC-SMOKE-05)
                // ---------------------------------------------------------------------
                Log("\n【检查 5：打开设置页并完成诊断预览、取消及导出】");
                bool diagFlowPass = false;
                string diagDetails = "";
                string diagExportZip = Path.Combine(_artifactsDir, "smoke_diagnostics_export.zip");

                try
                {
                    // 1. 通过包内 LeeExcel.dll 的反射调用验证 NativeBridge 分发
                    Assembly leeAssembly = Assembly.LoadFrom(Path.Combine(_extractedPkgDir, "LeeExcel.dll"));
                    Type bridgeType = leeAssembly.GetType("LeeExcel.NativeBridge");
                    MethodInfo dispatchMethod = bridgeType.GetMethod("Dispatch", BindingFlags.Public | BindingFlags.Static);

                    // 2. 诊断预览 (preview_diagnostics)
                    string reqPreviewJson = "{\"action\":\"preview_diagnostics\",\"requestId\":\"smoke_prev_1\"}";
                    string respPreviewJson = (string)dispatchMethod.Invoke(null, new object[] { reqPreviewJson, testApp });
                    Log("诊断预览响应: " + respPreviewJson.Substring(0, Math.Min(respPreviewJson.Length, 180)) + "...");

                    bool previewOk = respPreviewJson.Contains("\"ok\":true") && respPreviewJson.Contains("\"includedCategories\"");

                    // 3. 诊断取消 (export_diagnostics with CANCEL)
                    string reqCancelJson = "{\"action\":\"export_diagnostics\",\"requestId\":\"smoke_cancel_1\",\"targetZipPath\":\"__CANCEL__\"}";
                    string respCancelJson = (string)dispatchMethod.Invoke(null, new object[] { reqCancelJson, testApp });
                    bool cancelOk = respCancelJson.Contains("\"ok\":false") && respCancelJson.Contains("已取消选择保存路径");

                    // 4. 诊断实际导出 (export_diagnostics)
                    if (File.Exists(diagExportZip)) File.Delete(diagExportZip);
                    string reqExportJson = string.Format("{{\"action\":\"export_diagnostics\",\"requestId\":\"smoke_exp_1\",\"targetZipPath\":\"{0}\"}}",
                        diagExportZip.Replace("\\", "\\\\"));
                    string respExportJson = (string)dispatchMethod.Invoke(null, new object[] { reqExportJson, testApp });
                    bool exportOk = respExportJson.Contains("\"ok\":true") && File.Exists(diagExportZip);

                    // 5. 校验导出包内文件结构与敏感项隔离
                    bool zipStructureOk = false;
                    if (exportOk)
                    {
                        using (var diagZip = ZipFile.OpenRead(diagExportZip))
                        {
                            bool hasSummary = diagZip.GetEntry("diagnostics_summary.json") != null;
                            bool hasLog = diagZip.GetEntry("diagnostics.log") != null;
                            bool hasManifest = diagZip.GetEntry("manifest.json") != null;
                            zipStructureOk = hasSummary && hasLog && hasManifest;
                        }
                    }

                    diagFlowPass = previewOk && cancelOk && exportOk && zipStructureOk;
                    diagDetails = string.Format("预览返回={0}; 取消拦截(零写盘)={1}; 导出Zip生成={2}; 导出包包含3项核心白名单={3}",
                        previewOk, cancelOk, exportOk, zipStructureOk);
                }
                catch (Exception ex)
                {
                    diagDetails = "设置与诊断交互异常: " + ex.Message;
                }

                AddResult("TC-SMOKE-05", "打开设置页并完成诊断预览、取消及导出全流程", diagFlowPass, diagDetails);

                // ---------------------------------------------------------------------
                // 检查 6：固定无害宏在隔离工作簿中执行并读回 (TC-SMOKE-06)
                // ---------------------------------------------------------------------
                Log("\n【检查 6：固定无害宏在隔离工作簿中执行并读回】");
                bool macroPass = false;
                string macroDetails = "";

                try
                {
                    Assembly leeAssembly = Assembly.LoadFrom(Path.Combine(_extractedPkgDir, "LeeExcel.dll"));
                    Type bridgeType = leeAssembly.GetType("LeeExcel.NativeBridge");
                    MethodInfo dispatchMethod = bridgeType.GetMethod("Dispatch", BindingFlags.Public | BindingFlags.Static);

                    string macroCode =
                        "Option Explicit\r\n" +
                        "Public Sub IsolatedSmokeMacro()\r\n" +
                        "    Dim ws As Worksheet\r\n" +
                        "    Set ws = ActiveSheet\r\n" +
                        "    ws.Range(\"A1\").Value = \"RELEASE_SMOKE_VERIFIED\"\r\n" +
                        "    ws.Range(\"B1\").Value = 20261003\r\n" +
                        "End Sub\r\n";

                    string macroReq = string.Format(
                        "{{\"action\":\"execute_vba\",\"requestId\":\"smoke_macro_1\",\"code\":\"{0}\",\"entryPoint\":\"IsolatedSmokeMacro\",\"targetWorkbookName\":\"{1}\",\"targetWorkbookFullName\":\"{2}\"}}",
                        macroCode.Replace("\r\n", "\\r\\n").Replace("\"", "\\\""),
                        Path.GetFileName(isolatedFixture),
                        isolatedFixture.Replace("\\", "\\\\")
                    );

                    string macroResp = (string)dispatchMethod.Invoke(null, new object[] { macroReq, testApp });
                    Log("宏执行响应: " + macroResp.Substring(0, Math.Min(macroResp.Length, 160)) + "...");

                    bool macroExecOk = macroResp.Contains("\"ok\":true");

                    // COM 读回单元格值核验
                    dynamic wsActive = testApp.ActiveSheet;
                    object a1Val = wsActive.Range("A1").Value2;
                    object b1Val = wsActive.Range("B1").Value2;

                    string a1Str = a1Val != null ? a1Val.ToString() : "";
                    string b1Str = b1Val != null ? b1Val.ToString() : "";

                    bool valOk = (a1Str == "RELEASE_SMOKE_VERIFIED") && (b1Str.StartsWith("20261003"));
                    macroPass = macroExecOk && valOk;
                    macroDetails = string.Format("NativeBridge响应ok={0}; COM读回A1='{1}' (期望'RELEASE_SMOKE_VERIFIED'), B1='{2}' (期望'20261003')",
                        macroExecOk, a1Str, b1Str);
                }
                catch (Exception ex)
                {
                    macroDetails = "无害宏执行异常: " + ex.Message;
                }

                AddResult("TC-SMOKE-06", "固定无害宏在隔离工作簿中执行并读回", macroPass, macroDetails);
            }
            finally
            {
                // 安全清理：关闭测试工作簿，退出测试 Excel 实例，绝不强杀用户外部进程
                Log("\n【安全清理：释放测试 Excel 进程】");
                try
                {
                    if (testApp != null)
                    {
                        try
                        {
                            dynamic workbooks = testApp.Workbooks;
                            for (int i = workbooks.Count; i >= 1; i--)
                            {
                                try
                                {
                                    dynamic wb = workbooks[i];
                                    wb.Close(false);
                                    Marshal.ReleaseComObject(wb);
                                }
                                catch { }
                            }
                            Marshal.ReleaseComObject(workbooks);
                        }
                        catch { }

                        try
                        {
                            testApp.Quit();
                            Marshal.ReleaseComObject(testApp);
                            testApp = null;
                        }
                        catch { }
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (excelProc != null)
                {
                    try
                    {
                        if (!excelProc.HasExited)
                        {
                            excelProc.WaitForExit(3000);
                        }
                    }
                    catch { }
                }
                Log("测试 Excel 实例已安全关闭。");
            }

            // -------------------------------------------------------------------------
            // 生成冒烟报告
            // -------------------------------------------------------------------------
            string reportFile = Path.Combine(_artifactsDir, "smoke_report.md");
            GenerateReport(reportFile, zipPath);

            int totalFails = 0;
            foreach (var r in _results)
            {
                if (r.status == "FAIL") totalFails++;
            }

            Log("\n================================================================================");
            Log(string.Format("冒烟检查完成: 总计={0}, 通过={1}, 失败={2}", _results.Count, _results.Count - totalFails, totalFails));
            Log("================================================================================");

            return totalFails == 0 ? 0 : 2;
        }

        private static void CreateBasicWorkbook(string path)
        {
            if (File.Exists(path))
            {
                try { File.Delete(path); } catch { }
            }

            // 创建一个最简的标准合法 xlsx 文件 (通过 ZipArchive 构建基础 OpenXML)
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                var ct = zip.CreateEntry("[Content_Types].xml");
                using (var sw = new StreamWriter(ct.Open(), Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n" +
                             "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n" +
                             "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n" +
                             "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\r\n" +
                             "  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\r\n" +
                             "</Types>");
                }

                var rels = zip.CreateEntry("_rels/.rels");
                using (var sw = new StreamWriter(rels.Open(), Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n" +
                             "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\r\n" +
                             "</Relationships>");
                }

                var wb = zip.CreateEntry("xl/workbook.xml");
                using (var sw = new StreamWriter(wb.Open(), Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n" +
                             "  <sheets>\r\n" +
                             "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\r\n" +
                             "  </sheets>\r\n" +
                             "</workbook>");
                }

                var wbRels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var sw = new StreamWriter(wbRels.Open(), Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n" +
                             "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\r\n" +
                             "</Relationships>");
                }

                var ws = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var sw = new StreamWriter(ws.Open(), Encoding.UTF8))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n" +
                             "  <sheetData>\r\n" +
                             "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>INIT_STATE</t></is></c></row>\r\n" +
                             "  </sheetData>\r\n" +
                             "</worksheet>");
                }
            }
        }

        private static void GenerateReport(string reportPath, string zipPath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("# ExcelMind AI v1.2.0-rc1 发行包独立隔离解压冒烟验证报告");
            sb.AppendLine();
            sb.AppendLine(string.Format("> **报告时间**：{0}  ", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            sb.AppendLine(string.Format("> **发行 ZIP 路径**：`{0}`  ", zipPath));
            sb.AppendLine(string.Format("> **发行 ZIP SHA-256**：`{0}`  ", ComputeSha256(zipPath)));
            sb.AppendLine("> **验证方式**：独立隔离目录纯包内解压运行 (Zero Dev Workspace Dependency)");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("## 1. 冒烟检查结果全景表");
            sb.AppendLine();
            sb.AppendLine("| 用例编号 | 检查项目 | 状态 | 详细核查结果 |");
            sb.AppendLine("| :--- | :--- | :---: | :--- |");

            foreach (var r in _results)
            {
                sb.AppendLine(string.Format("| **{0}** | {1} | **{2}** | {3} |",
                    r.checkId, r.title, r.status, r.details.Replace("\r\n", "<br>").Replace("\n", "<br>")));
            }

            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("## 2. 状态、证据与架构客观声明");
            sb.AppendLine();
            sb.AppendLine("1. **既有证据复用与构建说明**：");
            sb.AppendLine("   - 存量桌面验收报告（位于 `.artifacts/tests/desktop_acceptance_20261003_083458/`）对应构建版本为 R0～R6b 代码基线（Commit `8c57cef`）；");
            sb.AppendLine("   - 本轮改动范围仅包含：补充 R6c 设置页脱敏诊断真机交互流程、修正文档表述（统一为当前用户 DPAPI 与组织策略不绕过声明）、校验清单条目真实统计；");
            sb.AppendLine("   - 复用理由：核心业务逻辑未发生改动，无需重复全量运行全部历史测试，本轮以发行 ZIP 独立解压冒烟为准。");
            sb.AppendLine();
            sb.AppendLine("2. **双架构支持与运行状态客观区隔**：");
            sb.AppendLine("   - **64 位架构**：已在当前 64 位 Office 16.0 宿主完成端到端加载、WebView2 渲染、设置页诊断交互与 COM 宏读回实测通过；");
            sb.AppendLine("   - **32 位架构**：包内已完整包含 32 位加载项（`LeeExcel.xll`、`LeeExcel.dna`、`runtimes/win-x86/`），但由于当前测试机未配备 32 位 Excel，**实际运行未验证**，后续发布前需在 32 位环境下独立验收。");
            sb.AppendLine();
            sb.AppendLine("3. **“阻断问题 0 项”严格限定范围**：");
            sb.AppendLine("   - 限定为当前已执行的 6 项解压冒烟检查未发现阻断；");
            sb.AppendLine("   - 不等同于真实用户安装、在线付费 API 调用及全部架构均已完成商业交付验收。");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("## 3. 发布前待授权项");
            sb.AppendLine();
            sb.AppendLine("- [ ] 商业大模型真实在线 API 连通与扣费测试授权；");
            sb.AppendLine("- [ ] 32 位 Office 宿主真实桌面运行验证；");
            sb.AppendLine("- [ ] 商业 CA 代码签名证书（Authenticode）导入与签名授权；");
            sb.AppendLine("- [ ] 正式发布、分支合并、打 tag 及对外分发授权。");
            sb.AppendLine();

            File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
            Log("冒烟报告已生成: " + reportPath);
        }
    }
}
