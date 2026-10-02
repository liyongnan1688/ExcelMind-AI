using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

namespace LeeExcelTests
{
    public class ProbeDesktopUia
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

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

        private static StringBuilder _sbLog = new StringBuilder();

        private static void Log(string msg)
        {
            Console.WriteLine(msg);
            _sbLog.AppendLine(msg);
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("        ExcelMind AI: 真实桌面端 Excel + UIA + WebView2 能力核查探测器         ");
            Console.WriteLine("================================================================================");

            string artifactsDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
            string logPath = Path.Combine(artifactsDir, "probe_capability_report.txt");
            string screenshotPath = Path.Combine(artifactsDir, "probe_excel_full.png");

            Log("探测启动时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 1. 记录现有已存在的 EXCEL 进程 PID，确保绝不误杀用户实例
            var existingPids = new HashSet<int>();
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                existingPids.Add(p.Id);
                Log(string.Format("[保护] 发现系统现有 Excel 进程 PID={0}, 标题='{1}' (绝不接管/关闭)", p.Id, p.MainWindowTitle));
            }

            // 2. 寻找 Excel.exe 路径与插件 xll
            string excelPath = null;
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
                {
                    if (key != null) excelPath = key.GetValue(null) as string;
                }
            }
            catch { }

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                excelPath = @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE";
            }

            string projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));
            string xllPath = Path.Combine(projectRoot, "bin", "LeeExcel64.xll");

            Log("Excel 路径: " + excelPath);
            Log("插件路径: " + xllPath);

            if (!File.Exists(excelPath))
            {
                Log("[FATAL] Excel.exe 不存在: " + excelPath);
                File.WriteAllText(logPath, _sbLog.ToString(), Encoding.UTF8);
                return 1;
            }
            if (!File.Exists(xllPath))
            {
                Log("[FATAL] 插件 xll 不存在: " + xllPath);
                File.WriteAllText(logPath, _sbLog.ToString(), Encoding.UTF8);
                return 1;
            }

            string testWbFile = Path.Combine(artifactsDir, "probe_test.xlsx");
            if (!File.Exists(testWbFile))
            {
                File.Copy(Path.Combine(projectRoot, ".artifacts", ".artifacts", "tests", "macro_verification", "IsolatedRealMacroTest.xlsx"), testWbFile, true);
            }

            Process testProcess = null;
            int testPid = 0;

            try
            {
                // 3. 启动隔离测试 Excel 实例（同时传入 xll 加载项与测试工作簿）
                var psi = new ProcessStartInfo
                {
                    FileName = excelPath,
                    Arguments = "\"" + xllPath + "\" \"" + testWbFile + "\"",
                    UseShellExecute = false
                };

                Log("正在启动独立测试 Excel 实例: " + psi.Arguments);
                testProcess = Process.Start(psi);
                if (testProcess == null)
                {
                    Log("[FATAL] 启动 Excel 进程失败");
                    return 1;
                }

                testPid = testProcess.Id;
                Log("测试 Excel 已启动，PID: " + testPid);

                // 4. 等待 Excel 主窗口出现并定位 XLMAIN
                IntPtr mainHwnd = IntPtr.Zero;
                AutomationElement xlMainElement = null;

                for (int i = 0; i < 60; i++)
                {
                    Thread.Sleep(500);
                    testProcess.Refresh();

                    var procCond = new PropertyCondition(AutomationElement.ProcessIdProperty, testPid);
                    var topWins = AutomationElement.RootElement.FindAll(TreeScope.Children, procCond);

                    for (int w = 0; w < topWins.Count; w++)
                    {
                        var win = topWins[w];
                        string cls = win.Current.ClassName;
                        Log(string.Format("  [检查窗口] HWND=0x{0:X8}, Class='{1}', Name='{2}'", win.Current.NativeWindowHandle, cls, win.Current.Name));
                        if (cls == "XLMAIN")
                        {
                            xlMainElement = win;
                            mainHwnd = new IntPtr(win.Current.NativeWindowHandle);
                            break;
                        }
                    }

                    if (mainHwnd != IntPtr.Zero) break;
                }

                if (mainHwnd == IntPtr.Zero)
                {
                    Log("[FATAL] 等待 Excel 主窗口 (XLMAIN) 超时！");
                    return 1;
                }

                Log(string.Format("Excel 主窗口 (XLMAIN) 句柄已捕获: 0x{0:X8}, Title='{1}'", mainHwnd.ToInt64(), xlMainElement.Current.Name));
                ShowWindow(mainHwnd, 3); // SW_MAXIMIZE
                SetForegroundWindow(mainHwnd);

                // 5. 循环等待并探测窗口子层级（等待 WebView2 与 EXCEL7 完全渲染）
                IntPtr hwndExcel7 = IntPtr.Zero;
                IntPtr hwndWebView = IntPtr.Zero;
                var foundChildWindows = new List<string>();

                for (int poll = 0; poll < 20; poll++)
                {
                    Thread.Sleep(500);
                    foundChildWindows.Clear();
                    hwndExcel7 = IntPtr.Zero;
                    hwndWebView = IntPtr.Zero;

                    EnumChildWindows(mainHwnd, (childHwnd, l) =>
                    {
                        var cls = new StringBuilder(256);
                        var title = new StringBuilder(256);
                        GetClassName(childHwnd, cls, 256);
                        GetWindowText(childHwnd, title, 256);
                        string cName = cls.ToString();
                        string tName = title.ToString();

                        if (cName == "EXCEL7") hwndExcel7 = childHwnd;
                        if (cName.Contains("Chrome_RenderWidgetHostHWND")) hwndWebView = childHwnd;

                        if (cName.Contains("EXCEL") || cName.Contains("NetUI") || cName.Contains("WindowsForms") || cName.Contains("Chrome"))
                        {
                            foundChildWindows.Add(string.Format("0x{0:X8} Class='{1}', Title='{2}'", childHwnd.ToInt64(), cName, tName));
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (hwndExcel7 != IntPtr.Zero && hwndWebView != IntPtr.Zero)
                    {
                        Log(string.Format("  [就绪] 在第 {0} 次轮询时检测到 EXCEL7 与 WebView2 (Chrome_RenderWidgetHostHWND) 全部就绪！", poll + 1));
                        break;
                    }
                }

                Log(string.Format("发现关键子窗口 {0} 个 (EXCEL7=0x{1:X8}, WebView=0x{2:X8}):", 
                    foundChildWindows.Count, hwndExcel7.ToInt64(), hwndWebView.ToInt64()));
                foreach (var w in foundChildWindows) Log("  " + w);

                // 6. 核查特定进程的 COM 连接能力
                if (hwndExcel7 != IntPtr.Zero)
                {
                    Log("发现工作表网格窗口 EXCEL7: 0x" + hwndExcel7.ToString("X8"));
                    object pAcc = null;
                    Guid iid = IID_IDispatch;
                    int hr = AccessibleObjectFromWindow(hwndExcel7, OBJID_NATIVEOM, ref iid, out pAcc);
                    if (hr == 0 && pAcc != null)
                    {
                        dynamic win = pAcc;
                        dynamic app = win.Application;
                        Log(string.Format("[PASS] 成功通过 AccessibleObjectFromWindow 直连测试 Excel COM 实例: Version='{0}', Workbooks.Count={1}",
                            app.Version, app.Workbooks.Count));
                    }
                    else
                    {
                        Log(string.Format("[WARN] AccessibleObjectFromWindow 失败: hr=0x{0:X8}", hr));
                    }
                }

                // 7. UI Automation (UIA) 深度探测
                Log("\n--- UI Automation 探测 ---");
                AutomationElement rootElement = AutomationElement.FromHandle(mainHwnd);
                if (rootElement == null)
                {
                    Log("[FAIL] 无法从主窗口句柄创建 AutomationElement");
                }
                else
                {
                    Log(string.Format("[PASS] 成功获取主窗口 AutomationElement: Name='{0}'", rootElement.Current.Name));

                    // 探测 Ribbon
                    var ribbonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab);
                    var allTabs = rootElement.FindAll(TreeScope.Descendants, ribbonCondition);
                    Log(string.Format("发现功能区 Tab 控件数量: {0}", allTabs.Count));

                    bool foundExcelMindTab = false;
                    for (int t = 0; t < allTabs.Count; t++)
                    {
                        var tab = allTabs[t];
                        Log(string.Format("  Tab [{0}]: Name='{1}', AutomationId='{2}'", t, tab.Current.Name, tab.Current.AutomationId));
                        if (tab.Current.Name.Contains("ExcelMind") || tab.Current.Name.Contains("LeeExcel"))
                        {
                            foundExcelMindTab = true;
                        }
                    }
                    Log("ExcelMind AI 功能区 Tab 检测结果: " + (foundExcelMindTab ? "[PASS]" : "[WARN] 未在初始页签中激活"));

                    // 探测任务窗格 TaskPane 与 WebView2
                    Log("\n--- 任务窗格与 WebView2 内部控件探测 ---");
                    var paneCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane);
                    var allPanes = rootElement.FindAll(TreeScope.Descendants, paneCondition);
                    Log(string.Format("发现 Pane 控件数量: {0}", allPanes.Count));

                    AutomationElement taskPaneElem = null;
                    for (int p = 0; p < allPanes.Count; p++)
                    {
                        var pane = allPanes[p];
                        string pName = pane.Current.Name;
                        string cName = pane.Current.ClassName;
                        if (pName.Contains("ExcelMind") || cName.Contains("Chrome") || cName.Contains("TaskPane") || cName.Contains("WindowsForms"))
                        {
                            Log(string.Format("  候选任务窗格/WebView2 Pane: Name='{0}', Class='{1}', AutomationId='{2}'", pName, cName, pane.Current.AutomationId));
                            if (pName.Contains("ExcelMind") || cName.Contains("Chrome_WidgetWin_0"))
                            {
                                taskPaneElem = pane;
                            }
                        }
                    }

                    // 尝试深层遍历 WebView2 内部元素
                    if (taskPaneElem != null)
                    {
                        Log(string.Format("锁定目标任务窗格: Class='{0}', Name='{1}'", taskPaneElem.Current.ClassName, taskPaneElem.Current.Name));
                        var innerButtons = taskPaneElem.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                        Log(string.Format("任务窗格下发现 Button 控件数量: {0}", innerButtons.Count));
                        for (int b = 0; b < innerButtons.Count; b++)
                        {
                            var btn = innerButtons[b];
                            Log(string.Format("    Button [{0}]: Name='{1}', AutomationId='{2}'", b, btn.Current.Name, btn.Current.AutomationId));
                        }

                        var innerEdits = taskPaneElem.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                        Log(string.Format("任务窗格下发现 Edit/文本框 控件数量: {0}", innerEdits.Count));
                        for (int e = 0; e < innerEdits.Count; e++)
                        {
                            var ed = innerEdits[e];
                            Log(string.Format("    Edit [{0}]: Name='{1}', AutomationId='{2}'", e, ed.Current.Name, ed.Current.AutomationId));
                        }
                    }
                    else
                    {
                        Log("[WARN] 未直接锁定命名为 ExcelMind 的 TaskPane Pane，尝试全局搜索 WebView 内部元素...");
                        var docCond = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
                        var docs = rootElement.FindAll(TreeScope.Descendants, docCond);
                        Log(string.Format("全局发现 Document 控件数量: {0}", docs.Count));
                        for (int d = 0; d < docs.Count; d++)
                        {
                            Log(string.Format("  Document [{0}]: Name='{1}', Class='{2}'", d, docs[d].Current.Name, docs[d].Current.ClassName));
                        }
                    }
                }

                // 8. 截图保存验证证据 (处理最大化窗口坐标负数边界)
                try
                {
                    RECT r;
                    if (GetWindowRect(mainHwnd, out r))
                    {
                        int srcX = Math.Max(0, r.Left);
                        int srcY = Math.Max(0, r.Top);
                        int w = Math.Max(1, r.Right - srcX);
                        int h = Math.Max(1, r.Bottom - srcY);
                        if (w > 100 && h > 100)
                        {
                            using (var bmp = new Bitmap(w, h))
                            {
                                using (var g = Graphics.FromImage(bmp))
                                {
                                    g.CopyFromScreen(srcX, srcY, 0, 0, new Size(w, h));
                                }
                                bmp.Save(screenshotPath, ImageFormat.Png);
                                Log("[PASS] 窗口完整截图已保存: " + screenshotPath);
                            }
                        }
                    }
                }
                catch (Exception exScr)
                {
                    Log("[WARN] 截图生成警告: " + exScr.Message);
                }

                File.WriteAllText(logPath, _sbLog.ToString(), Encoding.UTF8);
                Log("\n能力核查报告已生成: " + logPath);
                return 0;
            }
            catch (Exception ex)
            {
                Log("[EXCEPTION] 探测异常: " + ex.Message);
                _sbLog.AppendLine(ex.ToString());
                File.WriteAllText(logPath, _sbLog.ToString(), Encoding.UTF8);
                return 1;
            }
            finally
            {
                // 9. 安全清理：仅且只能关闭本次测试创建的 testPid，绝不误触用户原有进程
                if (testProcess != null && !testProcess.HasExited)
                {
                    Log(string.Format("[清理] 正在安全退出本次测试 Excel 实例 (PID={0})...", testPid));
                    try
                    {
                        testProcess.CloseMainWindow();
                        Thread.Sleep(2000);
                        if (!testProcess.HasExited)
                        {
                            testProcess.Kill();
                        }
                        Log("[清理] 测试 Excel 实例已安全退出");
                    }
                    catch (Exception exClose)
                    {
                        Log("[清理异常] " + exClose.Message);
                    }
                }
            }
        }
    }
}
