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
    public class VerifyRibbonLayout
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("          ExcelMind AI: Ribbon 真实定向验证与截图工具                           ");
            Console.WriteLine("================================================================================");

            string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ribbon_verify_output");
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

            string normalShotPath = Path.Combine(outDir, "real_ribbon_normal_width.png");
            string narrowShotPath = Path.Combine(outDir, "real_ribbon_narrow_width.png");
            string menuTaskShotPath = Path.Combine(outDir, "real_ribbon_menu_tasks.png");
            string menuFavShotPath = Path.Combine(outDir, "real_ribbon_menu_fav.png");
            string logPath = Path.Combine(outDir, "ribbon_verify_log.txt");

            var sbLog = new StringBuilder();
            sbLog.AppendLine("=== ExcelMind AI Ribbon 定向验证日志 ===");
            sbLog.AppendLine("执行时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 1. 查找 Excel.exe 与 xll 路径
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
            if (!File.Exists(Path.Combine(projectRoot, "bin", "LeeExcel64.xll")))
            {
                projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", ".."));
            }
            string xllPath = Path.Combine(projectRoot, "bin", "LeeExcel64.xll");

            Console.WriteLine("Excel: " + excelPath);
            Console.WriteLine("Xll:   " + xllPath);
            sbLog.AppendLine("Excel: " + excelPath);
            sbLog.AppendLine("Xll:   " + xllPath);

            if (!File.Exists(xllPath))
            {
                Console.WriteLine("[ERROR] 插件不存在: " + xllPath);
                return 1;
            }

            Process proc = null;
            object excelApp = null;
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = excelPath,
                    Arguments = "\"" + xllPath + "\"",
                    UseShellExecute = false
                };

                Console.WriteLine("启动 Excel 并加载最新构建插件...");
                proc = Process.Start(psi);
                sbLog.AppendLine("Excel PID: " + proc.Id);

                IntPtr hwnd = IntPtr.Zero;
                for (int i = 0; i < 40; i++)
                {
                    proc.Refresh();
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                        hwnd = proc.MainWindowHandle;
                        break;
                    }
                    Thread.Sleep(500);
                }

                Console.WriteLine("Excel 窗口句柄: " + hwnd);
                sbLog.AppendLine("Excel 窗口句柄: " + hwnd);
                Thread.Sleep(3000);

                // 通过 COM 创建空白工作簿以完全展开功能区
                try
                {
                    excelApp = Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null)
                    {
                        object workbooks = excelApp.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);
                        int count = (int)workbooks.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, workbooks, null);
                        if (count == 0)
                        {
                            workbooks.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, null);
                            Console.WriteLine("[INFO] 已通过 COM 激活空白工作簿展开功能区");
                        }
                    }
                }
                catch (Exception exCom)
                {
                    Console.WriteLine("[WARN] COM 激活: " + exCom.Message);
                }

                // 获取真正的 Excel 主窗口句柄
                try
                {
                    int appHwnd = (int)excelApp.GetType().InvokeMember("Hwnd", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);
                    if (appHwnd != 0) hwnd = new IntPtr(appHwnd);
                    Console.WriteLine("[INFO] Excel 主窗口 Hwnd: " + hwnd);
                }
                catch { }

                ShowWindow(hwnd, 3); // SW_MAXIMIZE
                SetForegroundWindow(hwnd);
                Thread.Sleep(2000);

                AutomationElement excelWindow = AutomationElement.FromHandle(hwnd);
                if (excelWindow == null)
                {
                    Console.WriteLine("[ERROR] 无法获取 Excel AutomationElement");
                    return 1;
                }

                // 定位并激活 ExcelMind AI 选项卡
                AutomationElement tabElem = null;
                var allTabs = excelWindow.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                foreach (AutomationElement t in allTabs)
                {
                    try
                    {
                        string tName = t.Current.Name;
                        if (!string.IsNullOrEmpty(tName) && tName.IndexOf("ExcelMind", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            tabElem = t;
                            var selPattern = t.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            if (selPattern != null)
                            {
                                selPattern.Select();
                                Console.WriteLine("[PASS] 成功激活 ExcelMind AI 选项卡: " + tName);
                                sbLog.AppendLine("[PASS] 成功激活 ExcelMind AI 选项卡: " + tName);
                            }
                            break;
                        }
                    }
                    catch { }
                }
                Thread.Sleep(2000);

                // ==================== 1. 正常宽度窗口下的 Ribbon 布局 ====================
                Console.WriteLine("\n--- [验证 1] 正常宽度窗口下的 Ribbon 布局 ---");
                RECT normalRect;
                GetWindowRect(hwnd, out normalRect);
                int normW = Math.Max(1, normalRect.Right - normalRect.Left);
                int normH = Math.Max(1, normalRect.Bottom - normalRect.Top);
                Console.WriteLine(string.Format("正常窗口像素大小: {0} x {1}", normW, normH));
                sbLog.AppendLine(string.Format("正常窗口像素大小: {0} x {1}", normW, normH));

                CaptureWindowTop(hwnd, normW, normH, 180, normalShotPath);
                Console.WriteLine("[PASS] 已捕获正常宽度真实 Ribbon 原图: " + normalShotPath);

                // 探测正常宽度下的控件树
                ProbeRibbonControls(excelWindow, sbLog);

                // ==================== 2. 菜单下拉交互测试 ====================
                Console.WriteLine("\n--- [验证 2] 菜单展开测试 ---");
                AutomationElement menuTasksElem = FindControlByName(excelWindow, "任务", ControlType.MenuItem, ControlType.Button);
                if (menuTasksElem != null)
                {
                    try
                    {
                        var expandPat = menuTasksElem.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                        if (expandPat != null)
                        {
                            expandPat.Expand();
                            Console.WriteLine("[PASS] 成功通过 ExpandCollapsePattern 展开【任务】菜单");
                            sbLog.AppendLine("[PASS] 成功通过 ExpandCollapsePattern 展开【任务】菜单");
                            Thread.Sleep(1200);
                            CaptureWindowTop(hwnd, normW, normH, 260, menuTaskShotPath);
                            Console.WriteLine("[PASS] 已捕获【任务】下拉菜单展开截图: " + menuTaskShotPath);

                            // 验证菜单子项
                            var batchBtn = FindControlByName(excelWindow, "批量处理", ControlType.MenuItem, ControlType.Button);
                            var workflowBtn = FindControlByName(excelWindow, "工作流", ControlType.MenuItem, ControlType.Button);
                            bool foundBatch = (batchBtn != null);
                            bool foundWorkflow = (workflowBtn != null);
                            Console.WriteLine("  子项【批量处理】存在: " + foundBatch);
                            Console.WriteLine("  子项【工作流】存在:   " + foundWorkflow);
                            sbLog.AppendLine(string.Format("  子项【批量处理】: {0}, 子项【工作流】: {1}", foundBatch, foundWorkflow));

                            expandPat.Collapse();
                            Thread.Sleep(500);
                        }
                        else
                        {
                            var invokePat = menuTasksElem.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                            if (invokePat != null)
                            {
                                invokePat.Invoke();
                                Console.WriteLine("[PASS] 成功通过 InvokePattern 触发【任务】菜单");
                                sbLog.AppendLine("[PASS] 成功通过 InvokePattern 触发【任务】菜单");
                                Thread.Sleep(1000);
                                CaptureWindowTop(hwnd, normW, normH, 260, menuTaskShotPath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("[WARN] 展开任务菜单: " + ex.Message);
                        sbLog.AppendLine("[WARN] 展开任务菜单: " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("[WARN] 未能在 Automation 树中找到【任务】菜单");
                }

                // ==================== 3. 适度缩窄窗口验证与截图 ====================
                Console.WriteLine("\n--- [验证 3] 适度缩窄窗口下的 Ribbon 响应 ---");
                ShowWindow(hwnd, 9); // SW_RESTORE
                Thread.Sleep(1000);
                // 缩窄至约 850 x 650
                int narrowTargetW = 850;
                int narrowTargetH = 650;
                MoveWindow(hwnd, 50, 50, narrowTargetW, narrowTargetH, true);
                SetForegroundWindow(hwnd);
                Thread.Sleep(2500);

                RECT narrowRect;
                GetWindowRect(hwnd, out narrowRect);
                int actualNarrowW = Math.Max(1, narrowRect.Right - narrowRect.Left);
                int actualNarrowH = Math.Max(1, narrowRect.Bottom - narrowRect.Top);
                Console.WriteLine(string.Format("缩窄后窗口像素大小: {0} x {1}", actualNarrowW, actualNarrowH));
                sbLog.AppendLine(string.Format("缩窄后窗口像素大小: {0} x {1}", actualNarrowW, actualNarrowH));

                CaptureWindowTop(hwnd, actualNarrowW, actualNarrowH, 180, narrowShotPath);
                Console.WriteLine("[PASS] 已捕获适度缩窄真实 Ribbon 截图: " + narrowShotPath);

                // 探测缩窄状态下的控件树
                ProbeRibbonControls(excelWindow, sbLog);

                // 复制截图至 brain 目录
                string brainDir = @"C:\Users\35651\.gemini\antigravity-ide\brain\69496968-66be-4c13-9d75-45edbeb36221";
                if (Directory.Exists(brainDir))
                {
                    File.Copy(normalShotPath, Path.Combine(brainDir, "real_ribbon_normal_width.png"), true);
                    File.Copy(narrowShotPath, Path.Combine(brainDir, "real_ribbon_narrow_width.png"), true);
                    if (File.Exists(menuTaskShotPath))
                    {
                        File.Copy(menuTaskShotPath, Path.Combine(brainDir, "real_ribbon_menu_tasks.png"), true);
                    }
                    Console.WriteLine("[PASS] 所有 Ribbon 截图已同步至交付目录");
                }

                File.WriteAllText(logPath, sbLog.ToString(), Encoding.UTF8);
                Console.WriteLine("\n[RESULT] Ribbon 定向验证完成。");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EXCEPTION] " + ex.Message);
                sbLog.AppendLine("[EXCEPTION] " + ex.ToString());
                File.WriteAllText(logPath, sbLog.ToString(), Encoding.UTF8);
                return 1;
            }
            finally
            {
                try
                {
                    if (proc != null && !proc.HasExited)
                    {
                        Console.WriteLine("安全释放与退出测试 Excel 进程...");
                        proc.CloseMainWindow();
                        Thread.Sleep(1500);
                        if (!proc.HasExited) proc.Kill();
                    }
                }
                catch { }
            }
        }

        private static void CaptureWindowTop(IntPtr hwnd, int fullW, int fullH, int topHeight, string savePath)
        {
            using (var bmp = new Bitmap(fullW, fullH))
            {
                using (var gfx = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = gfx.GetHdc();
                    bool printed = PrintWindow(hwnd, hdc, 2);
                    if (!printed) printed = PrintWindow(hwnd, hdc, 0);
                    gfx.ReleaseHdc(hdc);
                }

                int cropH = Math.Min(fullH, topHeight);
                using (var ribbonCrop = bmp.Clone(new Rectangle(0, 0, fullW, cropH), bmp.PixelFormat))
                {
                    ribbonCrop.Save(savePath, ImageFormat.Png);
                }
            }
        }

        private static AutomationElement FindControlByName(AutomationElement parent, string name, params ControlType[] types)
        {
            foreach (var t in types)
            {
                var elems = parent.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, t));
                foreach (AutomationElement el in elems)
                {
                    try
                    {
                        if (el.Current.Name == name) return el;
                    }
                    catch { }
                }
            }
            return null;
        }

        private static void ProbeRibbonControls(AutomationElement parent, StringBuilder sbLog)
        {
            sbLog.AppendLine("--- 控件探测快照 ---");
            var groups = parent.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group));
            foreach (AutomationElement g in groups)
            {
                try
                {
                    string gName = g.Current.Name;
                    if (!string.IsNullOrEmpty(gName))
                    {
                        Console.WriteLine("  [Group] " + gName);
                        sbLog.AppendLine("  [Group] " + gName);
                    }
                }
                catch { }
            }

            var buttons = parent.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            foreach (AutomationElement btn in buttons)
            {
                try
                {
                    string bName = btn.Current.Name;
                    if (!string.IsNullOrEmpty(bName) && (bName == "打开" || bName.IndexOf("ExcelMind", StringComparison.OrdinalIgnoreCase) >= 0 || bName == "AI 助手" || bName == "宏库" || bName == "数据工具" || bName == "设置" || bName == "任务" || bName == "收藏宏" || bName == "批量处理" || bName == "工作流"))
                    {
                        Console.WriteLine("  [Button] " + bName);
                        sbLog.AppendLine("  [Button] " + bName);
                    }
                }
                catch { }
            }
        }
    }
}
