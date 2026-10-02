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
    public class CaptureRibbonEvidence
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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
            Console.WriteLine("        ExcelMind AI: 真实 Excel 功能区 UI Automation 自动化取证工具         ");
            Console.WriteLine("================================================================================");

            string artifactsDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(artifactsDir))
            {
                Directory.CreateDirectory(artifactsDir);
            }

            string screenshotPath = Path.Combine(artifactsDir, "excel_ribbon_macro_management.png");
            string logPath = Path.Combine(artifactsDir, "ribbon_automation_evidence.txt");
            var sbLog = new StringBuilder();
            sbLog.AppendLine("=== ExcelMind AI 真实功能区 UI Automation 探测证据 ===");
            sbLog.AppendLine("测试时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 1. 寻找 Excel.exe 路径与插件 xll
            string excelPath = null;
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
                {
                    if (key != null)
                    {
                        excelPath = key.GetValue(null) as string;
                    }
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

            Console.WriteLine("Excel 路径: " + excelPath);
            Console.WriteLine("插件路径: " + xllPath);
            sbLog.AppendLine("Excel 路径: " + excelPath);
            sbLog.AppendLine("插件路径: " + xllPath);

            if (!File.Exists(xllPath))
            {
                Console.WriteLine("[ERROR] 未找到插件文件: " + xllPath);
                return 1;
            }

            Process proc = null;
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = excelPath,
                    Arguments = "\"" + xllPath + "\"",
                    UseShellExecute = false
                };

                Console.WriteLine("正在启动 Excel 并挂载加载项...");
                proc = Process.Start(psi);
                sbLog.AppendLine("Excel 进程 PID: " + proc.Id);

                // 等待主窗口并初始化
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

                Console.WriteLine("Excel 主窗口句柄: " + hwnd);
                sbLog.AppendLine("Excel 主窗口句柄: " + hwnd);
                Thread.Sleep(3000); // 留足 TaskPane 与 Ribbon 加载时间

                // 确保有打开的工作簿，使 Ribbon 选项卡完全展现
                try
                {
                    object excelApp = Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null)
                    {
                        object workbooks = excelApp.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);
                        int count = (int)workbooks.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, workbooks, null);
                        if (count == 0)
                        {
                            workbooks.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, null);
                            Console.WriteLine("[INFO] 已通过 COM 创建空白工作簿以完全展开功能区");
                        }

                        // 获取真实当前主窗口 HWND
                        try
                        {
                            int appHwnd = (int)excelApp.GetType().InvokeMember("Hwnd", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);
                            if (appHwnd != 0)
                            {
                                hwnd = new IntPtr(appHwnd);
                                Console.WriteLine("[INFO] 获取到当前 Excel COM 主窗口句柄: " + hwnd);
                            }
                        }
                        catch { }
                    }
                }
                catch (Exception exCom)
                {
                    Console.WriteLine("[WARN] COM 激活工作簿: " + exCom.Message);
                }

                ShowWindow(hwnd, 3); // SW_MAXIMIZE
                SetForegroundWindow(hwnd);
                Thread.Sleep(2000);

                // 2. UI Automation 探测
                bool foundDiagErrorDialog = false;
                string diagErrorDetails = "";
                bool foundTab = false;
                bool foundAssistantBtn = false;
                bool foundMacroLibBtn = false;
                bool foundImportBtn = false;
                bool foundSettingsBtn = false;
                bool foundAssistantGroup = false;
                bool foundMacroToolsGroup = false;
                bool foundSettingsGroup = false;

                // 检查是否有 Excel-DNA 诊断弹窗
                try
                {
                    var diagCond = new PropertyCondition(AutomationElement.NameProperty, "Excel-DNA Diagnostic Display");
                    var diagWin = AutomationElement.RootElement.FindFirst(TreeScope.Children, diagCond);
                    if (diagWin != null)
                    {
                        foundDiagErrorDialog = true;
                        diagErrorDetails = "发现 Excel-DNA Diagnostic Display 弹窗！";
                        Console.WriteLine("[ERROR] " + diagErrorDetails);
                        sbLog.AppendLine("[ERROR] " + diagErrorDetails);
                    }
                }
                catch { }

                List<AutomationElement> searchRoots = new List<AutomationElement>();
                try
                {
                    var rootFromHwnd = AutomationElement.FromHandle(hwnd);
                    if (rootFromHwnd != null) searchRoots.Add(rootFromHwnd);
                }
                catch { }

                try
                {
                    var procCond = new PropertyCondition(AutomationElement.ProcessIdProperty, proc.Id);
                    var appWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, procCond);
                    foreach (AutomationElement win in appWindows)
                    {
                        if (!searchRoots.Contains(win)) searchRoots.Add(win);
                    }
                }
                catch { }

                sbLog.AppendLine("\n搜索根元素数量: " + searchRoots.Count);

                foreach (var excelWindow in searchRoots)
                {
                    try
                    {
                        // 查找 TabItem
                        var allTabs = excelWindow.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                        foreach (AutomationElement t in allTabs)
                        {
                            try
                            {
                                string tName = t.Current.Name;
                                sbLog.AppendLine("  - 发现选项卡: " + tName);
                                if (!string.IsNullOrEmpty(tName) && tName.IndexOf("ExcelMind", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    foundTab = true;
                                    Console.WriteLine("[CONFIRMED] 成功定位到功能区选项卡: " + tName);
                                    sbLog.AppendLine("[CONFIRMED] 成功定位到功能区选项卡: " + tName);

                                    var selPattern = t.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                    if (selPattern != null)
                                    {
                                        selPattern.Select();
                                        Console.WriteLine("[CONFIRMED] 已激活 ExcelMind AI 选项卡");
                                        sbLog.AppendLine("[CONFIRMED] 已激活 ExcelMind AI 选项卡");
                                        Thread.Sleep(1500);
                                    }
                                    break;
                                }
                            }
                            catch { }
                        }

                        // 查找按钮
                        var buttons = excelWindow.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                        foreach (AutomationElement btn in buttons)
                        {
                            try
                            {
                                string bName = btn.Current.Name;
                                if (!string.IsNullOrEmpty(bName))
                                {
                                    sbLog.AppendLine("  - 按钮: " + bName);
                                    if (bName == "ExcelMind AI" || bName == "AI助手") foundAssistantBtn = true;
                                    if (bName == "宏库") foundMacroLibBtn = true;
                                    if (bName == "导入") foundImportBtn = true;
                                    if (bName == "API设置") foundSettingsBtn = true;
                                }
                            }
                            catch { }
                        }

                        // 查找分组
                        var groups = excelWindow.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group));
                        foreach (AutomationElement g in groups)
                        {
                            try
                            {
                                string gName = g.Current.Name;
                                if (!string.IsNullOrEmpty(gName))
                                {
                                    sbLog.AppendLine("  - 分组: " + gName);
                                    if (gName == "智能助手") foundAssistantGroup = true;
                                    if (gName == "宏工具") foundMacroToolsGroup = true;
                                    if (gName == "设置") foundSettingsGroup = true;
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception exElem)
                    {
                        sbLog.AppendLine("[WARN] 元素探测异常: " + exElem.Message);
                    }
                }

                // 3. 截屏
                try
                {
                    RECT rect;
                    if (GetWindowRect(hwnd, out rect))
                    {
                        int width = Math.Max(1, rect.Right - rect.Left);
                        int height = Math.Max(1, rect.Bottom - rect.Top);
                        if (width > 300 && height > 200)
                        {
                            using (var bmp = new Bitmap(width, height))
                            {
                                using (var gfx = Graphics.FromImage(bmp))
                                {
                                    IntPtr hdc = gfx.GetHdc();
                                    bool printed = PrintWindow(hwnd, hdc, 2); // PW_RENDERFULLCONTENT
                                    if (!printed)
                                    {
                                        printed = PrintWindow(hwnd, hdc, 0);
                                    }
                                    gfx.ReleaseHdc(hdc);

                                    if (!printed)
                                    {
                                        try
                                        {
                                            int srcX = Math.Max(0, rect.Left);
                                            int srcY = Math.Max(0, rect.Top);
                                            gfx.CopyFromScreen(srcX, srcY, 0, 0, new Size(width, height));
                                        }
                                        catch { }
                                    }
                                }
                                bmp.Save(screenshotPath, ImageFormat.Png);
                                string secondPath = Path.Combine(artifactsDir, "excel_ribbon_new_layout.png");
                                bmp.Save(secondPath, ImageFormat.Png);
                                Console.WriteLine("[CONFIRMED] 功能区窗口完整截图已保存: " + screenshotPath);
                                sbLog.AppendLine("[CONFIRMED] 功能区窗口完整截图已保存: " + screenshotPath);
                            }
                        }
                    }
                }
                catch (Exception exBmp)
                {
                    Console.WriteLine("[WARN] 截图处理: " + exBmp.Message);
                    sbLog.AppendLine("[WARN] 截图处理: " + exBmp.Message);
                }

                Console.WriteLine("\n--- 功能区元素验证总结 ---");
                Console.WriteLine("ExcelMind AI 选项卡: " + (foundTab ? "已检测到 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("Excel-DNA 注册异常弹窗: " + (!foundDiagErrorDialog ? "无异常弹窗 [PASS]" : "异常出现 [FAIL]"));
                Console.WriteLine("分组【智能助手】: " + (foundAssistantGroup ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("分组【宏工具】: " + (foundMacroToolsGroup ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("分组【设置】: " + (foundSettingsGroup ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("按钮【AI助手】: " + (foundAssistantBtn ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("按钮【宏库】: " + (foundMacroLibBtn ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("按钮【导入】: " + (foundImportBtn ? "已就绪 [PASS]" : "未检测到 [FAIL]"));
                Console.WriteLine("按钮【API设置】: " + (foundSettingsBtn ? "已就绪 [PASS]" : "未检测到 [FAIL]"));

                sbLog.AppendLine("\n--- 关键验收指标判定 ---");
                sbLog.AppendLine("ExcelMind AI 选项卡存在: " + foundTab);
                sbLog.AppendLine("Excel-DNA 注册诊断异常: " + foundDiagErrorDialog);
                sbLog.AppendLine("智能助手分组存在: " + foundAssistantGroup);
                sbLog.AppendLine("宏工具分组存在: " + foundMacroToolsGroup);
                sbLog.AppendLine("设置分组存在: " + foundSettingsGroup);
                sbLog.AppendLine("AI助手按钮存在: " + foundAssistantBtn);
                sbLog.AppendLine("宏库按钮存在: " + foundMacroLibBtn);
                sbLog.AppendLine("导入按钮存在: " + foundImportBtn);
                sbLog.AppendLine("API设置按钮存在: " + foundSettingsBtn);

                File.WriteAllText(logPath, sbLog.ToString(), Encoding.UTF8);
                Console.WriteLine("取证证据报告已写入: " + logPath);

                bool allPass = foundTab && !foundDiagErrorDialog && foundAssistantGroup && foundMacroToolsGroup && foundSettingsGroup && foundAssistantBtn && foundMacroLibBtn && foundImportBtn && foundSettingsBtn;
                return allPass ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EXCEPTION] 自动化取证异常: " + ex.Message);
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
                        Console.WriteLine("正在关闭测试 Excel 进程...");
                        proc.CloseMainWindow();
                        Thread.Sleep(2000);
                        if (!proc.HasExited) proc.Kill();
                    }
                }
                catch { }
            }
        }
    }
}
