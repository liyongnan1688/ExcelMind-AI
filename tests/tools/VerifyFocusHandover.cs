using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using LeeExcel;

namespace LeeExcel.Tests
{
    class Program
    {
        static int _passed = 0;
        static int _failed = 0;

        static void Assert(bool condition, string testName, string detail = "")
        {
            if (condition)
            {
                Console.WriteLine("[PASS] " + testName);
                _passed++;
            }
            else
            {
                Console.WriteLine("[FAIL] " + testName + (string.IsNullOrEmpty(detail) ? "" : (" -> " + detail)));
                _failed++;
            }
        }

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, resolveArgs) =>
            {
                string assemblyName = new AssemblyName(resolveArgs.Name).Name + ".dll";
                string binPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "bin", assemblyName);
                if (File.Exists(binPath))
                {
                    return Assembly.LoadFrom(binPath);
                }
                string rootBin = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "bin", assemblyName));
                if (File.Exists(rootBin))
                {
                    return Assembly.LoadFrom(rootBin);
                }
                return null;
            };

            RunTests();
        }

        static void RunTests()
        {
            Console.WriteLine("=== TaskPane 与 Excel 焦点交接机制单元校验 ===");

            // 1. IsChildOrSame 基础边界校验
            Assert(TaskPaneFocusHelper.IsChildOrSame(new IntPtr(100), new IntPtr(100)) == true, 
                "IsChildOrSame: 相同句柄返回 true");
            Assert(TaskPaneFocusHelper.IsChildOrSame(IntPtr.Zero, new IntPtr(100)) == false, 
                "IsChildOrSame: 父句柄为 Zero 返回 false");
            Assert(TaskPaneFocusHelper.IsChildOrSame(new IntPtr(100), IntPtr.Zero) == false, 
                "IsChildOrSame: 目标句柄为 Zero 返回 false");
            Assert(TaskPaneFocusHelper.IsChildOrSame(IntPtr.Zero, IntPtr.Zero) == false, 
                "IsChildOrSame: 双 Zero 返回 false");

            // 2. 真实 WinForms 控件树判定校验
            using (Form parentForm = new Form())
            {
                IntPtr hParent = parentForm.Handle; // 触发句柄创建
                using (Panel childPanel = new Panel())
                {
                    childPanel.Parent = parentForm;
                    IntPtr hChild = childPanel.Handle;

                    using (TextBox textInPanel = new TextBox())
                    {
                        textInPanel.Parent = childPanel;
                        IntPtr hText = textInPanel.Handle;

                        using (Form unrelatedForm = new Form())
                        {
                            IntPtr hUnrelated = unrelatedForm.Handle;

                            Assert(TaskPaneFocusHelper.IsChildOrSame(hParent, hChild) == true,
                                "IsChildOrSame: Panel 是 Form 的子窗口");
                            Assert(TaskPaneFocusHelper.IsChildOrSame(hParent, hText) == true,
                                "IsChildOrSame: TextBox 是 Form 的孙窗口");
                            Assert(TaskPaneFocusHelper.IsChildOrSame(hParent, hUnrelated) == false,
                                "IsChildOrSame: 独立 Form 不是子窗口");

                            // 3. TaskPaneFocusMessageFilter 消息过滤验证
                            TaskPaneFocusMessageFilter filter = new TaskPaneFocusMessageFilter(hParent, null);

                            // 验证所有消息无论何时都返回 false (绝不吞掉 Excel 原生点击或编辑消息)
                            Message clickInPane = Message.Create(hText, 0x0201, IntPtr.Zero, IntPtr.Zero); // WM_LBUTTONDOWN 内部
                            Assert(filter.PreFilterMessage(ref clickInPane) == false,
                                "PreFilterMessage: 窗格内点击消息返回 false (不吞掉)");

                            Message clickOutside = Message.Create(hUnrelated, 0x0201, IntPtr.Zero, IntPtr.Zero); // WM_LBUTTONDOWN 外部
                            Assert(filter.PreFilterMessage(ref clickOutside) == false,
                                "PreFilterMessage: 窗格外点击消息返回 false (不吞掉)");

                            Message rClickOutside = Message.Create(hUnrelated, 0x0204, IntPtr.Zero, IntPtr.Zero); // WM_RBUTTONDOWN 外部
                            Assert(filter.PreFilterMessage(ref rClickOutside) == false,
                                "PreFilterMessage: 窗格外右键消息返回 false (不吞掉)");

                            Message keyMsg = Message.Create(hText, 0x0100, IntPtr.Zero, IntPtr.Zero); // WM_KEYDOWN
                            Assert(filter.PreFilterMessage(ref keyMsg) == false,
                                "PreFilterMessage: 键盘按键消息返回 false (无干涉)");

                            // 4. 异常安全边界测试
                            try
                            {
                                TaskPaneFocusHelper.RelinquishFocusToExcel(IntPtr.Zero, null);
                                TaskPaneFocusHelper.RelinquishFocusToExcel(parentForm.Handle, null, IntPtr.Zero);
                                Assert(true, "RelinquishFocusToExcel: 句柄为空或 app 为空时不抛出异常 (抗崩溃保护)");
                            }
                            catch (Exception ex)
                            {
                                Assert(false, "RelinquishFocusToExcel: 异常未捕获", ex.Message);
                            }

                            try
                            {
                                IntPtr activeSheet = TaskPaneFocusHelper.GetActiveWorksheetHwnd(null);
                                Assert(activeSheet == IntPtr.Zero, "GetActiveWorksheetHwnd: app 为空时安全返回 Zero");
                            }
                            catch (Exception ex)
                            {
                                Assert(false, "GetActiveWorksheetHwnd: 异常未捕获", ex.Message);
                            }
                        }
                    }
                }
            }

            Console.WriteLine(string.Format("\n焦点机制测试结果汇总: 通过 {0} 项, 失败 {1} 项", _passed, _failed));
            if (_failed > 0)
            {
                Environment.Exit(1);
            }
        }
    }
}
