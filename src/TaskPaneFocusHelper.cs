using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace LeeExcel
{
    /// <summary>
    /// 任务窗格与 Excel 宿主之间的 Win32 焦点平滑交接辅助类
    /// 解决 WebView2 获得焦点后，鼠标点击 Excel 单元格无法夺回键盘输入的交互缺陷
    /// </summary>
    public static class TaskPaneFocusHelper
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>
        /// 判断 target 是否为 parent 本身或其子孙窗口（含 IsChild 与 GetParent 递归回溯）
        /// </summary>
        public static bool IsChildOrSame(IntPtr parent, IntPtr target)
        {
            if (parent == IntPtr.Zero || target == IntPtr.Zero) return false;
            if (parent == target) return true;
            if (IsChild(parent, target)) return true;

            // 向上回溯父窗口树
            IntPtr cur = target;
            for (int i = 0; i < 64; i++)
            {
                cur = GetParent(cur);
                if (cur == IntPtr.Zero) break;
                if (cur == parent) return true;
            }

            return false;
        }

        /// <summary>
        /// 获取窗口类名
        /// </summary>
        public static string GetWindowClassName(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return string.Empty;
            StringBuilder sb = new StringBuilder(256);
            GetClassName(hWnd, sb, 256);
            return sb.ToString();
        }

        /// <summary>
        /// 查找当前 Excel 实例中可见的 EXCEL7（工作表网格）窗口句柄
        /// </summary>
        public static IntPtr GetActiveWorksheetHwnd(dynamic excelApp)
        {
            try
            {
                if (excelApp == null) return IntPtr.Zero;
                int mainHwndInt = 0;
                try { mainHwndInt = (int)excelApp.Hwnd; } catch { }
                if (mainHwndInt == 0) return IntPtr.Zero;
                IntPtr xlMain = new IntPtr(mainHwndInt);

                // 1. 标准 MDI 结构遍历：XLMAIN -> XLDESK -> EXCEL< -> EXCEL7
                IntPtr xlDesk = FindWindowEx(xlMain, IntPtr.Zero, "XLDESK", null);
                if (xlDesk != IntPtr.Zero)
                {
                    IntPtr excelBook = IntPtr.Zero;
                    while ((excelBook = FindWindowEx(xlDesk, excelBook, "EXCEL<", null)) != IntPtr.Zero)
                    {
                        if (IsWindowVisible(excelBook))
                        {
                            IntPtr excel7 = FindWindowEx(excelBook, IntPtr.Zero, "EXCEL7", null);
                            if (excel7 != IntPtr.Zero && IsWindowVisible(excel7))
                            {
                                return excel7;
                            }
                        }
                    }
                }

                // 2. 递归枚举备选方案 (适用于 SDI 或自定义窗口层级)
                IntPtr foundExcel7 = IntPtr.Zero;
                EnumChildWindows(xlMain, (hWnd, lParam) =>
                {
                    if (IsWindowVisible(hWnd) && GetWindowClassName(hWnd) == "EXCEL7")
                    {
                        foundExcel7 = hWnd;
                        return false; // 找到即停止枚举
                    }
                    return true;
                }, IntPtr.Zero);

                if (foundExcel7 != IntPtr.Zero)
                {
                    return foundExcel7;
                }

                // 3. 保底返回 XLMAIN 自身
                return xlMain;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[TaskPaneFocusHelper] GetActiveWorksheetHwnd 异常: " + ex.Message);
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 仅当焦点当前确实停留在任务窗格内时，将键盘焦点平滑归还给 Excel（工作表或目标控件）
        /// </summary>
        public static void RelinquishFocusToExcel(IntPtr taskPaneHwnd, dynamic excelApp, IntPtr preferredTargetHwnd = default(IntPtr))
        {
            try
            {
                if (taskPaneHwnd == IntPtr.Zero) return;

                IntPtr currentFocus = GetFocus();
                if (currentFocus == IntPtr.Zero) return;

                // 只有当当前拥有 Win32 焦点的窗口确实是任务窗格或其子控件时才交出焦点
                if (!IsChildOrSame(taskPaneHwnd, currentFocus))
                {
                    return;
                }

                // 若指定了目标窗口（且不在任务窗格内部），直接激活目标窗口
                if (preferredTargetHwnd != IntPtr.Zero && IsWindow(preferredTargetHwnd) && !IsChildOrSame(taskPaneHwnd, preferredTargetHwnd))
                {
                    SetFocus(preferredTargetHwnd);
                    return;
                }

                // 否则定位活动工作表窗口
                IntPtr sheetHwnd = GetActiveWorksheetHwnd(excelApp);
                if (sheetHwnd != IntPtr.Zero)
                {
                    SetFocus(sheetHwnd);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[TaskPaneFocusHelper] RelinquishFocusToExcel 异常: " + ex.Message);
            }
        }
    }

    /// <summary>
    /// 进程内 WinForms 消息过滤器：
    /// 当用户点击任务窗格外部的 Excel 界面（如单元格、公式栏、Ribbon）时，安全交出焦点
    /// 永远不吞掉任何消息，保持 Excel 原生交互 100% 顺畅
    /// </summary>
    public class TaskPaneFocusMessageFilter : IMessageFilter
    {
        private readonly IntPtr _taskPaneHwnd;
        private readonly dynamic _excelApp;

        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int WM_NCRBUTTONDOWN = 0x00A4;

        public TaskPaneFocusMessageFilter(IntPtr taskPaneHwnd, dynamic excelApp)
        {
            _taskPaneHwnd = taskPaneHwnd;
            _excelApp = excelApp;
        }

        public bool PreFilterMessage(ref Message m)
        {
            try
            {
                if (m.Msg == WM_LBUTTONDOWN || m.Msg == WM_RBUTTONDOWN || m.Msg == WM_MBUTTONDOWN ||
                    m.Msg == WM_NCLBUTTONDOWN || m.Msg == WM_NCRBUTTONDOWN)
                {
                    IntPtr targetHwnd = m.HWnd;
                    // 如果鼠标点击落在任务窗格外
                    if (!TaskPaneFocusHelper.IsChildOrSame(_taskPaneHwnd, targetHwnd))
                    {
                        TaskPaneFocusHelper.RelinquishFocusToExcel(_taskPaneHwnd, _excelApp, targetHwnd);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[TaskPaneFocusMessageFilter] 过滤异常: " + ex.Message);
            }

            return false; // 绝不拦截吞掉任何消息
        }
    }
}
