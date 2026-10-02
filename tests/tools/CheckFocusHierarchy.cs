using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace LeeExcel.Diagnostics
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public System.Drawing.Rectangle rcCaret;
        }

        [DllImport("user32.dll")]
        static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        static string GetWndInfo(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return "(null)";
            StringBuilder cls = new StringBuilder(256);
            StringBuilder title = new StringBuilder(256);
            GetClassName(hWnd, cls, 256);
            GetWindowText(hWnd, title, 256);
            return string.Format("0x{0:X8} [Class: '{1}', Title: '{2}']", hWnd.ToInt64(), cls.ToString(), title.ToString());
        }

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumChildProc lpEnumFunc, IntPtr lParam);

        static void Main(string[] args)
        {
            Process[] procs = Process.GetProcessesByName("EXCEL");
            if (procs.Length == 0)
            {
                Console.WriteLine("没有找到运行中的 Excel 进程。");
                return;
            }

            Process excel = procs[0];
            Console.WriteLine("Excel PID: " + excel.Id);

            EnumWindows((hWnd, lParam) =>
            {
                uint pid = 0;
                GetWindowThreadProcessId(hWnd, out pid);
                if (pid == excel.Id)
                {
                    Console.WriteLine("\n[顶层窗口] " + GetWndInfo(hWnd));
                    EnumChildWindows(hWnd, (childHwnd, l) =>
                    {
                        StringBuilder cls = new StringBuilder(256);
                        GetClassName(childHwnd, cls, 256);
                        string className = cls.ToString();
                        if (className == "EXCEL7" || className == "XLDESK" || className == "EXCEL<" || 
                            className.Contains("Chrome") || className.Contains("Task") || className.Contains("NetUI") ||
                            className.Contains("WindowsForms"))
                        {
                            StringBuilder title = new StringBuilder(256);
                            GetWindowText(childHwnd, title, 256);
                            Console.WriteLine("    -> 0x{0:X8} Class='{1}', Title='{2}'", childHwnd.ToInt64(), className, title.ToString());
                        }
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
        }
    }
}
