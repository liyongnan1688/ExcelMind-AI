using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace LeeExcelTests
{
    public class ProbeExcelProcess
    {
        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        public const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        public static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

        public static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            var procs = Process.GetProcessesByName("EXCEL");
            Console.WriteLine("Excel 进程数量: " + procs.Length);

            foreach (var proc in procs)
            {
                Console.WriteLine("\n--- 进程 PID: " + proc.Id + " | 响应: " + proc.Responding + " ---");
                try
                {
                    foreach (ProcessModule mod in proc.Modules)
                    {
                        if (mod.ModuleName.IndexOf("LeeExcel", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine("  [Module] " + mod.ModuleName + " -> " + mod.FileName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  模块读取异常: " + ex.Message);
                }

                var foundWindows = new List<string>();
                IntPtr excel7Hwnd = IntPtr.Zero;

                EnumWindows((hWnd, lParam) =>
                {
                    uint pid;
                    GetWindowThreadProcessId(hWnd, out pid);
                    if (pid == proc.Id)
                    {
                        var sbClass = new StringBuilder(256);
                        GetClassName(hWnd, sbClass, 256);
                        var sbTitle = new StringBuilder(256);
                        GetWindowText(hWnd, sbTitle, 256);
                        bool vis = IsWindowVisible(hWnd);
                        foundWindows.Add(string.Format("TopWindow: hWnd=0x{0:X}, Class={1}, Visible={2}, Title={3}", (long)hWnd, sbClass, vis, sbTitle));

                        EnumChildWindows(hWnd, (childHwnd, cLParam) =>
                        {
                            var cClass = new StringBuilder(256);
                            GetClassName(childHwnd, cClass, 256);
                            var cTitle = new StringBuilder(256);
                            GetWindowText(childHwnd, cTitle, 256);
                            bool cVis = IsWindowVisible(childHwnd);
                            foundWindows.Add(string.Format("  Child: hWnd=0x{0:X}, Class={1}, Visible={2}, Title={3}", (long)childHwnd, cClass, cVis, cTitle));

                            if (cClass.ToString() == "EXCEL7")
                            {
                                excel7Hwnd = childHwnd;
                            }
                            return true;
                        }, IntPtr.Zero);
                    }
                    return true;
                }, IntPtr.Zero);

                Console.WriteLine("发现窗口数: " + foundWindows.Count);
                foreach (var w in foundWindows)
                {
                    Console.WriteLine("  " + w);
                }

                if (excel7Hwnd != IntPtr.Zero)
                {
                    Console.WriteLine("找到 EXCEL7 句柄: 0x" + excel7Hwnd.ToString("X"));
                    object nativeObj = null;
                    Guid iid = IID_IDispatch;
                    int hr = AccessibleObjectFromWindow(excel7Hwnd, OBJID_NATIVEOM, ref iid, out nativeObj);
                    if (hr == 0 && nativeObj != null)
                    {
                        dynamic win = nativeObj;
                        dynamic app = win.Application;
                        Console.WriteLine("COM 挂接成功! Excel Version: " + app.Version + ", Workbooks.Count: " + app.Workbooks.Count);
                    }
                    else
                    {
                        Console.WriteLine("AccessibleObjectFromWindow 失败，hr = 0x" + hr.ToString("X"));
                    }
                }
            }
        }
    }
}
