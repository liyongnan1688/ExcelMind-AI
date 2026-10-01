using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public class MemoryDumper
{
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern int VirtualQueryEx(IntPtr hProcess, IntPtr lpAddress, out MEMORY_BASIC_INFORMATION lpBuffer, uint dwLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [Out] byte[] lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);

    [StructLayout(LayoutKind.Sequential)]
    struct MEMORY_BASIC_INFORMATION
    {
        public IntPtr BaseAddress;
        public IntPtr AllocationBase;
        public uint AllocationProtect;
        public IntPtr RegionSize;
        public uint State;
        public uint Protect;
        public uint Type;
    }

    const int PROCESS_QUERY_INFORMATION = 0x0400;
    const int PROCESS_VM_READ = 0x0010;
    const uint MEM_COMMIT = 0x1000;
    const uint PAGE_NOACCESS = 0x01;
    const uint PAGE_GUARD = 0x100;

    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        int[] targetPids = new int[] { 14184, 16600, 19728, 3364, 18536, 30668, 18972 };
        string[] searchPatterns = new string[] { "3c80eca4", "3e262634", "ws.Cells(1, 1).Select" };

        Console.WriteLine("=== 开始扫描进程内存 ===");
        int hitCount = 0;

        foreach (int pid in targetPids)
        {
            Process p = null;
            try { p = Process.GetProcessById(pid); } catch { continue; }
            Console.WriteLine(string.Format("扫描进程: {0} (PID: {1})...", p.ProcessName, pid));

            IntPtr hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, pid);
            if (hProcess == IntPtr.Zero)
            {
                Console.WriteLine(string.Format("  无法打开进程 {0}", pid));
                continue;
            }

            try
            {
                long address = 0;
                while (address < 0x7FFFFFFF0000L)
                {
                    MEMORY_BASIC_INFORMATION mbi;
                    int res = VirtualQueryEx(hProcess, (IntPtr)address, out mbi, (uint)Marshal.SizeOf(typeof(MEMORY_BASIC_INFORMATION)));
                    if (res == 0) break;

                    long regionSize = mbi.RegionSize.ToInt64();
                    if (mbi.State == MEM_COMMIT && (mbi.Protect & PAGE_GUARD) == 0 && (mbi.Protect & PAGE_NOACCESS) == 0)
                    {
                        // 只读合理大小的内存块 (<= 32MB)
                        if (regionSize > 0 && regionSize <= 32 * 1024 * 1024)
                        {
                            byte[] buffer = new byte[regionSize];
                            IntPtr bytesRead;
                            if (ReadProcessMemory(hProcess, mbi.BaseAddress, buffer, (int)regionSize, out bytesRead))
                            {
                                int readLen = bytesRead.ToInt32();
                                // 分别检查 UTF-8 与 UTF-16
                                string utf8 = null;
                                try { utf8 = Encoding.UTF8.GetString(buffer, 0, readLen); } catch { }
                                string unicode = null;
                                try { unicode = Encoding.Unicode.GetString(buffer, 0, readLen); } catch { }

                                foreach (string pat in searchPatterns)
                                {
                                    bool foundInUtf8 = utf8 != null && utf8.Contains(pat);
                                    bool foundInUni = unicode != null && unicode.Contains(pat);

                                    if (foundInUtf8 || foundInUni)
                                    {
                                        hitCount++;
                                        string source = foundInUtf8 ? utf8 : unicode;
                                        int idx = source.IndexOf(pat);
                                        int start = Math.Max(0, idx - 1500);
                                        int length = Math.Min(source.Length - start, 4000);
                                        string snippet = source.Substring(start, length);

                                        string outDir = args != null && args.Length > 0 ? args[0] : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../.artifacts/tests/memory_dumps");
                                        outDir = Path.GetFullPath(outDir);
                                        if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);
                                        string dumpFileName = string.Format("dump_hit_{0}_pid{1}_{2}.txt", hitCount, pid, pat.Replace('.', '_').Replace(' ', '_').Replace('(', '_').Replace(')', '_'));
                                        string dumpFile = Path.Combine(outDir, dumpFileName);
                                        File.WriteAllText(dumpFile, snippet, Encoding.UTF8);
                                        Console.WriteLine(string.Format("  [命中 {0}] PID {1} 发现 '{2}' -> 导出至 {3}", hitCount, pid, pat, dumpFile));
                                    }
                                }
                            }
                        }
                    }

                    address += regionSize;
                    if (address <= 0) break;
                }
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        Console.WriteLine(string.Format("=== 扫描结束，共命中 {0} 处 ===", hitCount));
    }
}
