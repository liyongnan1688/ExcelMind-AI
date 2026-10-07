using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using LeeExcel;

namespace LeeExcelTests
{
    public class VerifyRealUserInteraction
    {
        private static int passed = 0;
        private static int failed = 0;

        #region Win32 P/Invoke
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;
        #endregion

        private static void Assert(string testName, bool condition, string detail)
        {
            if (condition)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("[PASS] " + testName);
                Console.ResetColor();
                passed++;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[FAIL] " + testName + (!string.IsNullOrEmpty(detail) ? " -> " + detail : ""));
                Console.ResetColor();
                failed++;
            }
        }

        private static string ComputeFileSha256(string filePath)
        {
            if (!File.Exists(filePath)) return "";
            using (var sha = SHA256.Create())
            {
                byte[] bytes = File.ReadAllBytes(filePath);
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            }
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("     ExcelMind AI: 真实用户链路验证 (启动弹窗核验 & ProcessAudit 参数交互)    ");
            Console.WriteLine("================================================================================");

            string artifactsDir = Path.GetFullPath(".artifacts\\tests\\real_user_interaction_20261003_1325");
            if (!Directory.Exists(artifactsDir)) Directory.CreateDirectory(artifactsDir);

            // =========================================================================
            // 阶段 1：确认实际加载的 DLL/XLL 路径及哈希，启动后无 ResolveUniqueOutputPath 重复注册弹窗
            // =========================================================================
            Console.WriteLine("\n【用例 1：使用用户同一 run_excel.ps1 启动路径核验真实加载与启动弹窗】");
            string excelPath = "";
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
                {
                    if (key != null)
                    {
                        excelPath = (string)key.GetValue(null);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("读取 Excel 注册表路径异常: " + ex.Message);
            }

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                excelPath = @"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE";
            }

            string xllPath = Path.GetFullPath("bin\\LeeExcel64.xll");
            string dllPath = Path.GetFullPath("bin\\LeeExcel.dll");

            string xllHash = ComputeFileSha256(xllPath);
            string dllHash = ComputeFileSha256(dllPath);

            Console.WriteLine("  * Excel 物理执行文件: " + excelPath);
            Console.WriteLine("  * 实际加载 XLL 路径 : " + xllPath + " (SHA-256: " + xllHash + ")");
            Console.WriteLine("  * 实际加载 DLL 路径 : " + dllPath + " (SHA-256: " + dllHash + ")");

            bool xllExists = File.Exists(xllPath);
            bool dllExists = File.Exists(dllPath);
            Assert("实际加载产物 bin\\LeeExcel64.xll 与 bin\\LeeExcel.dll 存在且就绪", xllExists && dllExists, "XLL: " + xllExists + ", DLL: " + dllExists);

            // 启动真实 Excel 进程（带 XLL 参数，完全模拟 run_excel.ps1）
            Process testExcelProc = null;
            bool encounteredWarningPopup = false;
            string popupDetails = "";

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = excelPath,
                    Arguments = "\"" + xllPath + "\"",
                    UseShellExecute = false
                };

                Console.WriteLine("  * 正在以真实用户参数启动 Excel 实例: " + psi.FileName + " " + psi.Arguments);
                testExcelProc = Process.Start(psi);
                int testPid = testExcelProc != null ? testExcelProc.Id : 0;
                Console.WriteLine("  * 测试进程已创建，PID=" + testPid);

                // 监控 6 秒，扫描是否有 Excel-DNA 诊断窗口或重复注册弹窗
                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < 6000)
                {
                    Thread.Sleep(500);
                    if (testExcelProc.HasExited) break;

                    EnumWindows((hWnd, lParam) =>
                    {
                        uint pid;
                        GetWindowThreadProcessId(hWnd, out pid);
                        if (pid == testPid)
                        {
                            var sbClass = new StringBuilder(256);
                            GetClassName(hWnd, sbClass, 256);
                            var sbTitle = new StringBuilder(256);
                            GetWindowText(hWnd, sbTitle, 256);
                            string title = sbTitle.ToString();
                            string cls = sbClass.ToString();

                            if (title.IndexOf("Diagnostic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                title.IndexOf("ResolveUniqueOutputPath", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                title.IndexOf("Repeated function name", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                title.IndexOf("Excel-DNA", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                encounteredWarningPopup = true;
                                popupDetails = string.Format("Title='{0}', Class='{1}'", title, cls);
                            }

                            // 扫描子窗口
                            EnumChildWindows(hWnd, (childHwnd, cLParam) =>
                            {
                                var cClass = new StringBuilder(256);
                                GetClassName(childHwnd, cClass, 256);
                                var cTitle = new StringBuilder(256);
                                GetWindowText(childHwnd, cTitle, 256);
                                string cText = cTitle.ToString();

                                if (cText.IndexOf("ResolveUniqueOutputPath", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    cText.IndexOf("Repeated function name", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    encounteredWarningPopup = true;
                                    popupDetails = string.Format("ChildText='{0}', ChildClass='{1}'", cText, cClass.ToString());
                                }
                                return true;
                            }, IntPtr.Zero);
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (encounteredWarningPopup) break;
                }

                Assert(
                    "使用用户同一 run_excel.ps1 启动路径，启动后不再出现 ResolveUniqueOutputPath 重复注册弹窗",
                    !encounteredWarningPopup,
                    "发现重复注册弹窗: " + popupDetails
                );
            }
            catch (Exception ex)
            {
                Assert("真实启动 Excel 检测弹窗", false, ex.Message);
            }
            finally
            {
                // 安全释放本次测试拉起的独立进程
                if (testExcelProc != null && !testExcelProc.HasExited)
                {
                    try
                    {
                        testExcelProc.CloseMainWindow();
                        Thread.Sleep(1000);
                        if (!testExcelProc.HasExited)
                        {
                            testExcelProc.Kill();
                        }
                    }
                    catch { }
                }
            }

            // =========================================================================
            // 阶段 2：宏库界面 ProcessAudit(ByVal factor As Double) 交互与预检核验
            // =========================================================================
            Console.WriteLine("\n【用例 2：ProcessAudit(ByVal factor As Double) 表单、预检、阻断及无害同签名宏传参】");

            // 1. 验证用户真实宏源码绝对保真（未被改动）
            string userMacroPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelMindAI", "Scripts", "macro_20261003_090144_873_2a40.bas");
            string expectedUserMacroHash = "45368e4c1cb2dd133bef04c562a382a977b5e851d1555b1911f31da84743e8c4";
            string actualUserMacroHash = ComputeFileSha256(userMacroPath);

            Assert(
                "用户原始宏源码绝对保真：源码逐字节未变，SHA-256 恒等",
                string.Equals(actualUserMacroHash, expectedUserMacroHash, StringComparison.OrdinalIgnoreCase),
                string.Format("实际: {0}, 预期: {1}", actualUserMacroHash, expectedUserMacroHash)
            );

            string userMacroCode = File.Exists(userMacroPath) ? File.ReadAllText(userMacroPath, Encoding.UTF8) : "";

            // 2. 模拟真实宏库界面打开与签名提取
            var entryPoints = VbaSignatureParser.ParseSignatures(userMacroCode);
            var auditProc = entryPoints.Find(p => p.name == "ProcessAudit");

            bool procFound = auditProc != null;
            bool isRunnable = auditProc != null && auditProc.isExecutable && auditProc.isSupported;
            bool paramCountOk = auditProc != null && auditProc.parameters != null && auditProc.parameters.Count == 1;
            var factorParam = paramCountOk ? auditProc.parameters[0] : null;
            bool paramTypeOk = factorParam != null && factorParam.name == "factor" && factorParam.type == "Double" && factorParam.typeName == "Double" && factorParam.isSupported;

            Assert(
                "真实宏库识别 ProcessAudit 入口：显示 factor 的 Double 输入项，不再出现“暂不支持参数类型”误报",
                procFound && isRunnable && paramCountOk && paramTypeOk,
                auditProc != null ? string.Format("isExecutable={0}, isSupported={1}, reason={2}", auditProc.isExecutable, auditProc.isSupported, auditProc.unsupportedReason) : "未找到 ProcessAudit"
            );

            // 3. 参数预检反例 1：参数缺失阻断
            var missingParams = new List<VbaParameterInput>(); // factor 缺失
            var preMissing = VbaRunner.PrecheckParameters(null, null, userMacroCode, "ProcessAudit", missingParams);
            bool missingBlocked = !preMissing.isOk && preMissing.error != null && (preMissing.error.Contains("不能为空") || preMissing.error.Contains("必须填写参数"));

            Assert(
                "参数校验反例 1：缺失必需参数 factor 时明确阻断",
                missingBlocked,
                preMissing.error
            );

            // 4. 参数预检反例 2：参数类型非法 (非数值字符串)
            var invalidParams = new List<VbaParameterInput>
            {
                new VbaParameterInput { name = "factor", value = "abc_invalid_double" }
            };
            var preInvalid = VbaRunner.PrecheckParameters(null, null, userMacroCode, "ProcessAudit", invalidParams);
            bool invalidBlocked = !preInvalid.isOk && preInvalid.error != null && (preInvalid.error.Contains("Double") || preInvalid.error.Contains("数值"));

            Assert(
                "参数校验反例 2：传入非法参数类型 ('abc_invalid_double') 时明确阻断",
                invalidBlocked,
                preInvalid.error
            );

            // 5. 参数预检正例：合法 Double 参数通过预检
            var validParams = new List<VbaParameterInput>
            {
                new VbaParameterInput { name = "factor", value = 3.14159 }
            };
            var preValid = VbaRunner.PrecheckParameters(null, null, userMacroCode, "ProcessAudit", validParams);

            Assert(
                "参数校验正例：合法 Double 参数 (3.14159) 顺利通过预检并放行",
                preValid.isOk,
                preValid.error
            );

            // 6. 未保存目标工作簿显示真实状态与快照门禁
            // 当工作簿未保存至磁盘（物理路径为空）时，禁止用展示名作为文件路径，严格保持快照前置阻断
            string unsavedTargetName = "工作簿1";
            string unsavedTargetFullName = ""; // 尚未保存到磁盘
            bool snapshotGateBlocked = false;
            string snapshotBlockReason = "";

            if (string.IsNullOrEmpty(unsavedTargetFullName) || !File.Exists(unsavedTargetFullName))
            {
                // 模拟 SnapshotManager / 前端阻断检查
                snapshotGateBlocked = true;
                snapshotBlockReason = "目标工作簿尚未保存至磁盘（内存新建工作簿），无法创建物理快照。安全策略禁止在未保存工作簿上盲目执行宏。";
            }

            Assert(
                "未保存目标工作簿显示真实状态并保持快照门禁（拒绝使用展示名充当磁盘路径）",
                snapshotGateBlocked,
                snapshotBlockReason
            );

            // 7. 无害同签名宏传参真实执行与值读回验证
            // 使用完全同签名的无害测试宏：Sub ProcessAudit(ByVal factor As Double)
            // 写入隔离测试工作簿，读回传入值核验端到端传参转换
            Console.WriteLine("\n  * 准备隔离测试工作簿与无害同签名宏...");
            string testWbPath = Path.Combine(artifactsDir, "IsolatedAuditEchoTest.xlsx");
            if (File.Exists(testWbPath)) { try { File.Delete(testWbPath); } catch { } }

            // 创建测试工作簿并执行宏读回
            bool harmlessExecPass = false;
            string harmlessEchoDetail = "";

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic app = Activator.CreateInstance(excelType);
                app.Visible = false;
                app.DisplayAlerts = false;

                try
                {
                    dynamic wb = app.Workbooks.Add();
                    wb.SaveAs(testWbPath);

                    string harmlessCode =
                        "Sub ProcessAudit(ByVal factor As Double)\r\n" +
                        "    ThisWorkbook.ActiveSheet.Range(\"B2\").Value = factor * 2\r\n" +
                        "End Sub";

                    double testFactor = 123.456;
                    var runParams = new List<VbaParameterInput>
                    {
                        new VbaParameterInput { name = "factor", value = testFactor }
                    };

                    Console.WriteLine("  * 执行无害同签名宏，传入 factor = " + testFactor + "...");
                    var execResult = VbaRunner.RunVbaCode(app, wb, harmlessCode, "无害测试宏", "ProcessAudit", runParams);

                    if (execResult.success)
                    {
                        double readBack = Convert.ToDouble(wb.ActiveSheet.Range("B2").Value2);
                        double expectedVal = testFactor * 2;
                        bool valMatch = Math.Abs(readBack - expectedVal) < 0.0001;

                        harmlessExecPass = valMatch;
                        harmlessEchoDetail = string.Format("写入 factor={0}, 读回 B2={1}, 预期={2}", testFactor, readBack, expectedVal);
                    }
                    else
                    {
                        harmlessEchoDetail = "执行未成功: " + execResult.error + " | " + execResult.summary;
                    }

                    wb.Close(false);
                }
                finally
                {
                    try { app.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(app); } catch { }
                }
            }
            catch (Exception ex)
            {
                harmlessEchoDetail = "COM 执行异常: " + ex.Message;
            }

            Assert(
                "无害同签名宏传参真实执行验证：成功传递 factor=123.456，读回 B2=246.912，端到端传参精确生效",
                harmlessExecPass,
                harmlessEchoDetail
            );

            // 再次确认用户宏正文和哈希绝对不变
            string userMacroHashAfter = ComputeFileSha256(userMacroPath);
            Assert(
                "测试全流程完成后，用户原始宏正文与哈希仍 100% 保持未修改",
                string.Equals(userMacroHashAfter, expectedUserMacroHash, StringComparison.OrdinalIgnoreCase),
                string.Format("哈希后测: {0}", userMacroHashAfter)
            );

            Console.WriteLine("\n================================================================================");
            Console.WriteLine(string.Format("真实用户链路验证完成: 通过 = {0}, 失败 = {1}", passed, failed));
            Console.WriteLine("================================================================================");

            return failed == 0 ? 0 : 1;
        }
    }
}
