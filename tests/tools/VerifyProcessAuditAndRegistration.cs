using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using LeeExcel;

namespace LeeExcelTests
{
    public class VerifyProcessAuditAndRegistration
    {
        private static int passed = 0;
        private static int failed = 0;

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

        private static bool IsExcelDnaEligibleFunction(MethodInfo method)
        {
            if (!method.IsPublic || !method.IsStatic) return false;
            if (method.IsSpecialName) return false;
            if (method.IsGenericMethod || method.IsGenericMethodDefinition) return false;

            // Excel-DNA 无法导出包含 ref 或 out 参数的方法为工作表函数
            foreach (var p in method.GetParameters())
            {
                if (p.IsOut || p.ParameterType.IsByRef) return false;
            }

            // Excel-DNA 仅对 Excel 兼容的返回值类型注册为工作表函数
            Type ret = method.ReturnType;
            if (ret == typeof(string) || ret == typeof(double) || ret == typeof(int) ||
                ret == typeof(bool) || ret == typeof(DateTime) || ret == typeof(object) ||
                ret == typeof(short) || ret == typeof(long) || ret == typeof(float) ||
                ret == typeof(decimal) || ret == typeof(string[]) || ret == typeof(double[]) ||
                ret == typeof(object[]) || ret == typeof(double[,]) || ret == typeof(object[,]))
            {
                return true;
            }

            return false;
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            Console.WriteLine("=======================================================");
            Console.WriteLine("定向验证：Excel-DNA 重复注册修复与 ProcessAudit 签名解析");
            Console.WriteLine("=======================================================");

            // -------------------------------------------------------------------------
            // 验证 1：Excel-DNA 导出 UDF 反射核验（无 ResolveUniqueOutputPath 重复注册）
            // -------------------------------------------------------------------------
            try
            {
                string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LeeExcel.dll");
                if (!File.Exists(dllPath))
                {
                    dllPath = Path.GetFullPath("bin\\LeeExcel.dll");
                }
                Assembly asm = Assembly.LoadFrom(dllPath);

                // 检查 DataToolsService.ResolveUniqueOutputPath
                Type dtType = asm.GetType("LeeExcel.DataToolsService");
                MethodInfo dtMethod = dtType != null ? dtType.GetMethod("ResolveUniqueOutputPath", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static) : null;
                bool dtIsInternal = dtMethod != null && !dtMethod.IsPublic;

                // 检查 BatchRunnerService.ResolveUniqueOutputPath
                Type brType = asm.GetType("LeeExcel.BatchRunnerService");
                MethodInfo brMethod = brType != null ? brType.GetMethod("ResolveUniqueOutputPath", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static) : null;
                bool brIsInternal = brMethod != null && !brMethod.IsPublic;

                Assert(
                    "ResolveUniqueOutputPath 访问控制收敛为 internal（非 public static，防止 Excel-DNA 自动导出为 Excel 函数）",
                    dtIsInternal && brIsInternal,
                    string.Format("DataTools: isPublic={0}, BatchRunner: isPublic={1}", dtMethod != null ? dtMethod.IsPublic.ToString() : "null", brMethod != null ? brMethod.IsPublic.ToString() : "null")
                );

                // 统计所有 public class 的 public static 方法（模拟 Excel-DNA 自动注册扫描）
                var exportedFunctionNames = new List<string>();
                var duplicates = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var type in asm.GetExportedTypes())
                {
                    // 排除 ExcelRibbon 和 IExcelAddIn
                    if (typeof(ExcelDna.Integration.CustomUI.ExcelRibbon).IsAssignableFrom(type)) continue;
                    if (typeof(ExcelDna.Integration.IExcelAddIn).IsAssignableFrom(type)) continue;

                    foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                    {
                        if (!IsExcelDnaEligibleFunction(method)) continue;

                        string mName = method.Name;
                        if (seen.Contains(mName))
                        {
                            duplicates.Add(type.Name + "." + mName);
                        }
                        else
                        {
                            seen.Add(mName);
                            exportedFunctionNames.Add(mName);
                        }
                    }
                }

                Assert(
                    "Excel-DNA 自动导出全局函数名无重名冲突（重复注册列表为空，杜绝 Repeated function name 弹窗）",
                    !seen.Contains("ResolveUniqueOutputPath") && duplicates.Count == 0,
                    "重复列表: " + string.Join(", ", duplicates.ToArray())
                );
            }
            catch (Exception ex)
            {
                Assert("Excel-DNA 导出 UDF 反射核验", false, ex.Message);
            }

            // -------------------------------------------------------------------------
            // 验证 2：ProcessAudit(ByVal factor As Double) 签名与双向契约
            // -------------------------------------------------------------------------
            try
            {
                string processAuditCode = "Attribute VB_Name = \"Module1\"\r\n" +
                    "' 财务高精度报表与审计计算\r\n" +
                    "Sub ProcessAudit(ByVal factor As Double)\r\n" +
                    "    Dim token As String\r\n" +
                    "    token = \"ID: 1000000000000000001, PI: 3.14159265358979323846, Exp: 1.23456789e18\"\r\n" +
                    "    MsgBox \"Result: \" & token\r\n" +
                    "End Sub";

                var entryPoints = VbaSignatureParser.ParseSignatures(processAuditCode);

                bool isCount1 = entryPoints.Count == 1;
                var ep = entryPoints.Count > 0 ? entryPoints[0] : null;
                bool isNameMatch = ep != null && ep.name == "ProcessAudit";
                bool isSub = ep != null && ep.kind == "Sub";
                bool isExec = ep != null && ep.isExecutable && ep.isSupported;
                bool paramOk = ep != null && ep.parameters != null && ep.parameters.Count == 1;
                var p = paramOk ? ep.parameters[0] : null;
                bool pMatch = p != null && p.name == "factor" && p.type == "Double" && p.typeName == "Double" && p.isSupported;

                Assert(
                    "ProcessAudit(ByVal factor As Double) 生产签名正确解析为可运行宏且参数为 Double",
                    isCount1 && isNameMatch && isSub && isExec && paramOk && pMatch,
                    ep != null ? string.Format("isExecutable={0}, isSupported={1}, reason={2}", ep.isExecutable, ep.isSupported, ep.unsupportedReason) : "null"
                );
            }
            catch (Exception ex)
            {
                Assert("ProcessAudit 签名解析", false, ex.Message);
            }

            // -------------------------------------------------------------------------
            // 验证 3：确实不支持的签名给出具体参数名、类型及原因
            // -------------------------------------------------------------------------
            try
            {
                string unsupportedCode = "Sub ComplexProc(ByVal col As Collection, ByRef flags() As Long)\r\n    MsgBox \"x\"\r\nEnd Sub";
                var unsupEntries = VbaSignatureParser.ParseSignatures(unsupportedCode);
                var uEp = unsupEntries.Count > 0 ? unsupEntries[0] : null;

                bool uNotExec = uEp != null && !uEp.isExecutable && !uEp.isSupported;
                bool uReasonDetail = uEp != null &&
                    uEp.unsupportedReason != null &&
                    uEp.unsupportedReason.Contains("col") &&
                    uEp.unsupportedReason.Contains("Collection") &&
                    uEp.unsupportedReason.Contains("flags") &&
                    uEp.unsupportedReason.Contains("数组参数");

                Assert(
                    "确实不支持的入口明确给出具体参数名、类型及原因（非笼统提示）",
                    uNotExec && uReasonDetail,
                    uEp != null ? uEp.unsupportedReason : "null"
                );
            }
            catch (Exception ex)
            {
                Assert("不支持签名具体原因诊断", false, ex.Message);
            }

            // -------------------------------------------------------------------------
            // 验证 4：用户真实宏源码字节与哈希恒定未变
            // -------------------------------------------------------------------------
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string macroPath = Path.Combine(appData, "ExcelMindAI", "Scripts", "macro_20261003_090144_873_2a40.bas");
                if (File.Exists(macroPath))
                {
                    byte[] rawBytes = File.ReadAllBytes(macroPath);
                    string hash;
                    using (var sha = SHA256.Create())
                    {
                        hash = BitConverter.ToString(sha.ComputeHash(rawBytes)).Replace("-", "").ToUpperInvariant();
                    }

                    // 预期哈希 (已在定位调查时固化)
                    string expectedHash = "45368E4C1CB2DD133BEF04C562A382A977B5E851D1555B1911F31DA84743E8C4";
                    bool hashMatch = string.Equals(hash, expectedHash, StringComparison.OrdinalIgnoreCase);

                    Assert(
                        "用户原始宏源码绝对保真：源码逐字节未变，SHA-256 哈希恒等",
                        hashMatch,
                        string.Format("实际: {0}, 预期: {1}", hash, expectedHash)
                    );
                }
                else
                {
                    Console.WriteLine("[INFO] 用户宏文件不存在于当前测试机路径，跳过物理文件哈希校验");
                }
            }
            catch (Exception ex)
            {
                Assert("用户宏源码保真", false, ex.Message);
            }

            Console.WriteLine("\n-------------------------------------------------------");
            Console.WriteLine(string.Format("定向验证完成: 通过 = {0}, 失败 = {1}", passed, failed));
            Console.WriteLine("-------------------------------------------------------");

            return failed == 0 ? 0 : 1;
        }
    }
}
