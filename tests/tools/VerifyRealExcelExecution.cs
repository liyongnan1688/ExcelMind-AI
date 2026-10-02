using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using LeeExcel;

namespace LeeExcelTests
{
    public class VerifyRealExcelExecution
    {
        private static int _passCount = 0;
        private static int _failCount = 0;
        private static StringBuilder _log = new StringBuilder();

        private static void LogAndAssert(bool condition, string testName, string details = "")
        {
            string line = (condition ? "[PASS] " : "[FAIL] ") + testName + (string.IsNullOrEmpty(details) ? "" : " -> " + details);
            Console.WriteLine(line);
            _log.AppendLine(line);
            if (condition) _passCount++;
            else _failCount++;
        }

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("================================================================================");
            Console.WriteLine("   ExcelMind AI: 真实 Excel 宿主 COM 宏运行、快照隔离与入口调度全链验证   ");
            Console.WriteLine("================================================================================");

            string artifactsDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".artifacts", "tests", "macro_verification"));
            if (!Directory.Exists(artifactsDir))
            {
                Directory.CreateDirectory(artifactsDir);
            }

            string testWbPath = Path.Combine(artifactsDir, "IsolatedRealMacroTest.xlsx");
            if (File.Exists(testWbPath))
            {
                try { File.Delete(testWbPath); } catch { }
            }

            dynamic app = null;
            dynamic wb = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Console.WriteLine("[ERROR] 本机未检测到 Microsoft Excel 安装 (ProgID: Excel.Application 未注册)");
                    return 2;
                }

                app = Activator.CreateInstance(excelType);
                app.Visible = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Add();
                dynamic sheet1 = wb.Worksheets[1];
                sheet1.Name = "DataSheet";
                sheet1.Cells[1, 1].Value2 = "初始未修改状态";
                wb.SaveAs(testWbPath);

                LogAndAssert(File.Exists(testWbPath), "已创建真实测试隔离工作簿", testWbPath);

                // ====================================================================
                // 1. 无害测试宏保存与首次执行 (无需 API Key，不调用大模型)
                // ====================================================================
                Console.WriteLine("\n--- 1. 真实运行测试宏与首份物理快照生成 ---");
                string macro1Code =
                    "Option Explicit\r\n" +
                    "Public Sub FillSalesSummary()\r\n" +
                    "    Dim ws As Worksheet\r\n" +
                    "    Set ws = ActiveSheet\r\n" +
                    "    ws.Range(\"A1\").Value = \"季度汇总\"\r\n" +
                    "    ws.Range(\"B1\").Value = \"销售额(元)\"\r\n" +
                    "    ws.Range(\"A2\").Value = \"2026年Q1\"\r\n" +
                    "    ws.Range(\"B2\").Value = 880000\r\n" +
                    "End Sub\r\n";

                string macro1Id;
                string saveErr;
                bool saved = ScriptManager.SaveScript(
                    id: null,
                    displayName: "季度销售汇总宏",
                    code: macro1Code,
                    description: "真实Excel执行测试",
                    category: "日常测试",
                    sourceType: "paste",
                    originalFileName: "",
                    encoding: "UTF-8",
                    entryPoint: "FillSalesSummary",
                    rawBytesBase64: "",
                    overwrite: false,
                    savedId: out macro1Id,
                    error: out saveErr
                );

                LogAndAssert(saved, "测试宏已成功入库保存", "ID=" + macro1Id);

                // 发送 execute_vba 请求 (第 1 次运行)
                var req1 = new Dictionary<string, string>
                {
                    { "action", "execute_vba" },
                    { "requestId", "run_test_1" },
                    { "code", macro1Code },
                    { "prompt", "运行宏: 季度销售汇总宏" },
                    { "targetWorkbookName", Path.GetFileName(testWbPath) },
                    { "targetWorkbookFullName", testWbPath },
                    { "entryPoint", "FillSalesSummary" },
                    { "scriptId", macro1Id }
                };

                string resp1Json = NativeBridge.Dispatch(SimpleJson.Serialize(req1), app);
                var resp1 = SimpleJson.ParseFlatObject(resp1Json);

                LogAndAssert(resp1.ContainsKey("ok") && resp1["ok"] == "true", "NativeBridge 第 1 次运行宏执行成功", resp1.ContainsKey("message") ? resp1["message"] : "");

                // 读回工作簿单元格实际写入值
                object valA1 = sheet1.Cells[1, 1].Value2;
                object valB2 = sheet1.Cells[2, 2].Value2;
                string cellA1 = valA1 != null ? valA1.ToString() : null;
                string cellB2 = valB2 != null ? valB2.ToString() : null;
                LogAndAssert(cellA1 == "季度汇总" && cellB2 == "880000", "目标工作簿实际单元格值已成功写入", string.Format("A1='{0}', B2='{1}'", cellA1, cellB2));

                // 检查快照列表
                var snapList1 = SnapshotManager.LoadSnapshots(testWbPath);
                LogAndAssert(snapList1.Count >= 1, "第 1 次运行成功创建独立的物理快照", "快照数=" + snapList1.Count + ", ID=" + snapList1[0].id);
                string snap1Id = snapList1[0].id;

                // 验证 ScriptManager 中的执行状态被同步更新为'执行成功'
                var scriptItem1 = ScriptManager.ListScripts().Find(s => s.id == macro1Id);
                LogAndAssert(scriptItem1 != null && scriptItem1.lastExecutionResult.Contains("执行成功"), "宏条目的最近一次执行结果成功持久化为'执行成功'");

                // ====================================================================
                // 2. 同一宏连续运行第 2 次：产生全新的独立快照，不复用上次快照
                // ====================================================================
                Console.WriteLine("\n--- 2. 同一宏连续运行第 2 次快照独立性验证 ---");
                System.Threading.Thread.Sleep(1100); // 确保时间戳区分

                // 宏代码稍作调整写入第2行
                string macro2Code =
                    "Option Explicit\r\n" +
                    "Public Sub FillSalesSummary()\r\n" +
                    "    Range(\"A3\").Value = \"2026年Q2\"\r\n" +
                    "    Range(\"B3\").Value = 960000\r\n" +
                    "End Sub\r\n";

                var req2 = new Dictionary<string, string>
                {
                    { "action", "execute_vba" },
                    { "requestId", "run_test_2" },
                    { "code", macro2Code },
                    { "prompt", "第2次运行宏: 季度销售汇总宏" },
                    { "targetWorkbookName", Path.GetFileName(testWbPath) },
                    { "targetWorkbookFullName", testWbPath },
                    { "entryPoint", "FillSalesSummary" },
                    { "scriptId", macro1Id }
                };

                string resp2Json = NativeBridge.Dispatch(SimpleJson.Serialize(req2), app);
                var resp2 = SimpleJson.ParseFlatObject(resp2Json);
                LogAndAssert(resp2.ContainsKey("ok") && resp2["ok"] == "true", "第 2 次运行宏执行成功");

                var snapList2 = SnapshotManager.LoadSnapshots(testWbPath);
                LogAndAssert(snapList2.Count >= 2, "第 2 次运行生成了独立的新快照", "当前总快照数=" + snapList2.Count);

                string snap2Id = snapList2[0].id; // 降序最新
                LogAndAssert(snap1Id != snap2Id, "两次执行快照 ID 完全独立不同 (严禁复用上次快照)", string.Format("Snap1={0}, Snap2={1}", snap1Id, snap2Id));

                // ====================================================================
                // 3. 多入口过程指定执行与参数拦截
                // ====================================================================
                Console.WriteLine("\n--- 3. 多入口过程指定执行验证 ---");
                string multiProcVba =
                    "Public Sub TaskAlpha()\r\n" +
                    "    Range(\"C1\").Value = \"AlphaDone\"\r\n" +
                    "End Sub\r\n\r\n" +
                    "Public Sub TaskBeta()\r\n" +
                    "    Range(\"D1\").Value = \"BetaDone\"\r\n" +
                    "End Sub\r\n\r\n" +
                    "Public Sub TaskWithArg(arg As String)\r\n" +
                    "    Range(\"E1\").Value = arg\r\n" +
                    "End Sub\r\n";

                // 指定入口为 TaskBeta，预期只执行 TaskBeta，写入 D1="BetaDone"，而 C1 保持空
                var reqBeta = new Dictionary<string, string>
                {
                    { "action", "execute_vba" },
                    { "requestId", "run_test_beta" },
                    { "code", multiProcVba },
                    { "prompt", "运行指定入口: TaskBeta" },
                    { "targetWorkbookName", Path.GetFileName(testWbPath) },
                    { "targetWorkbookFullName", testWbPath },
                    { "entryPoint", "TaskBeta" }
                };

                string respBetaJson = NativeBridge.Dispatch(SimpleJson.Serialize(reqBeta), app);
                var respBeta = SimpleJson.ParseFlatObject(respBetaJson);
                LogAndAssert(respBeta.ContainsKey("ok") && respBeta["ok"] == "true", "指定多入口宏中的 TaskBeta 成功执行");

                object valD1 = sheet1.Cells[1, 4].Value2;
                object valC1 = sheet1.Cells[1, 3].Value2;
                string cellD1 = valD1 != null ? valD1.ToString() : null;
                string cellC1 = valC1 != null ? valC1.ToString() : null;
                LogAndAssert(cellD1 == "BetaDone" && string.IsNullOrEmpty(cellC1), "仅指定的 TaskBeta 执行生效 (D1='BetaDone', C1 为空未被误调)");

                // 指定必填参数过程 TaskWithArg，预期安全拦截拒绝
                var reqWithArg = new Dictionary<string, string>
                {
                    { "action", "execute_vba" },
                    { "requestId", "run_test_arg" },
                    { "code", multiProcVba },
                    { "prompt", "运行带参过程: TaskWithArg" },
                    { "targetWorkbookName", Path.GetFileName(testWbPath) },
                    { "targetWorkbookFullName", testWbPath },
                    { "entryPoint", "TaskWithArg" }
                };
                string respArgJson = NativeBridge.Dispatch(SimpleJson.Serialize(reqWithArg), app);
                var respArg = SimpleJson.ParseFlatObject(respArgJson);
                LogAndAssert(respArg["ok"] == "false" && respArg["error"].Contains("包含必填参数"), "指定带必填参数入口被精准拦截并明确报错", respArg["error"]);

                // ====================================================================
                // 4. 清理测试产生的宏
                // ====================================================================
                string cleanErr;
                ScriptManager.DeleteScript(macro1Id, out cleanErr);

                // 保存完整测试报告
                string reportPath = Path.Combine(artifactsDir, "real_excel_macro_verification_report.txt");
                File.WriteAllText(reportPath, _log.ToString(), Encoding.UTF8);

                Console.WriteLine("\n================================================================================");
                Console.WriteLine(string.Format("真实 Excel COM 执行验证完成: 通过 = {0}, 失败 = {1}", _passCount, _failCount));
                Console.WriteLine("验证报告已写入: " + reportPath);
                Console.WriteLine("================================================================================");

                return _failCount == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[FATAL EXCEPTION] 真实 Excel 运行验证发生未捕获异常: " + ex.ToString());
                _log.AppendLine("[FATAL EXCEPTION] " + ex.ToString());
                return 1;
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                        wb = null;
                    }
                }
                catch { }

                try
                {
                    if (app != null)
                    {
                        app.Quit();
                        Marshal.ReleaseComObject(app);
                        app = null;
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
