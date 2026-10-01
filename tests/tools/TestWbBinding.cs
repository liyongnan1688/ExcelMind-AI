using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace LeeExcel.Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("=== Excel Interop Application.Run 与 Workbook 传参专项实测 (C# 原生) ===");

            string outputDir = args != null && args.Length > 0 ? args[0] : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../.artifacts/tests/wb_binding_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            outputDir = Path.GetFullPath(outputDir);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            string targetPath = Path.Combine(outputDir, "exp_target.xlsx");
            string otherPath = Path.Combine(outputDir, "exp_other.xlsx");
            string logPath = Path.Combine(outputDir, "exp_binding_result.txt");

            if (File.Exists(targetPath)) File.Delete(targetPath);
            if (File.Exists(otherPath)) File.Delete(otherPath);
            if (File.Exists(logPath)) File.Delete(logPath);

            var report = new StringBuilder();
            Action<string> log = (msg) =>
            {
                Console.WriteLine(msg);
                report.AppendLine(msg);
            };

            dynamic excel = null;
            dynamic wbTarget = null;
            dynamic wbOther = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                excel = Activator.CreateInstance(excelType);
                excel.Visible = false;
                excel.DisplayAlerts = false;

                log("Excel 版本: " + excel.Version + " (Build: " + excel.Build + ")");
                log("测试执行时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // 1. 创建目标工作簿和干扰工作簿
                wbTarget = excel.Workbooks.Add();
                wbTarget.Sheets[1].Range("A1").Value = "TARGET_INIT";
                wbTarget.SaveAs(targetPath);

                wbOther = excel.Workbooks.Add();
                wbOther.Sheets[1].Range("A1").Value = "OTHER_INIT";
                wbOther.SaveAs(otherPath);

                string targetFullName = (string)wbTarget.FullName;
                string otherFullName = (string)wbOther.FullName;
                string targetName = (string)wbTarget.Name;
                string otherName = (string)wbOther.Name;

                log("目标工作簿 (Target): " + targetFullName);
                log("干扰工作簿 (Other) : " + otherFullName);

                // 2. 注入测试 VBA 模块到 Target 工作簿
                string vbaCode = @"
Public g_LogResult As String

Sub TestParamWorkbook(targetWb As Workbook)
    On Error Resume Next
    Dim tn As String, fn As String, isTarget As Boolean, isActive As Boolean
    tn = TypeName(targetWb)
    fn = targetWb.FullName
    isTarget = (targetWb.FullName = ThisWorkbook.FullName)
    isActive = (targetWb Is ActiveWorkbook)
    g_LogResult = ""TYPE="" & tn & ""|FN="" & fn & ""|ISTARGET="" & isTarget & ""|ISACTIVE="" & isActive & ""|ERR="" & Err.Number & "" "" & Err.Description
    ThisWorkbook.Sheets(1).Range(""B1"").Value = g_LogResult
End Sub

Sub TestParamObject(targetWb As Object)
    On Error Resume Next
    Dim tn As String, fn As String, isTarget As Boolean, isActive As Boolean
    tn = TypeName(targetWb)
    fn = targetWb.FullName
    isTarget = (targetWb.FullName = ThisWorkbook.FullName)
    isActive = (targetWb Is ActiveWorkbook)
    g_LogResult = ""TYPE="" & tn & ""|FN="" & fn & ""|ISTARGET="" & isTarget & ""|ISACTIVE="" & isActive & ""|ERR="" & Err.Number & "" "" & Err.Description
    ThisWorkbook.Sheets(1).Range(""B2"").Value = g_LogResult
End Sub

Sub TestOptionalWorkbook(Optional targetWb As Workbook = Nothing)
    On Error Resume Next
    Dim tn As String, fn As String, isNothing As Boolean
    isNothing = (targetWb Is Nothing)
    If Not isNothing Then
        tn = TypeName(targetWb)
        fn = targetWb.FullName
    Else
        tn = ""Nothing""
        fn = ""None""
    End If
    g_LogResult = ""ISNOTHING="" & isNothing & ""|TYPE="" & tn & ""|FN="" & fn & ""|ERR="" & Err.Number
    ThisWorkbook.Sheets(1).Range(""B3"").Value = g_LogResult
End Sub

Sub TestNoParam()
    ' 未显式指定工作簿，直接操作 ActiveSheet
    ActiveSheet.Range(""C1"").Value = ""WRITTEN_BY_NOPARAM""
End Sub

Sub TestWrapperActivate(targetWb As Workbook)
    targetWb.Activate
    ThisWorkbook.Sheets(1).Range(""B4"").Value = ""ACTIVE_NOW="" & ActiveWorkbook.FullName
    ' 激活后调用无参宏
    Call TestNoParam
End Sub
";

                dynamic proj = wbTarget.VBProject;
                dynamic modComp = proj.VBComponents.Add(1); // 1 = vbext_ct_StdModule
                modComp.Name = "ModTestBinding";
                modComp.CodeModule.AddFromString(vbaCode);

                // =========================================================================
                // 实验 1: wbOther 处于活动状态，调用 TestParamWorkbook(targetWb As Workbook)
                // =========================================================================
                wbOther.Activate();
                log("\n--- 实验 1: ActiveWorkbook 为 Other 时，调用 TestParamWorkbook(Workbook 对象) ---");
                log("调用前 ActiveWorkbook: " + (string)excel.ActiveWorkbook.FullName);

                string macroAddress1 = "'" + targetName + "'!TestParamWorkbook";
                try
                {
                    excel.Run(macroAddress1, wbTarget);
                    string res1 = (string)wbTarget.Sheets[1].Range("B1").Value;
                    log("调用结果: 成功运行，无 COM 异常");
                    log("宏内实测记录 (Target!B1): " + res1);
                }
                catch (Exception ex)
                {
                    log("调用结果: 发生异常 -> " + ex.Message);
                }

                // =========================================================================
                // 实验 2: wbOther 处于活动状态，调用 TestParamObject(targetWb As Object)
                // =========================================================================
                wbOther.Activate();
                log("\n--- 实验 2: ActiveWorkbook 为 Other 时，调用 TestParamObject(Object 对象) ---");
                log("调用前 ActiveWorkbook: " + (string)excel.ActiveWorkbook.FullName);

                string macroAddress2 = "'" + targetName + "'!TestParamObject";
                try
                {
                    excel.Run(macroAddress2, wbTarget);
                    string res2 = (string)wbTarget.Sheets[1].Range("B2").Value;
                    log("调用结果: 成功运行，无 COM 异常");
                    log("宏内实测记录 (Target!B2): " + res2);
                }
                catch (Exception ex)
                {
                    log("调用结果: 发生异常 -> " + ex.Message);
                }

                // =========================================================================
                // 实验 3: 不传参数调用 TestOptionalWorkbook()
                // =========================================================================
                log("\n--- 实验 3: 不传参调用 TestOptionalWorkbook(Optional) ---");
                string macroAddress3 = "'" + targetName + "'!TestOptionalWorkbook";
                try
                {
                    excel.Run(macroAddress3);
                    string res3 = (string)wbTarget.Sheets[1].Range("B3").Value;
                    log("调用结果: 成功运行，无 COM 异常");
                    log("宏内实测记录 (Target!B3): " + res3);
                }
                catch (Exception ex)
                {
                    log("调用结果: 发生异常 -> " + ex.Message);
                }

                // =========================================================================
                // 实验 4: wbOther 处于活动状态，调用 Target 中的无参宏 TestNoParam()
                // =========================================================================
                wbOther.Activate();
                log("\n--- 实验 4: ActiveWorkbook 为 Other 时，直接调用 Target 中的无参宏 TestNoParam() ---");
                log("调用前 ActiveWorkbook: " + (string)excel.ActiveWorkbook.FullName);

                string macroAddress4 = "'" + targetName + "'!TestNoParam";
                try
                {
                    excel.Run(macroAddress4);
                    log("调用结果: 成功运行");
                    string targetC1 = (string)wbTarget.Sheets[1].Range("C1").Value;
                    string otherC1 = (string)wbOther.Sheets[1].Range("C1").Value;
                    log("Target!C1 内容: " + (targetC1 ?? "<空>"));
                    log("Other!C1 内容 : " + (otherC1 ?? "<空>"));
                    if (otherC1 == "WRITTEN_BY_NOPARAM")
                    {
                        log("【致命实测证实】无参宏若依赖 ActiveSheet/ActiveWorkbook 且未做 Activate，将把数据致命写入非目标的活动工作簿！");
                    }
                }
                catch (Exception ex)
                {
                    log("调用结果: 发生异常 -> " + ex.Message);
                }

                // =========================================================================
                // 实验 5: wbOther 处于活动状态，调用宿主包装器 TestWrapperActivate(wbTarget)
                // =========================================================================
                // 清理 C1
                wbTarget.Sheets[1].Range("C1").Value = null;
                wbOther.Sheets[1].Range("C1").Value = null;
                wbOther.Activate();
                log("\n--- 实验 5: ActiveWorkbook 为 Other 时，调用宿主包装器 TestWrapperActivate(wbTarget) ---");
                log("调用前 ActiveWorkbook: " + (string)excel.ActiveWorkbook.FullName);

                string macroAddress5 = "'" + targetName + "'!TestWrapperActivate";
                try
                {
                    excel.Run(macroAddress5, wbTarget);
                    log("调用结果: 成功运行");
                    string res5 = (string)wbTarget.Sheets[1].Range("B4").Value;
                    log("宏内记录 (Target!B4): " + res5);
                    string targetC1After = (string)wbTarget.Sheets[1].Range("C1").Value;
                    string otherC1After = (string)wbOther.Sheets[1].Range("C1").Value;
                    log("调用后 ActiveWorkbook: " + (string)excel.ActiveWorkbook.FullName);
                    log("Target!C1 内容: " + (targetC1After ?? "<空>"));
                    log("Other!C1 内容 : " + (otherC1After ?? "<空>"));
                    if (targetC1After == "WRITTEN_BY_NOPARAM" && string.IsNullOrEmpty(otherC1After))
                    {
                        log("【结论证实】宿主包装器执行 targetWb.Activate 成功将焦点收归目标工作簿，使后续无参子过程安全正确写入目标工作簿，未污染非目标簿！");
                    }
                }
                catch (Exception ex)
                {
                    log("调用结果: 发生异常 -> " + ex.Message);
                }

                File.WriteAllText(logPath, report.ToString(), Encoding.UTF8);
                log("\n全部测试已完成，测试日志已保存至: " + logPath);
            }
            catch (Exception ex)
            {
                log("严重未捕获异常: " + ex.ToString());
            }
            finally
            {
                if (wbTarget != null) { try { wbTarget.Close(false); } catch { } }
                if (wbOther != null) { try { wbOther.Close(false); } catch { } }
                if (excel != null) { try { excel.Quit(); } catch { } }
            }
        }
    }
}
