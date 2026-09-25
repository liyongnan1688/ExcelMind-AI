using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class WorkbookReadback
    {
        public string targetWorkbookName { get; set; }
        public string targetWorkbookFullName { get; set; }
        public string targetSheetName { get; set; }
        public string usedRangeAddress { get; set; }
        public int rowCount { get; set; }
        public int columnCount { get; set; }
        public string startCell { get; set; }
        public string endCell { get; set; }
        public List<string> sampleValues { get; set; }
        public bool hasFormulas { get; set; }
        public bool hasBorders { get; set; }
        public bool hasInteriorColor { get; set; }
        public int sheetCount { get; set; }
        public bool targetVerified { get; set; }
        public bool otherWorkbooksAffected { get; set; }
        public string affectedWorkbooksWarning { get; set; }
    }

    public class VbaExecutionResult
    {
        public bool success { get; set; }
        public string summary { get; set; }
        public string error { get; set; }
        public string originalVbaCode { get; set; }
        public string executedVbaCode { get; set; }
        public List<string> transformSteps { get; set; }
        public double elapsedMs { get; set; }
        public WorkbookReadback readback { get; set; }
    }

    public class VbaRunner
    {
        public static VbaExecutionResult RunVbaCode(dynamic app, dynamic targetWorkbook, string vbaCode)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var transformSteps = new List<string>();

            if (targetWorkbook == null)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：未指定有效的目标工作簿",
                    error = "目标工作簿对象为空或已关闭",
                    originalVbaCode = vbaCode ?? "",
                    executedVbaCode = "",
                    transformSteps = transformSteps,
                    elapsedMs = 0
                };
            }

            if (string.IsNullOrEmpty(vbaCode))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：传入的代码为空",
                    error = "未检测到有效 VBA 代码",
                    originalVbaCode = "",
                    executedVbaCode = "",
                    transformSteps = transformSteps,
                    elapsedMs = 0
                };
            }

            // 1. 结构完整性校验：必须包含过程定义与 End Sub
            var matchSub = Regex.Match(vbaCode, @"(?:Public\s+|Private\s+)?Sub\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*\((.*?)\)", RegexOptions.IgnoreCase);
            bool hasEndSub = Regex.IsMatch(vbaCode, @"End\s+Sub", RegexOptions.IgnoreCase);

            if (!matchSub.Success || !hasEndSub)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：VBA 代码不完整 (缺少 Sub 或 End Sub)",
                    error = "模型输出的代码结构不完整，已在注入前安全拦截，避免导致 Excel 编译错误。",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            string subName = matchSub.Groups[1].Value;
            string paramsList = matchSub.Groups[2].Value.Trim();

            // 2. 目标对象显式绑定约定校验：
            // 拒绝使用 Regex.Replace 全局重写 ThisWorkbook / ActiveWorkbook
            // 依据受控约定：宏过程必须声明接收目标工作簿参数 (如 targetWb As Workbook 或 targetWb As Object)
            // 若代码未声明参数，拒绝自动运行，防止无目标或漂移到其他前台活动窗口
            if (string.IsNullOrEmpty(paramsList))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "安全拦截：宏过程未声明目标工作簿参数",
                    error = "安全约定拦截：模型生成的宏未接收目标工作簿参数 (需声明 Sub LeeTaskEntry(targetWb As Workbook))，存在误改其他前台窗口的风险，已安全拒绝自动运行以保护工作簿。",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            string targetWbName = "";
            string targetWbFullName = "";
            int preSheetCount = 0;
            try
            {
                targetWbName = (string)targetWorkbook.Name;
                targetWbFullName = (string)targetWorkbook.FullName;
                preSheetCount = (int)targetWorkbook.Sheets.Count;
            }
            catch (Exception ex)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：无法访问目标工作簿属性",
                    error = "目标工作簿状态异常: " + ex.Message,
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 3. 仅对过程名标准化为英文 LeeTaskEntry，保持参数签名完整传递，绝不篡改代码正文中的任何文本
            string entrySubName = "LeeTaskEntry";
            string finalCode = vbaCode;
            if (!string.Equals(subName, entrySubName, StringComparison.OrdinalIgnoreCase))
            {
                finalCode = Regex.Replace(
                    finalCode,
                    @"(?:Public\s+|Private\s+)?Sub\s+" + Regex.Escape(subName) + @"\s*\(",
                    "Sub " + entrySubName + "(",
                    RegexOptions.IgnoreCase
                );
                transformSteps.Add("规整过程入口: " + subName + " -> " + entrySubName + " (保持目标工作簿参数: " + paramsList + ")");
            }

            // 4. 执行前：记录当前 Excel 实例中所有【非目标工作簿】的初始状态快照
            var otherWbSnapshots = new List<Tuple<string, int, string>>();
            try
            {
                foreach (dynamic wb in app.Workbooks)
                {
                    try
                    {
                        string oFull = (string)wb.FullName;
                        if (!string.Equals(oFull, targetWbFullName, StringComparison.OrdinalIgnoreCase))
                        {
                            string oName = (string)wb.Name;
                            int sCount = (int)wb.Sheets.Count;
                            string uAddr = "";
                            try { uAddr = ((string)wb.ActiveSheet.UsedRange.Address) ?? ""; } catch { }
                            otherWbSnapshots.Add(Tuple.Create(oName, sCount, uAddr));
                        }
                    }
                    catch { }
                }
            }
            catch { }

            dynamic vbComp = null;
            dynamic vbProj = null;
            bool screenUpdated = false;

            try
            {
                // 5. 检查 VBProject 访问权限
                try
                {
                    vbProj = targetWorkbook.VBProject;
                }
                catch (COMException comEx)
                {
                    if ((uint)comEx.ErrorCode == 0x800A03EC || comEx.Message.Contains("1004"))
                    {
                        return new VbaExecutionResult
                        {
                            success = false,
                            summary = "执行中断：Excel 未开启对 VBA 工程对象模型的访问信任",
                            error = "错误 1004：请在 Excel“文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置”中勾选【信任对 VBA 工程对象模型的访问】。",
                            originalVbaCode = vbaCode,
                            executedVbaCode = finalCode,
                            transformSteps = transformSteps,
                            elapsedMs = sw.ElapsedMilliseconds
                        };
                    }
                    throw;
                }

                // 6. 动态创建标准临时模块并注入宏
                string moduleName = "LeeMod_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                vbComp = vbProj.VBComponents.Add(1);
                vbComp.Name = moduleName;
                vbComp.CodeModule.AddFromString(finalCode);

                // 7. 挂起屏幕刷新与系统弹窗
                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    screenUpdated = true;
                }
                catch { }

                // 8. 调用标准化受控宏，将真实 targetWorkbook COM 对象通过参数安全传入
                string macroAddress = "'" + targetWbName + "'!" + entrySubName;
                app.Run(macroAddress, targetWorkbook);

                sw.Stop();

                // 9. 执行后检查：核验非目标工作簿是否受到意外影响
                bool otherAffected = false;
                string otherWarning = null;
                foreach (var snap in otherWbSnapshots)
                {
                    try
                    {
                        dynamic oWb = app.Workbooks[snap.Item1];
                        int postCount = (int)oWb.Sheets.Count;
                        string postAddr = "";
                        try { postAddr = ((string)oWb.ActiveSheet.UsedRange.Address) ?? ""; } catch { }

                        if (postCount != snap.Item2 || postAddr != snap.Item3)
                        {
                            otherAffected = true;
                            otherWarning = "检测到非目标工作簿 [" + snap.Item1 + "] 发生附带变更！";
                            break;
                        }
                    }
                    catch { }
                }

                // 10. 目标工作簿写后读回与核验
                var readback = PerformReadback(targetWorkbook, targetWbName, targetWbFullName, preSheetCount);
                readback.otherWorkbooksAffected = otherAffected;
                readback.affectedWorkbooksWarning = otherWarning;

                return new VbaExecutionResult
                {
                    success = !otherAffected,
                    summary = otherAffected ? ("执行告警：" + otherWarning) : ("宏已运行（耗时 " + sw.ElapsedMilliseconds + " ms）"),
                    error = otherWarning,
                    originalVbaCode = vbaCode,
                    executedVbaCode = finalCode,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds,
                    readback = readback
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errDetail = ex.Message;
                if (ex.InnerException != null) errDetail += " (" + ex.InnerException.Message + ")";

                if (vbProj != null)
                {
                    try
                    {
                        vbProj.VBE.MainWindow.Visible = false;
                    }
                    catch { }
                }

                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行失败：VBA 运行期抛出异常",
                    error = errDetail,
                    originalVbaCode = vbaCode,
                    executedVbaCode = finalCode,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }
            finally
            {
                // 11. 瞬时清理临时模块，保持工作簿纯净无宏
                if (vbComp != null && vbProj != null)
                {
                    try
                    {
                        vbProj.VBComponents.Remove(vbComp);
                    }
                    catch (Exception rmEx)
                    {
                        System.Diagnostics.Debug.WriteLine("Remove temp module warning: " + rmEx.Message);
                    }
                }

                if (screenUpdated)
                {
                    try
                    {
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                    catch { }
                }
            }
        }

        private static WorkbookReadback PerformReadback(dynamic targetWb, string expectedName, string expectedFullName, int preSheetCount)
        {
            var rb = new WorkbookReadback
            {
                targetWorkbookName = expectedName,
                targetWorkbookFullName = expectedFullName,
                targetVerified = false,
                otherWorkbooksAffected = false,
                affectedWorkbooksWarning = null,
                sampleValues = new List<string>(),
                rowCount = 0,
                columnCount = 0,
                usedRangeAddress = "",
                startCell = "",
                endCell = "",
                hasFormulas = false,
                hasBorders = false,
                hasInteriorColor = false,
                sheetCount = 0
            };

            try
            {
                // 1. 验证目标工作簿身份
                string currentName = (string)targetWb.Name;
                rb.targetVerified = string.Equals(currentName, expectedName, StringComparison.OrdinalIgnoreCase);
                rb.sheetCount = (int)targetWb.Sheets.Count;

                // 2. 获取当前活动工作表或新建的工作表
                dynamic sheet = null;
                try
                {
                    sheet = targetWb.ActiveSheet;
                }
                catch { }

                if (sheet == null)
                {
                    try
                    {
                        sheet = targetWb.Sheets[1];
                    }
                    catch { }
                }

                if (sheet != null)
                {
                    rb.targetSheetName = (string)sheet.Name;

                    dynamic usedRange = null;
                    try
                    {
                        usedRange = sheet.UsedRange;
                    }
                    catch { }

                    if (usedRange != null)
                    {
                        string rawAddr = (string)usedRange.Address;
                        rb.usedRangeAddress = rawAddr != null ? rawAddr.Replace("$", "") : "";
                        rb.rowCount = (int)usedRange.Rows.Count;
                        rb.columnCount = (int)usedRange.Columns.Count;

                        if (!string.IsNullOrEmpty(rb.usedRangeAddress))
                        {
                            var parts = rb.usedRangeAddress.Split(':');
                            rb.startCell = parts[0];
                            rb.endCell = parts.Length > 1 ? parts[1] : parts[0];
                        }

                        // 3. 抽样单元格文本
                        int sampleMaxRows = Math.Min(rb.rowCount, 9);
                        int sampleMaxCols = Math.Min(rb.columnCount, 9);

                        for (int r = 1; r <= sampleMaxRows && rb.sampleValues.Count < 12; r++)
                        {
                            for (int c = 1; c <= sampleMaxCols && rb.sampleValues.Count < 12; c++)
                            {
                                try
                                {
                                    dynamic cell = usedRange.Cells[r, c];
                                    object val = cell.Value2;
                                    if (val != null)
                                    {
                                        string text = val.ToString().Trim();
                                        if (!string.IsNullOrEmpty(text) && !rb.sampleValues.Contains(text))
                                        {
                                            rb.sampleValues.Add(text);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }

                        // 4. 检查公式
                        try
                        {
                            object hf = usedRange.HasFormula;
                            if (hf != null && !(hf is DBNull) && hf is bool && (bool)hf)
                            {
                                rb.hasFormulas = true;
                            }
                            else
                            {
                                for (int r = 1; r <= Math.Min(rb.rowCount, 5) && !rb.hasFormulas; r++)
                                {
                                    for (int c = 1; c <= Math.Min(rb.columnCount, 5); c++)
                                    {
                                        try
                                        {
                                            dynamic cell = usedRange.Cells[r, c];
                                            if ((bool)cell.HasFormula)
                                            {
                                                rb.hasFormulas = true;
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        catch { }

                        // 5. 检查边框 (xlNone = -4142)
                        try
                        {
                            dynamic borders = usedRange.Borders;
                            object lineStyle = borders.LineStyle;
                            if (lineStyle != null && Convert.ToInt32(lineStyle) != -4142)
                            {
                                rb.hasBorders = true;
                            }
                        }
                        catch { }

                        // 6. 检查背景填充色 (xlNone = -4142)
                        try
                        {
                            dynamic interior = usedRange.Interior;
                            object colorIndex = interior.ColorIndex;
                            if (colorIndex != null && Convert.ToInt32(colorIndex) != -4142)
                            {
                                rb.hasInteriorColor = true;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("PerformReadback warning: " + ex.Message);
            }

            return rb;
        }
    }
}
