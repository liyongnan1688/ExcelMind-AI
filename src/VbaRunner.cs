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

            // 1. 严格校验代码是否包含完整的 Sub 与 End Sub
            if (!Regex.IsMatch(vbaCode, @"(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(vbaCode, @"End\s+Sub", RegexOptions.IgnoreCase))
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

            // 2. 规整过程入口为标准化英文 LeeTaskEntry，彻底根除中文过程名引发的 COM 0x800A9C68 异常
            string entrySubName = "LeeTaskEntry";
            string finalCode = vbaCode;
            var matchSub = Regex.Match(finalCode, @"(?:Public\s+|Private\s+)?Sub\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*\(", RegexOptions.IgnoreCase);
            if (matchSub.Success)
            {
                string originalSubName = matchSub.Groups[1].Value;
                finalCode = Regex.Replace(
                    finalCode,
                    @"(?:Public\s+|Private\s+)?Sub\s+" + Regex.Escape(originalSubName) + @"\s*\(",
                    "Sub " + entrySubName + "(",
                    RegexOptions.IgnoreCase
                );
                transformSteps.Add("规整过程入口: " + originalSubName + " -> " + entrySubName);
            }

            // 3. 显式目标工作簿绑定：
            // 将代码中可能存在的 ThisWorkbook 或 ActiveWorkbook 显式绑定为 Application.Workbooks("<targetWbName>")
            // 防止执行期间因用户切换前台活动窗口造成对非目标工作簿的串改
            bool replacedThis = Regex.IsMatch(finalCode, @"\bThisWorkbook\b", RegexOptions.IgnoreCase);
            bool replacedActive = Regex.IsMatch(finalCode, @"\bActiveWorkbook\b", RegexOptions.IgnoreCase);

            if (replacedThis)
            {
                finalCode = Regex.Replace(
                    finalCode,
                    @"\bThisWorkbook\b",
                    "Application.Workbooks(\"" + targetWbName.Replace("\"", "\"\"") + "\")",
                    RegexOptions.IgnoreCase
                );
                transformSteps.Add("规整 ThisWorkbook -> Application.Workbooks(\"" + targetWbName + "\")");
            }

            if (replacedActive)
            {
                finalCode = Regex.Replace(
                    finalCode,
                    @"\bActiveWorkbook\b",
                    "Application.Workbooks(\"" + targetWbName.Replace("\"", "\"\"") + "\")",
                    RegexOptions.IgnoreCase
                );
                transformSteps.Add("规整 ActiveWorkbook -> Application.Workbooks(\"" + targetWbName + "\")");
            }

            dynamic vbComp = null;
            dynamic vbProj = null;
            bool screenUpdated = false;

            try
            {
                // 4. 检查 VBProject 访问权限
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

                // 5. 动态创建标准临时模块 (1 = vbext_ct_StdModule)
                string moduleName = "LeeMod_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                vbComp = vbProj.VBComponents.Add(1);
                vbComp.Name = moduleName;
                vbComp.CodeModule.AddFromString(finalCode);

                // 6. 执行前置激活目标工作簿并挂起屏幕刷新与告警
                try
                {
                    targetWorkbook.Activate();
                }
                catch { }

                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    screenUpdated = true;
                }
                catch { }

                // 7. 调用标准化宏
                string macroAddress = "'" + targetWbName + "'!" + entrySubName;
                app.Run(macroAddress);

                sw.Stop();

                // 8. 执行后身份核验与实际状态写后读回 (Readback)
                var readback = PerformReadback(targetWorkbook, targetWbName, targetWbFullName, preSheetCount);

                return new VbaExecutionResult
                {
                    success = true,
                    summary = "宏已运行（耗时 " + sw.ElapsedMilliseconds + " ms）",
                    error = null,
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

                // 若触发了编译错误导致 VBE 弹窗，强制将 VBE 窗口隐藏
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
                // 9. 瞬时清理临时模块，保持工作簿纯净无宏
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

                // 恢复屏幕刷新
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
                string currentFullName = (string)targetWb.FullName;
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

                        // 3. 抽样单元格文本（优先提取前 5 行 5 列及对角线采样）
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
