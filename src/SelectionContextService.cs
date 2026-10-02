using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace LeeExcel
{
    public class CellSampleItem
    {
        public int row { get; set; }
        public int col { get; set; }
        public string address { get; set; }
        public object value { get; set; }
        public string displayText { get; set; }
        public string formula { get; set; }
        public string valueType { get; set; } // "string", "number", "boolean", "empty", "error"
        public bool isTextTruncated { get; set; }
    }

    public class SelectionContextData
    {
        public string workbookName { get; set; }
        public string workbookFullName { get; set; }
        public string sheetName { get; set; }
        public string address { get; set; }
        public int totalRows { get; set; }
        public int totalColumns { get; set; }
        public int startRow { get; set; }
        public int startColumn { get; set; }
        public int endRow { get; set; }
        public int endColumn { get; set; }
        public bool isSingleArea { get; set; }
        public int sampleRowCount { get; set; }
        public int sampleColumnCount { get; set; }
        public string sampleAddress { get; set; }
        public List<string> candidateHeaders { get; set; }
        public List<List<CellSampleItem>> sampleRows { get; set; }
        public string formulaStatus { get; set; } // "has_formula", "no_formula", "sample_mixed", "unknown"
        public string mergeStatus { get; set; } // "has_merged", "no_merged", "sample_mixed", "unknown"
        public string visibilityStatus { get; set; } // "sample_scanned_only", "unknown"
        public bool isRowTruncated { get; set; }
        public bool isColumnTruncated { get; set; }
        public int maxTextLengthLimit { get; set; }
        public string unscannedNotes { get; set; }
    }

    public class SelectionContextResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string errorType { get; set; } // "no_workbook", "not_range", "multi_area", "busy", "exception"
        public SelectionContextData data { get; set; }
    }

    public class SelectionContextService
    {
        public const int DefaultMaxSampleRows = 5;
        public const int DefaultMaxSampleCols = 15;
        public const int MaxTextCharLimit = 100;

        /// <summary>
        /// 批量、安全、纯只读地获取 Excel 当前选区元数据与限定样本子区域
        /// 绝不写入、格式化、修改选区，绝不整列整表加载，绝不调用大模型或 API Key
        /// </summary>
        public static SelectionContextResult GetSelectionContext(dynamic app, int requestedSampleRows = DefaultMaxSampleRows, int requestedSampleCols = DefaultMaxSampleCols)
        {
            if (app == null)
            {
                return new SelectionContextResult
                {
                    ok = false,
                    error = "Excel 宿主实例未就绪",
                    errorType = "no_workbook"
                };
            }

            try
            {
                dynamic wb = null;
                try
                {
                    wb = app.ActiveWorkbook;
                }
                catch (COMException comEx)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "Excel 当前正处于编辑状态或忙碌中，请按 Enter 或 Esc 退出单元格编辑后再试 (COM: 0x" + comEx.ErrorCode.ToString("X") + ")",
                        errorType = "busy"
                    };
                }

                if (wb == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "未检测到已打开的活动工作簿",
                        errorType = "no_workbook"
                    };
                }

                dynamic sheet = null;
                try
                {
                    sheet = wb.ActiveSheet;
                }
                catch { }

                if (sheet == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "未检测到活动工作表",
                        errorType = "no_workbook"
                    };
                }

                dynamic selection = null;
                try
                {
                    selection = app.Selection;
                }
                catch (COMException comEx)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "获取选区失败，Excel 正处于忙碌或单元格编辑状态 (COM: 0x" + comEx.ErrorCode.ToString("X") + ")",
                        errorType = "busy"
                    };
                }

                if (selection == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "当前没有活动选区",
                        errorType = "not_range"
                    };
                }

                // 校验选中对象是否为 Range（排除选中文档、图表、图形、形状等非单元格对象）
                string typeName = "";
                try
                {
                    // 在 COM 互操作中可通过 Microsoft.VisualBasic.Information.TypeName，或通过反射/属性判断
                    typeName = selection.GetType().Name;
                }
                catch { }

                // 检查 Areas 属性
                dynamic areas = null;
                int areaCount = 1;
                try
                {
                    areas = selection.Areas;
                    if (areas != null)
                    {
                        areaCount = (int)areas.Count;
                    }
                }
                catch
                {
                    // 若无 Areas 属性，说明不是 Range
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "当前选中的对象不是单元格区域（可能为图表、形状或图片），暂不支持附加",
                        errorType = "not_range"
                    };
                }

                if (areaCount > 1)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "当前选区包含 " + areaCount + " 个不连续子区域。第一版暂仅支持单一连续矩形选区，请按住鼠标框选单一矩形区域。",
                        errorType = "multi_area"
                    };
                }

                dynamic range = selection;

                // 1. 获取全局元信息（不读取全表单元格）
                string wbName = "";
                string wbFullName = "";
                string sheetName = "";
                string address = "";
                int totalRows = 0;
                int totalCols = 0;
                int startRow = 1;
                int startCol = 1;

                try { wbName = wb.Name ?? ""; } catch { }
                try { wbFullName = wb.FullName ?? ""; } catch { }
                try { sheetName = sheet.Name ?? ""; } catch { }
                try { address = range.Address ?? ""; } catch { }
                try { totalRows = (int)range.Rows.Count; } catch { }
                try { totalCols = (int)range.Columns.Count; } catch { }
                try { startRow = (int)range.Row; } catch { }
                try { startCol = (int)range.Column; } catch { }

                int endRow = startRow + Math.Max(0, totalRows - 1);
                int endCol = startCol + Math.Max(0, totalCols - 1);

                // 2. 确定样本读取边界（严格限制采样上限，杜绝整行/整列读取造成假死）
                int maxAllowedRows = Math.Max(1, Math.Min(requestedSampleRows, DefaultMaxSampleRows));
                int maxAllowedCols = Math.Max(1, Math.Min(requestedSampleCols, DefaultMaxSampleCols));

                int fetchRows = Math.Min(totalRows, maxAllowedRows);
                int fetchCols = Math.Min(totalCols, maxAllowedCols);

                bool isRowTruncated = totalRows > fetchRows;
                bool isColTruncated = totalCols > fetchCols;

                // 构造样本子区域 Range（只读前 fetchRows 行 × fetchCols 列）
                dynamic sampleRange = null;
                try
                {
                    dynamic startCell = sheet.Cells[startRow, startCol];
                    dynamic endCell = sheet.Cells[startRow + fetchRows - 1, startCol + fetchCols - 1];
                    sampleRange = sheet.Range[startCell, endCell];
                }
                catch (Exception rangeEx)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "构造样本子区域失败: " + rangeEx.Message,
                        errorType = "exception"
                    };
                }

                string sampleAddress = "";
                try { sampleAddress = sampleRange.Address ?? ""; } catch { }

                // 3. 批量读取样本的值与公式（仅在子区域调用一次，严禁逐格 COM 遍历）
                object rawValues = null;
                object rawFormulas = null;

                try
                {
                    rawValues = sampleRange.Value2;
                }
                catch (Exception valEx)
                {
                    System.Diagnostics.Debug.WriteLine("Value2 read error: " + valEx.Message);
                }

                try
                {
                    rawFormulas = sampleRange.Formula;
                }
                catch (Exception fEx)
                {
                    System.Diagnostics.Debug.WriteLine("Formula read error: " + fEx.Message);
                }

                // 4. 读取样本区域的可观察合并与公式状态（局部状态，不宣称为整表全量）
                string formulaStatus = "no_formula";
                try
                {
                    object hasFormulaObj = sampleRange.HasFormula;
                    if (hasFormulaObj != null && !(hasFormulaObj is DBNull))
                    {
                        if (hasFormulaObj is bool && (bool)hasFormulaObj) formulaStatus = "has_formula";
                        else if (hasFormulaObj is bool && !(bool)hasFormulaObj) formulaStatus = "no_formula";
                    }
                    else
                    {
                        formulaStatus = "sample_mixed";
                    }
                }
                catch
                {
                    formulaStatus = "unknown";
                }

                string mergeStatus = "no_merged";
                try
                {
                    object mergeObj = sampleRange.MergeCells;
                    if (mergeObj != null && !(mergeObj is DBNull))
                    {
                        if (mergeObj is bool && (bool)mergeObj) mergeStatus = "has_merged";
                        else if (mergeObj is bool && !(bool)mergeObj) mergeStatus = "no_merged";
                    }
                    else
                    {
                        mergeStatus = "sample_mixed";
                    }
                }
                catch
                {
                    mergeStatus = "unknown";
                }

                // 5. 格式化解析样本数据（保留类型，避免字符串简单归一化）
                var candidateHeaders = new List<string>();
                var sampleRows = new List<List<CellSampleItem>>();

                // 区分单单元格标量 vs 二维数组
                object[,] valArray = rawValues as object[,];
                object[,] formulaArray = rawFormulas as object[,];

                for (int r = 1; r <= fetchRows; r++)
                {
                    var rowItems = new List<CellSampleItem>();
                    for (int c = 1; c <= fetchCols; c++)
                    {
                        int actualRow = startRow + r - 1;
                        int actualCol = startCol + c - 1;

                        object rawVal = null;
                        if (valArray != null)
                        {
                            rawVal = valArray[r, c];
                        }
                        else if (fetchRows == 1 && fetchCols == 1)
                        {
                            rawVal = rawValues;
                        }

                        object rawFormula = null;
                        if (formulaArray != null)
                        {
                            rawFormula = formulaArray[r, c];
                        }
                        else if (fetchRows == 1 && fetchCols == 1)
                        {
                            rawFormula = rawFormulas;
                        }

                        var item = ParseCellItem(actualRow, actualCol, rawVal, rawFormula);
                        rowItems.Add(item);

                        // 收集第一行作为“候选表头”
                        if (r == 1)
                        {
                            string headerText = string.IsNullOrEmpty(item.displayText) ? ("列" + c) : item.displayText;
                            candidateHeaders.Add(headerText);
                        }
                    }
                    sampleRows.Add(rowItems);
                }

                // 6. 构造完整上下文 DTO
                var data = new SelectionContextData
                {
                    workbookName = wbName,
                    workbookFullName = wbFullName,
                    sheetName = sheetName,
                    address = address,
                    totalRows = totalRows,
                    totalColumns = totalCols,
                    startRow = startRow,
                    startColumn = startCol,
                    endRow = endRow,
                    endColumn = endCol,
                    isSingleArea = true,
                    sampleRowCount = fetchRows,
                    sampleColumnCount = fetchCols,
                    sampleAddress = sampleAddress,
                    candidateHeaders = candidateHeaders,
                    sampleRows = sampleRows,
                    formulaStatus = formulaStatus,
                    mergeStatus = mergeStatus,
                    visibilityStatus = "sample_scanned_only",
                    isRowTruncated = isRowTruncated,
                    isColumnTruncated = isColTruncated,
                    maxTextLengthLimit = MaxTextCharLimit,
                    unscannedNotes = string.Format(
                        "选区共 {0} 行 × {1} 列。本次仅安全抽样前 {2} 行 × 前 {3} 列。未扫描其余单元格内容，未扫描全表筛选/隐藏行状态。",
                        totalRows, totalCols, fetchRows, fetchCols
                    )
                };

                return new SelectionContextResult
                {
                    ok = true,
                    data = data
                };
            }
            catch (Exception ex)
            {
                return new SelectionContextResult
                {
                    ok = false,
                    error = "读取选区上下文异常: " + ex.Message,
                    errorType = "exception"
                };
            }
        }

        private static CellSampleItem ParseCellItem(int row, int col, object rawVal, object rawFormula)
        {
            string colName = GetColumnLetter(col);
            string addr = "$" + colName + "$" + row;

            var item = new CellSampleItem
            {
                row = row,
                col = col,
                address = addr,
                formula = null,
                isTextTruncated = false
            };

            // 公式提取
            if (rawFormula != null && !(rawFormula is DBNull))
            {
                string fStr = rawFormula.ToString();
                if (fStr.StartsWith("="))
                {
                    item.formula = fStr;
                }
            }

            // 值与类型解析（保留数值、布尔、空、错误区分）
            if (rawVal == null || rawVal is DBNull)
            {
                item.valueType = "empty";
                item.value = null;
                item.displayText = "";
            }
            else if (rawVal is bool)
            {
                item.valueType = "boolean";
                item.value = (bool)rawVal;
                item.displayText = ((bool)rawVal) ? "TRUE" : "FALSE";
            }
            else if (rawVal is double || rawVal is float || rawVal is decimal || rawVal is int || rawVal is long)
            {
                item.valueType = "number";
                item.value = rawVal;
                item.displayText = rawVal.ToString();
            }
            else if (rawVal is string)
            {
                item.valueType = "string";
                string s = (string)rawVal;
                if (s.Length > MaxTextCharLimit)
                {
                    item.value = s.Substring(0, MaxTextCharLimit);
                    item.displayText = s.Substring(0, MaxTextCharLimit) + "...[截断]";
                    item.isTextTruncated = true;
                }
                else
                {
                    item.value = s;
                    item.displayText = s;
                }
            }
            else
            {
                // 其他未知或错误类型
                item.valueType = "error";
                item.value = rawVal.ToString();
                item.displayText = rawVal.ToString();
            }

            return item;
        }

        private static string GetColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }
}
