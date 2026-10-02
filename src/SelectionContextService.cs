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
        public long capturedAt { get; set; }
        public string attachmentId { get; set; }
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
        public static SelectionContextResult GetSelectionContext(dynamic app, int requestedSampleRows = DefaultMaxSampleRows, int requestedSampleCols = DefaultMaxSampleCols, string attachmentId = null)
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

                // 校验选中对象是否为 Range
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

                return ExtractContextFromRange(wb, sheet, selection, requestedSampleRows, requestedSampleCols, attachmentId);
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

        /// <summary>
        /// 严格根据原工作簿身份、原工作表名和原地址只读刷新指定区域
        /// 绝不猜测同名表、绝不回退到活动表、绝不激活工作簿或改变用户当前选区
        /// </summary>
        public static SelectionContextResult GetSpecificRangeContext(
            dynamic app,
            string targetWbFullName,
            string targetWbName,
            string targetSheetName,
            string targetAddress,
            int requestedSampleRows = DefaultMaxSampleRows,
            int requestedSampleCols = DefaultMaxSampleCols,
            string attachmentId = null)
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
                // 1. 定位原工作簿
                dynamic targetWb = NativeBridge.FindTargetWorkbook(app, targetWbFullName, targetWbName);
                if (targetWb == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "原工作簿已关闭或未找到: " + (string.IsNullOrEmpty(targetWbName) ? targetWbFullName : targetWbName),
                        errorType = "workbook_not_found"
                    };
                }

                // 2. 严格按名称定位原工作表（不猜测同名表，不回退到 ActiveSheet）
                dynamic targetSheet = null;
                try
                {
                    foreach (dynamic ws in targetWb.Worksheets)
                    {
                        string sName = "";
                        try { sName = ws.Name; } catch { }
                        if (string.Equals(sName, targetSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetSheet = ws;
                            break;
                        }
                    }
                }
                catch (COMException comEx)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "定位原工作表失败，Excel 正处于忙碌或编辑状态 (COM: 0x" + comEx.ErrorCode.ToString("X") + ")",
                        errorType = "busy"
                    };
                }

                if (targetSheet == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "原工作表未找到或已被重命名/删除: " + targetSheetName,
                        errorType = "sheet_not_found"
                    };
                }

                // 3. 读取原地址区域
                dynamic range = null;
                try
                {
                    range = targetSheet.Range[targetAddress];
                }
                catch (Exception rangeEx)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "无法定位原区域地址 [" + targetAddress + "]: " + rangeEx.Message,
                        errorType = "range_not_found"
                    };
                }

                if (range == null)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "无法获取原区域: " + targetAddress,
                        errorType = "range_not_found"
                    };
                }

                // 校验 Areas 计数
                dynamic areas = null;
                int areaCount = 1;
                try
                {
                    areas = range.Areas;
                    if (areas != null) areaCount = (int)areas.Count;
                }
                catch { }

                if (areaCount > 1)
                {
                    return new SelectionContextResult
                    {
                        ok = false,
                        error = "原区域包含不连续子区域，暂不支持刷新",
                        errorType = "multi_area"
                    };
                }

                // 4. 纯只读提取，绝不激活工作簿或工作表，绝不改变当前用户选区
                return ExtractContextFromRange(targetWb, targetSheet, range, requestedSampleRows, requestedSampleCols, attachmentId);
            }
            catch (Exception ex)
            {
                return new SelectionContextResult
                {
                    ok = false,
                    error = "刷新原区域异常: " + ex.Message,
                    errorType = "exception"
                };
            }
        }

        private static SelectionContextResult ExtractContextFromRange(
            dynamic wb,
            dynamic sheet,
            dynamic range,
            int requestedSampleRows,
            int requestedSampleCols,
            string attachmentId)
        {
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

            try { rawValues = sampleRange.Value2; } catch { }
            try { rawFormulas = sampleRange.Formula; } catch { }

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
            catch { formulaStatus = "unknown"; }

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
            catch { mergeStatus = "unknown"; }

            // 5. 格式化解析样本数据（保留类型，避免字符串简单归一化）
            var candidateHeaders = new List<string>();
            var sampleRows = new List<List<CellSampleItem>>();

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
                    if (valArray != null) rawVal = valArray[r, c];
                    else if (fetchRows == 1 && fetchCols == 1) rawVal = rawValues;

                    object rawFormula = null;
                    if (formulaArray != null) rawFormula = formulaArray[r, c];
                    else if (fetchRows == 1 && fetchCols == 1) rawFormula = rawFormulas;

                    var item = ParseCellItem(actualRow, actualCol, rawVal, rawFormula);
                    rowItems.Add(item);

                    if (r == 1)
                    {
                        string headerText = string.IsNullOrEmpty(item.displayText) ? ("列" + c) : item.displayText;
                        candidateHeaders.Add(headerText);
                    }
                }
                sampleRows.Add(rowItems);
            }

            long nowMs = (long)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalMilliseconds;
            string finalAttId = string.IsNullOrEmpty(attachmentId) ? ("att_" + Guid.NewGuid().ToString("N").Substring(0, 12)) : attachmentId;

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
                ),
                capturedAt = nowMs,
                attachmentId = finalAttId
            };

            return new SelectionContextResult
            {
                ok = true,
                data = data
            };
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
