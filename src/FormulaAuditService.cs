using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class CellDependencyNode
    {
        public string address { get; set; }
        public string sheetName { get; set; }
        public string formula { get; set; }
        public string value { get; set; }
        public bool hasError { get; set; }
        public string errorType { get; set; }
        public bool isBlank { get; set; }
        public List<CellDependencyNode> precedents { get; set; }

        public CellDependencyNode()
        {
            precedents = new List<CellDependencyNode>();
        }
    }

    public class FormulaAuditResult
    {
        public string workbookName { get; set; }
        public string sheetName { get; set; }
        public string cellAddress { get; set; }
        public bool hasFormula { get; set; }
        public string formula { get; set; }
        public string evaluatedValue { get; set; }
        public string formattedText { get; set; }
        public bool hasError { get; set; }
        public string errorType { get; set; }
        public bool isCircular { get; set; }
        public List<string> functionsUsed { get; set; }
        public List<CellDependencyNode> directPrecedents { get; set; }
        public List<string> directDependents { get; set; }
        public string auditSummary { get; set; }
        public string explanationPrompt { get; set; }

        public FormulaAuditResult()
        {
            functionsUsed = new List<string>();
            directPrecedents = new List<CellDependencyNode>();
            directDependents = new List<string>();
        }
    }

    public class SheetOverviewDto
    {
        public string name { get; set; }
        public string visibility { get; set; }
        public string usedRange { get; set; }
        public int rowCount { get; set; }
        public int columnCount { get; set; }
        public int formulaCellCount { get; set; }
        public int tableCount { get; set; }
        public int chartCount { get; set; }
        public List<List<string>> headerPreview { get; set; }

        public SheetOverviewDto()
        {
            headerPreview = new List<List<string>>();
        }
    }

    public class WorkbookOverviewResult
    {
        public string workbookName { get; set; }
        public string fullName { get; set; }
        public string activeSheetName { get; set; }
        public int sheetCount { get; set; }
        public List<SheetOverviewDto> sheets { get; set; }
        public string blueprintSummary { get; set; }

        public WorkbookOverviewResult()
        {
            sheets = new List<SheetOverviewDto>();
        }
    }

    /// <summary>
    /// Excel 原生 COM 公式体检与血缘审计服务 (C# 5.0 兼容器)
    /// </summary>
    public static class FormulaAuditService
    {
        private static readonly Regex FunctionRegex = new Regex(
            @"\b([A-Z_][A-Z0-9_\.]*)\s*\(",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        private static readonly Regex CellRefRegex = new Regex(
            @"(?:(?<sheet>'[^']+'|[a-zA-Z0-9_\u4e00-\u9fa5]+)\s*!\s*)?\$?(?<col>[a-zA-Z]{1,3})\$?(?<row>[0-9]{1,7})(?::\$?(?<col2>[a-zA-Z]{1,3})\$?(?<row2>[0-9]{1,7}))?",
            RegexOptions.Compiled
        );

        private static readonly HashSet<string> KnownExcelFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS",
            "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "XMATCH", "OFFSET",
            "IF", "IFS", "IFERROR", "IFNA", "AND", "OR", "NOT", "SWITCH", "CHOOSE",
            "LEFT", "RIGHT", "MID", "LEN", "TRIM", "CONCATENATE", "TEXTJOIN", "SUBSTITUTE",
            "DATE", "DATEDIF", "YEAR", "MONTH", "DAY", "TODAY", "NOW", "EDATE", "EOMONTH",
            "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "MOD", "ABS", "CEILING", "FLOOR",
            "MAX", "MIN", "LARGE", "SMALL", "RANK", "UNIQUE", "FILTER", "SORT", "SORTBY",
            "SEQUENCE", "CHOOSEROWS", "CHOOSECOLS", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS"
        };

        /// <summary>
        /// 全簿结构透视 (get_workbook_overview)
        /// 读取工作表清单、表头预览、可见性及公式密度
        /// </summary>
        public static WorkbookOverviewResult GetWorkbookOverview(dynamic app, string targetFullName, string targetName)
        {
            if (app == null) throw new InvalidOperationException("Excel Application 实例未就绪");

            dynamic wb = NativeBridge.FindTargetWorkbook(app, targetFullName, targetName);
            if (wb == null)
            {
                wb = app.ActiveWorkbook;
            }
            if (wb == null)
            {
                throw new InvalidOperationException("未找到活动工作簿，请先在 Excel 中打开目标文件");
            }

            var result = new WorkbookOverviewResult
            {
                workbookName = (string)wb.Name,
                fullName = (string)wb.FullName,
                activeSheetName = wb.ActiveSheet != null ? (string)wb.ActiveSheet.Name : "",
                sheetCount = wb.Worksheets.Count
            };

            var sbSummary = new StringBuilder();
            sbSummary.AppendLine(string.Format("工作簿【{0}】概览 (共 {1} 个工作表)：", result.workbookName, result.sheetCount));

            foreach (dynamic ws in wb.Worksheets)
            {
                var sheetDto = new SheetOverviewDto
                {
                    name = (string)ws.Name
                };

                // 可见性判定
                int vis = 1;
                try { vis = (int)ws.Visible; } catch { }
                sheetDto.visibility = vis == -1 ? "Visible" : (vis == 0 ? "Hidden" : "VeryHidden");

                // 已用区域
                try
                {
                    dynamic usedRange = ws.UsedRange;
                    if (usedRange != null)
                    {
                        sheetDto.usedRange = (string)usedRange.Address(false, false);
                        sheetDto.rowCount = (int)usedRange.Rows.Count;
                        sheetDto.columnCount = (int)usedRange.Columns.Count;

                        // 预览前 3 行表头
                        int previewRows = Math.Min(sheetDto.rowCount, 3);
                        int previewCols = Math.Min(sheetDto.columnCount, 15);
                        for (int r = 1; r <= previewRows; r++)
                        {
                            var rowList = new List<string>();
                            for (int c = 1; c <= previewCols; c++)
                            {
                                dynamic cell = usedRange.Cells[r, c];
                                string val = cell.Text != null ? cell.Text.ToString() : "";
                                rowList.Add(val);
                            }
                            sheetDto.headerPreview.Add(rowList);
                        }

                        // 公式单元格估算 (Safe SpecialCells)
                        try
                        {
                            dynamic formulaCells = usedRange.SpecialCells(0x00000003 /* xlCellTypeFormulas */);
                            if (formulaCells != null)
                            {
                                sheetDto.formulaCellCount = (int)formulaCells.Count;
                            }
                        }
                        catch
                        {
                            sheetDto.formulaCellCount = 0;
                        }
                    }
                }
                catch { }

                // 统计表格与图表数
                try { sheetDto.tableCount = ws.ListObjects.Count; } catch { sheetDto.tableCount = 0; }
                try { sheetDto.chartCount = ws.ChartObjects().Count; } catch { sheetDto.chartCount = 0; }

                result.sheets.Add(sheetDto);
                sbSummary.AppendLine(string.Format("- 工作表 [{0}] ({1}): 区域 {2}, {3} 行 × {4} 列, 包含约 {5} 个公式",
                    sheetDto.name, sheetDto.visibility, sheetDto.usedRange ?? "空", sheetDto.rowCount, sheetDto.columnCount, sheetDto.formulaCellCount));
            }

            result.blueprintSummary = sbSummary.ToString().TrimEnd();
            return result;
        }

        /// <summary>
        /// 公式依赖穿透与审计 (trace_dependencies & explain_formula)
        /// </summary>
        public static FormulaAuditResult AuditCellFormula(
            dynamic app,
            string targetFullName,
            string targetName,
            string sheetName,
            string cellAddress)
        {
            if (app == null) throw new InvalidOperationException("Excel Application 实例未就绪");

            dynamic wb = NativeBridge.FindTargetWorkbook(app, targetFullName, targetName);
            if (wb == null) wb = app.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("未找到目标工作簿");

            dynamic ws = null;
            if (!string.IsNullOrEmpty(sheetName))
            {
                try { ws = wb.Worksheets[sheetName]; } catch { }
            }
            if (ws == null) ws = wb.ActiveSheet;
            if (ws == null) throw new InvalidOperationException("未找到目标工作表");

            if (string.IsNullOrEmpty(cellAddress))
            {
                dynamic activeCell = app.ActiveCell;
                if (activeCell != null)
                {
                    cellAddress = (string)activeCell.Address(false, false);
                }
                else
                {
                    cellAddress = "A1";
                }
            }

            dynamic cell = ws.Range[cellAddress];
            if (cell == null) throw new InvalidOperationException(string.Format("无法定位单元格: {0}", cellAddress));

            string normalizedAddress = (string)cell.Address(false, false);
            string actualSheetName = (string)ws.Name;
            bool hasFormula = (bool)cell.HasFormula;
            string formula = hasFormula ? (string)cell.Formula : "";
            string text = cell.Text != null ? cell.Text.ToString() : "";
            object valObj = cell.Value2;
            string valStr = valObj != null ? valObj.ToString() : "";

            var result = new FormulaAuditResult
            {
                workbookName = (string)wb.Name,
                sheetName = actualSheetName,
                cellAddress = normalizedAddress,
                hasFormula = hasFormula,
                formula = formula,
                evaluatedValue = valStr,
                formattedText = text
            };

            // 检查单元格自身是否包含计算错误
            string errorType = DetectCellError(text, valObj);
            result.hasError = !string.IsNullOrEmpty(errorType);
            result.errorType = errorType;

            if (!hasFormula)
            {
                result.auditSummary = string.Format("单元格 [{0}!{1}] 为纯数值/常量内容 (无公式)。当前显示值: '{2}'。",
                    actualSheetName, normalizedAddress, text);
                return result;
            }

            // 1. 提取公式中使用的标准 Excel 函数列表
            var matches = FunctionRegex.Matches(formula);
            foreach (Match m in matches)
            {
                string fn = m.Groups[1].Value.ToUpperInvariant();
                if (KnownExcelFunctions.Contains(fn) && !result.functionsUsed.Contains(fn))
                {
                    result.functionsUsed.Add(fn);
                }
            }

            // 2. 向上追溯：获取直接前序引用 (Precedents)
            var precedentsMap = new Dictionary<string, CellDependencyNode>(StringComparer.OrdinalIgnoreCase);

            // A. 原生 COM Precedents 遍历
            try
            {
                dynamic precedentsRange = cell.Precedents;
                if (precedentsRange != null)
                {
                    foreach (dynamic precCell in precedentsRange)
                    {
                        string addr = (string)precCell.Address(false, false);
                        string sName = (string)precCell.Worksheet.Name;
                        string key = string.Format("{0}!{1}", sName, addr);
                        if (!precedentsMap.ContainsKey(key))
                        {
                            var node = CreateDependencyNode(precCell, sName, addr);
                            precedentsMap[key] = node;
                        }
                    }
                }
            }
            catch
            {
                // COM 在无引用或跨工作表引用时会抛出异常，降级为正则跨表探测
            }

            // B. 跨工作表与正则补充探测
            var refMatches = CellRefRegex.Matches(formula);
            foreach (Match rm in refMatches)
            {
                string refSheet = rm.Groups["sheet"].Value.Trim('\'', ' ');
                if (string.IsNullOrEmpty(refSheet)) refSheet = actualSheetName;
                string col = rm.Groups["col"].Value;
                string row = rm.Groups["row"].Value;
                string col2 = rm.Groups["col2"].Value;
                string row2 = rm.Groups["row2"].Value;

                if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row))
                {
                    string refAddr = string.IsNullOrEmpty(col2) ? string.Format("{0}{1}", col, row) : string.Format("{0}{1}:{2}{3}", col, row, col2, row2);
                    string fullKey = string.Format("{0}!{1}", refSheet, refAddr);

                    if (!precedentsMap.ContainsKey(fullKey))
                    {
                        try
                        {
                            dynamic targetWs = wb.Worksheets[refSheet];
                            dynamic refRange = targetWs.Range[refAddr];
                            if (refRange != null)
                            {
                                string refVal = refRange.Text != null ? refRange.Text.ToString() : "";
                                var node = new CellDependencyNode
                                {
                                    sheetName = refSheet,
                                    address = refAddr,
                                    value = refVal,
                                    formula = refRange.HasFormula ? (string)refRange.Formula : "",
                                    isBlank = string.IsNullOrWhiteSpace(refVal)
                                };
                                node.errorType = DetectCellError(node.value, refRange.Value2);
                                node.hasError = !string.IsNullOrEmpty(node.errorType);
                                precedentsMap[fullKey] = node;
                            }
                        }
                        catch { }
                    }
                }
            }

            foreach (var node in precedentsMap.Values)
            {
                result.directPrecedents.Add(node);
            }

            // 3. 向下追溯：获取直接从属引用 (Dependents)
            try
            {
                dynamic dependentsRange = cell.Dependents;
                if (dependentsRange != null)
                {
                    foreach (dynamic depCell in dependentsRange)
                    {
                        string depAddr = (string)depCell.Address(false, false);
                        string depSheet = (string)depCell.Worksheet.Name;
                        string depKey = string.Format("{0}!{1}", depSheet, depAddr);
                        if (!result.directDependents.Contains(depKey))
                        {
                            result.directDependents.Add(depKey);
                        }
                    }
                }
            }
            catch
            {
                // 无从属引用时忽略
            }

            // 4. 循环引用与空值/错误断链检测
            int errorCount = 0;
            int blankCount = 0;
            foreach (var node in result.directPrecedents)
            {
                if (node.hasError) errorCount++;
                if (node.isBlank) blankCount++;
                if (string.Equals(node.sheetName, actualSheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(node.address, normalizedAddress, StringComparison.OrdinalIgnoreCase))
                {
                    result.isCircular = true;
                }
            }

            // 5. 组装自然语言审计与体检摘要
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("单元格 [{0}!{1}] 公式审计报告：", actualSheetName, normalizedAddress));
            sb.AppendLine(string.Format("- 公式表达式: `{0}`", formula));
            sb.AppendLine(string.Format("- 当前计算结果: '{0}' {1}", text, result.hasError ? string.Format("[⚠️ 异常: {0}]", result.errorType) : "[正常]"));
            sb.AppendLine(string.Format("- 调用的函数: {0}", result.functionsUsed.Count > 0 ? string.Join(", ", result.functionsUsed.ToArray()) : "基本四则运算"));
            sb.AppendLine(string.Format("- 前序输入源 (依赖项): 共引用 {0} 处单元格", result.directPrecedents.Count));

            if (result.isCircular)
            {
                sb.AppendLine("  ⚠️ 严重警告：检测到该公式包含对自身的循环引用！");
            }
            if (errorCount > 0)
            {
                sb.AppendLine(string.Format("  ⚠️ 警告：检测到前序输入源中有 {0} 处包含计算错误（如 #VALUE! / #N/A），直接导致了计算断链。", errorCount));
            }
            if (blankCount > 0)
            {
                sb.AppendLine(string.Format("  💡 提示：检测到前序输入源中有 {0} 处为空单元格，可能被默认隐式转换为 0 计算。", blankCount));
            }

            sb.AppendLine(string.Format("- 后续影响面 (从属项): 共影响 {0} 处下游单元格计算", result.directDependents.Count));

            result.auditSummary = sb.ToString().TrimEnd();
            result.explanationPrompt = string.Format("请用通俗易懂的中文向非技术人员解释以下 Excel 公式的工作逻辑与风险点：\n{0}", result.auditSummary);

            return result;
        }

        private static CellDependencyNode CreateDependencyNode(dynamic precCell, string sheetName, string addr)
        {
            string val = precCell.Text != null ? precCell.Text.ToString() : "";
            bool hasF = false;
            try { hasF = (bool)precCell.HasFormula; } catch { }
            string f = hasF ? (string)precCell.Formula : "";

            var node = new CellDependencyNode
            {
                sheetName = sheetName,
                address = addr,
                value = val,
                formula = f,
                isBlank = string.IsNullOrWhiteSpace(val)
            };
            node.errorType = DetectCellError(val, precCell.Value2);
            node.hasError = !string.IsNullOrEmpty(node.errorType);
            return node;
        }

        private static string DetectCellError(string text, object value2)
        {
            if (!string.IsNullOrEmpty(text) && text.StartsWith("#"))
            {
                if (text.Contains("#N/A")) return "#N/A (缺失值)";
                if (text.Contains("#VALUE!")) return "#VALUE! (类型不匹配)";
                if (text.Contains("#REF!")) return "#REF! (无效引用)";
                if (text.Contains("#DIV/0!")) return "#DIV/0! (除以零)";
                if (text.Contains("#NAME?")) return "#NAME? (未识别的函数或名称)";
                if (text.Contains("#NUM!")) return "#NUM! (数值计算越界)";
                if (text.Contains("#NULL!")) return "#NULL! (空交集错误)";
                return text;
            }

            if (value2 is int)
            {
                int errInt = (int)value2;
                switch (errInt)
                {
                    case -2146826288: return "#NULL!";
                    case -2146826281: return "#DIV/0!";
                    case -2146826273: return "#VALUE!";
                    case -2146826265: return "#REF!";
                    case -2146826259: return "#NAME?";
                    case -2146826252: return "#NUM!";
                    case -2146826246: return "#N/A";
                }
            }

            return null;
        }
    }
}
