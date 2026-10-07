using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace LeeExcel
{
    /// <summary>
    /// 快捷图表工具参数契约 (TASK-R5b-01)
    /// </summary>
    public class QuickChartParams
    {
        public string targetWorkbookName { get; set; }
        public string targetWorkbookFullName { get; set; }
        public string sourceSheet { get; set; }
        public string sourceRange { get; set; } // 例如 "A1:D10"
        public bool hasHeaders { get; set; } // 表头是否存在
        public int dataStartRow { get; set; } // 相对区域的数据起始行 (有表头默认 2, 无表头 1)
        public int dataEndRow { get; set; } // 相对区域的数据结束行 (0 表示自动根据区域高度)
        public int categoryColIndex { get; set; } // 类别 (X轴) 列索引 (1-based 相对区域)
        public string categoryColName { get; set; } // 类别列名 (用于提示和读回核验)
        public List<int> seriesColIndices { get; set; } // 数值系列列索引列表 (1-based 相对区域)
        public List<string> seriesNames { get; set; } // 系列名称列表
        public string chartType { get; set; } // "column" | "line" | "pie"
        public string title { get; set; } // 图表标题
        public string targetSheet { get; set; } // 图表放置工作表 (为空默认 sourceSheet)
        public string placementMode { get; set; } // "cell" | "coordinates"
        public string targetCell { get; set; } // 例如 "F2"
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public string action { get; set; } // "create_new" | "replace_existing"
        public string targetChartId { get; set; } // 替换时必须指定
        public string errorHandling { get; set; } // "reject_on_invalid" | "coerce_zero"

        // 别名兼容属性（保障与前端组件及各种入参形态的 100% 互操作性）
        public string rangeAddress { get { return sourceRange; } set { if (!string.IsNullOrEmpty(value)) sourceRange = value; } }
        public bool hasHeader { get { return hasHeaders; } set { hasHeaders = value; } }
        public int categoryColumn { get { return categoryColIndex; } set { if (value > 0) categoryColIndex = value; } }
        public List<int> valueSeriesColumns { get { return seriesColIndices; } set { if (value != null && value.Count > 0) seriesColIndices = value; } }
        public string chartTitle { get { return title; } set { if (!string.IsNullOrEmpty(value)) title = value; } }
        public string placementSheet { get { return targetSheet; } set { if (!string.IsNullOrEmpty(value)) targetSheet = value; } }
        public string placementCell { get { return targetCell; } set { if (!string.IsNullOrEmpty(value)) targetCell = value; } }
        public string mode { get { return action; } set { if (!string.IsNullOrEmpty(value)) action = value; } }
        public string errorHandlingRule { get { return errorHandling; } set { if (!string.IsNullOrEmpty(value)) errorHandling = value; } }

        public QuickChartParams()
        {
            hasHeaders = true;
            dataStartRow = 2;
            dataEndRow = 0;
            categoryColIndex = 1;
            seriesColIndices = new List<int>();
            seriesNames = new List<string>();
            chartType = "column";
            placementMode = "cell";
            targetCell = "F2";
            width = 480;
            height = 300;
            action = "create_new";
            errorHandling = "reject_on_invalid";
        }
    }

    /// <summary>
    /// 图表数据绑定与属性读回 DTO
    /// </summary>
    public class ChartReadbackDto
    {
        public string chartId { get; set; }
        public string chartName { get; set; }
        public string sheetName { get; set; }
        public string chartType { get; set; }
        public int actualChartTypeNum { get; set; }
        public string title { get; set; }
        public int seriesCount { get; set; }
        public List<string> seriesNames { get; set; }
        public string categoryAddress { get; set; }
        public List<string> valuesAddresses { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public bool isReplaced { get; set; }

        public ChartReadbackDto()
        {
            seriesNames = new List<string>();
            valuesAddresses = new List<string>();
        }
    }

    /// <summary>
    /// 快捷图表执行结果
    /// </summary>
    public class QuickChartResult
    {
        public bool ok { get; set; }
        public string chartId { get; set; }
        public string chartName { get; set; }
        public string action { get; set; }
        public string summaryText { get; set; }
        public string snapshotId { get; set; }
        public ChartReadbackDto readback { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }
        public string recoveryNotice { get; set; }
    }

    /// <summary>
    /// 本工具管理的已有图表信息 (供前端下拉选择替换)
    /// </summary>
    public class ManagedChartInfo
    {
        public string chartId { get; set; }
        public string chartName { get; set; }
        public string title { get; set; }
        public string chartType { get; set; }
        public string sheetName { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
    }

    /// <summary>
    /// 本地图表工具服务 (TASK-R5b-01: 汇总/明细数据到图表轻量生成)
    /// </summary>
    public static class ChartService
    {
        public const string ChartNamePrefix = "__EM_CHART_";
        public const string GeneratorSignature = "\"generator\":\"ExcelMindAI\"";

        // Excel 常量定义
        public const int xlColumnClustered = 51;
        public const int xlLineMarkers = 65;
        public const int xlPie = 5;

        /// <summary>
        /// 解析图表类型枚举值
        /// </summary>
        public static bool TryResolveChartType(string chartType, out int chartTypeNum, out string normalizedName)
        {
            chartTypeNum = 0;
            normalizedName = "";
            if (string.IsNullOrEmpty(chartType)) return false;

            string lower = chartType.Trim().ToLowerInvariant();
            if (lower == "column" || lower == "bar" || lower == "51")
            {
                chartTypeNum = xlColumnClustered;
                normalizedName = "column";
                return true;
            }
            if (lower == "line" || lower == "65")
            {
                chartTypeNum = xlLineMarkers;
                normalizedName = "line";
                return true;
            }
            if (lower == "pie" || lower == "5")
            {
                chartTypeNum = xlPie;
                normalizedName = "pie";
                return true;
            }
            return false;
        }

        /// <summary>
        /// 扫描指定工作表上属于本工具管理的图表对象列表 (绝不列出或碰触用户的普通图表)
        /// </summary>
        public static List<ManagedChartInfo> ListManagedCharts(dynamic targetWb, string sheetName)
        {
            var list = new List<ManagedChartInfo>();
            if (targetWb == null || string.IsNullOrEmpty(sheetName)) return list;

            try
            {
                dynamic ws = targetWb.Worksheets[sheetName];
                if (ws == null) return list;

                dynamic chartObjects = ws.ChartObjects();
                int count = chartObjects.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic chObj = chartObjects.Item(i);
                    string name = (string)chObj.Name;
                    if (!string.IsNullOrEmpty(name) && name.StartsWith(ChartNamePrefix))
                    {
                        string chartId = name.Substring(ChartNamePrefix.Length);
                        string altText = "";
                        try { altText = (string)chObj.ShapeRange.AlternativeText; } catch { }

                        // 验证元数据签名
                        if (!string.IsNullOrEmpty(altText) && altText.Contains(GeneratorSignature))
                        {
                            string titleText = "";
                            try
                            {
                                if ((bool)chObj.Chart.HasTitle)
                                {
                                    titleText = (string)chObj.Chart.ChartTitle.Text;
                                }
                            }
                            catch { }

                            string cType = "unknown";
                            try
                            {
                                int ctNum = (int)chObj.Chart.ChartType;
                                if (ctNum == xlColumnClustered) cType = "column";
                                else if (ctNum == xlLineMarkers) cType = "line";
                                else if (ctNum == xlPie) cType = "pie";
                            }
                            catch { }

                            list.Add(new ManagedChartInfo
                            {
                                chartId = chartId,
                                chartName = name,
                                title = titleText,
                                chartType = cType,
                                sheetName = sheetName,
                                left = (double)chObj.Left,
                                top = (double)chObj.Top,
                                width = (double)chObj.Width,
                                height = (double)chObj.Height
                            });
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        /// <summary>
        /// 执行快捷图表生成或替换 (参数校验、数据质量核查、强制前置快照、精准绑定与写后读回)
        /// </summary>
        public static QuickChartResult ExecuteQuickChart(
            dynamic app,
            QuickChartParams p,
            bool skipSnapshot = false,
            string existingSnapshotId = null)
        {
            var res = new QuickChartResult
            {
                ok = false,
                action = p != null ? p.action : "create_new",
                failureStage = "precheck"
            };

            if (p == null)
            {
                res.error = "图表入参为空。";
                return res;
            }

            // 1. 图表类型支持性校验
            int chartTypeNum;
            string normalizedChartType;
            if (!TryResolveChartType(p.chartType, out chartTypeNum, out normalizedChartType))
            {
                res.error = string.Format(
                    "快捷图表工具第一切片仅支持柱状图 (column)、折线图 (line) 与饼图 (pie)，不支持图表类型【{0}】。注：此限制仅为快捷工具的初始支持边界，并不限制大模型自主生成任意原生 Excel 图表宏。",
                    p.chartType);
                return res;
            }

            // 2. 饼图单系列约束检查
            if (normalizedChartType == "pie" && p.seriesColIndices != null && p.seriesColIndices.Count > 1)
            {
                res.error = string.Format(
                    "饼图仅支持单数值系列，当前参数指定了 {0} 个数值系列。请仅指定 1 个数值系列，或改用柱状图/折线图展示多系列数据。",
                    p.seriesColIndices.Count);
                return res;
            }

            // 3. 维度与数值列配置检查
            if (p.categoryColIndex < 1)
            {
                res.error = "必须指定有效的类别 (X轴) 列索引 (categoryColIndex 必须 >= 1)。";
                return res;
            }
            if (p.seriesColIndices == null || p.seriesColIndices.Count == 0)
            {
                res.error = "必须指定至少一个数值系列列索引 (seriesColIndices)。";
                return res;
            }

            // 4. 锁定目标工作簿 (严格按 FullName / Name 查找，绝不 fallback 到 ActiveWorkbook)
            if (app == null)
            {
                res.error = "Excel 宿主实例为空。";
                return res;
            }

            dynamic targetWb = null;
            try
            {
                if (!string.IsNullOrEmpty(p.targetWorkbookFullName))
                {
                    foreach (dynamic wb in app.Workbooks)
                    {
                        if (string.Equals((string)wb.FullName, p.targetWorkbookFullName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = wb;
                            break;
                        }
                    }
                }
                if (targetWb == null && !string.IsNullOrEmpty(p.targetWorkbookName))
                {
                    foreach (dynamic wb in app.Workbooks)
                    {
                        if (string.Equals((string)wb.Name, p.targetWorkbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = wb;
                            break;
                        }
                    }
                }
            }
            catch (Exception exWb)
            {
                res.error = "查找目标工作簿异常: " + exWb.Message;
                return res;
            }

            if (targetWb == null)
            {
                res.error = string.Format(
                    "未找到指定的目标工作簿 (FullName: {0}, Name: {1})。系统严格拒绝隐式回退当前活动工作簿，防止串改！",
                    p.targetWorkbookFullName ?? "空", p.targetWorkbookName ?? "空");
                return res;
            }

            // 5. 查找数据源工作表与目标放置工作表
            if (string.IsNullOrEmpty(p.sourceSheet))
            {
                res.error = "未指定数据源工作表名称 (sourceSheet)。系统拒绝猜测工作表！";
                return res;
            }

            dynamic wsSource = null;
            try
            {
                foreach (dynamic ws in targetWb.Worksheets)
                {
                    if (string.Equals((string)ws.Name, p.sourceSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        wsSource = ws;
                        break;
                    }
                }
            }
            catch { }

            if (wsSource == null)
            {
                res.error = string.Format("目标工作簿中未找到数据源工作表【{0}】。", p.sourceSheet);
                return res;
            }

            string targetSheetName = string.IsNullOrEmpty(p.targetSheet) ? p.sourceSheet : p.targetSheet;
            dynamic wsTarget = null;
            try
            {
                foreach (dynamic ws in targetWb.Worksheets)
                {
                    if (string.Equals((string)ws.Name, targetSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        wsTarget = ws;
                        break;
                    }
                }
            }
            catch { }

            if (wsTarget == null)
            {
                res.error = string.Format("目标工作簿中未找到放置图表的工作表【{0}】。系统拒绝猜测工作表！", targetSheetName);
                return res;
            }

            // 6. 确定数据源区域
            if (string.IsNullOrEmpty(p.sourceRange))
            {
                res.error = "未指定数据源区域 (sourceRange)。系统拒绝盲目猜选全表！";
                return res;
            }

            dynamic srcRng = null;
            try
            {
                srcRng = wsSource.Range[p.sourceRange];
            }
            catch (Exception exRng)
            {
                res.error = string.Format("解析数据源区域【{0}】失败: {1}", p.sourceRange, exRng.Message);
                return res;
            }

            int totalRows = srcRng.Rows.Count;
            int totalCols = srcRng.Columns.Count;

            // 校验类别列与系列列索引是否超界
            if (p.categoryColIndex > totalCols)
            {
                res.error = string.Format("类别列索引 ({0}) 超出数据区域总列数 ({1})。", p.categoryColIndex, totalCols);
                return res;
            }
            foreach (int sIdx in p.seriesColIndices)
            {
                if (sIdx < 1 || sIdx > totalCols)
                {
                    res.error = string.Format("数值系列列索引 ({0}) 超出数据区域总列数 ({1})。", sIdx, totalCols);
                    return res;
                }
            }

            // 确定行边界
            int headerRelRow = p.hasHeaders ? 1 : 0;
            int dataStartRow = p.dataStartRow > 0 ? p.dataStartRow : (p.hasHeaders ? 2 : 1);
            int dataEndRow = p.dataEndRow > 0 ? p.dataEndRow : totalRows;

            if (dataStartRow > totalRows || dataStartRow > dataEndRow)
            {
                res.error = string.Format("数据起始行 ({0}) 非法，区域总行数为 {1}。", dataStartRow, totalRows);
                return res;
            }

            int dataPointCount = dataEndRow - dataStartRow + 1;
            if (dataPointCount <= 0)
            {
                res.error = "可绘图的数据点行数必须大于 0。";
                return res;
            }

            // 7. 读取并严格校验数值系列列与类别列数据质量
            object[,] matrix = null;
            try
            {
                object rawVal = srcRng.Value2;
                if (rawVal is object[,])
                {
                    matrix = (object[,])rawVal;
                }
                else
                {
                    // 单个单元格
                    matrix = new object[2, 2];
                    matrix[1, 1] = rawVal;
                }
            }
            catch (Exception exRead)
            {
                res.error = "读取数据区域 Value2 失败: " + exRead.Message;
                return res;
            }

            // 校验非数值内容、错误值与空值规则
            bool isCoerceZero = string.Equals(p.errorHandling, "coerce_zero", StringComparison.OrdinalIgnoreCase);
            for (int r = dataStartRow; r <= dataEndRow; r++)
            {
                foreach (int sIdx in p.seriesColIndices)
                {
                    object cellVal = matrix[r, sIdx];
                    bool isInvalid = false;
                    string cellDesc = "";

                    if (cellVal == null || cellVal == DBNull.Value)
                    {
                        isInvalid = true;
                        cellDesc = "空值";
                    }
                    else if (cellVal is int && (int)cellVal < 0)
                    {
                        // COM 单元格错误代码 (如 #N/A, #DIV/0!)
                        isInvalid = true;
                        cellDesc = "Excel公式错误 (" + cellVal.ToString() + ")";
                    }
                    else
                    {
                        string strVal = cellVal.ToString().Trim();
                        if (strVal.StartsWith("#") && strVal.Length > 2)
                        {
                            isInvalid = true;
                            cellDesc = "Excel错误值 (" + strVal + ")";
                        }
                        else
                        {
                            double dblVal;
                            if (!double.TryParse(strVal, out dblVal))
                            {
                                isInvalid = true;
                                cellDesc = "非数值文本 ('" + strVal + "')";
                            }
                        }
                    }

                    if (isInvalid)
                    {
                        if (!isCoerceZero)
                        {
                            // reject_on_invalid 模式下直接透明阻断
                            string cellAddr = "";
                            try
                            {
                                cellAddr = (string)srcRng.Cells[r, sIdx].Address;
                            }
                            catch { }

                            res.error = string.Format(
                                "数值系列列 (第 {0} 列) 在第 {1} 行单元格【{2}】发现非法数据: {3}。已安全阻断。请检查数据或在参数中指定容错规则 (coerce_zero)。",
                                sIdx, r, cellAddr, cellDesc);
                            return res;
                        }
                    }
                }
            }

            // 8. 替换模式专属身份检查 (防旧结果叠加与误删已有图表)
            dynamic oldChartObj = null;
            double targetLeft = p.left;
            double targetTop = p.top;
            double targetWidth = p.width > 0 ? p.width : 480;
            double targetHeight = p.height > 0 ? p.height : 300;

            if (string.Equals(p.action, "replace_existing", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(p.targetChartId))
                {
                    res.error = "选择【替换指定图表】模式时，必须指定待替换的图表标识 (targetChartId)。";
                    return res;
                }

                string cleanTargetId = p.targetChartId.Trim();
                string expectedShapeName = cleanTargetId.StartsWith(ChartNamePrefix) ? cleanTargetId : ChartNamePrefix + cleanTargetId;
                try
                {
                    dynamic chartObjects = wsTarget.ChartObjects();
                    int cCount = chartObjects.Count;
                    for (int i = 1; i <= cCount; i++)
                    {
                        dynamic obj = chartObjects.Item(i);
                        string oName = (string)obj.Name;
                        if (string.Equals(oName, expectedShapeName, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(oName, cleanTargetId, StringComparison.OrdinalIgnoreCase))
                        {
                            oldChartObj = obj;
                            break;
                        }
                    }
                }
                catch { }

                if (oldChartObj == null)
                {
                    res.error = string.Format(
                        "未在工作表【{0}】上找到指定的待替换图表对象【{1}】（可能已被移动或删除）。系统已安全阻断，未碰触或删除任何已有图表！",
                        targetSheetName, expectedShapeName);
                    return res;
                }

                // 核实对象身份，绝不仅凭标题认定！
                string altText = "";
                try { altText = (string)oldChartObj.ShapeRange.AlternativeText; } catch { }
                if (string.IsNullOrEmpty(altText) || !altText.Contains(GeneratorSignature))
                {
                    res.error = string.Format(
                        "目标图表对象【{0}】缺少本工具的安全管理签名，属于用户手工创建或其他来源图表。系统严格拒绝替换或删除非本工具拥有的图表！",
                        expectedShapeName);
                    return res;
                }

                // 继承原图表位置与尺寸
                try
                {
                    targetLeft = (double)oldChartObj.Left;
                    targetTop = (double)oldChartObj.Top;
                    targetWidth = (double)oldChartObj.Width;
                    targetHeight = (double)oldChartObj.Height;
                }
                catch { }
            }
            else
            {
                // 新建图表模式：如果未指定显式坐标，计算 targetCell 的位置
                if (string.Equals(p.placementMode, "cell", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(p.targetCell))
                {
                    try
                    {
                        dynamic cellRng = wsTarget.Range[p.targetCell];
                        targetLeft = (double)cellRng.Left;
                        targetTop = (double)cellRng.Top;
                    }
                    catch { }
                }
            }

            // 9. 强制前置整本物理快照 (承诺快照失败零图表修改)
            res.failureStage = "snapshot";
            if (!skipSnapshot)
            {
                string snapPrompt = string.Format("快捷图表工具: 【{0}】({1}) 执行前整本物理备份",
                    !string.IsNullOrEmpty(p.title) ? p.title : normalizedChartType,
                    p.action == "replace_existing" ? "替换" : "新建");

                SnapshotItem snapItem = null;
                try
                {
                    snapItem = SnapshotManager.CreateSnapshot(targetWb, snapPrompt, "");
                }
                catch (Exception exSnap)
                {
                    res.error = "执行前整本快照创建异常 (" + exSnap.Message + ")。承诺快照失败零图表修改，已安全阻断。";
                    return res;
                }

                if (snapItem == null || string.IsNullOrEmpty(snapItem.id) || !SnapshotManager.SnapshotExists((string)targetWb.FullName, snapItem.id))
                {
                    res.error = "执行前整本快照物理副本写入失败。承诺快照失败零图表修改，已安全阻断。";
                    return res;
                }

                res.snapshotId = snapItem.id;
            }
            else
            {
                res.snapshotId = existingSnapshotId;
            }

            // 10. 执行图表创建 / 替换
            res.failureStage = "execution";
            string chartId = string.Equals(p.action, "replace_existing", StringComparison.OrdinalIgnoreCase)
                ? p.targetChartId.Trim()
                : Guid.NewGuid().ToString("N").Substring(0, 8);
            string newChartName = ChartNamePrefix + chartId;

            dynamic newChartObj = null;
            dynamic catRange = null;
            try
            {
                // 若为替换，先安全删除旧对象
                if (oldChartObj != null)
                {
                    try { oldChartObj.Delete(); } catch { }
                    oldChartObj = null;
                }

                dynamic chartObjects = wsTarget.ChartObjects();
                newChartObj = chartObjects.Add(targetLeft, targetTop, targetWidth, targetHeight);
                newChartObj.Name = newChartName;

                // 写入身份签名元数据至 AlternativeText
                string metaJson = string.Format(
                    "{{\"generator\":\"ExcelMindAI\",\"tool\":\"quick_chart\",\"chartId\":\"{0}\",\"chartType\":\"{1}\",\"title\":\"{2}\",\"createdAt\":\"{3}\"}}",
                    chartId, normalizedChartType, (p.title ?? "").Replace("\"", "\\\""), DateTime.Now.ToString("o"));
                try { newChartObj.ShapeRange.AlternativeText = metaJson; } catch { }

                dynamic chart = newChartObj.Chart;
                chart.ChartType = chartTypeNum;

                // 设置标题
                if (!string.IsNullOrEmpty(p.title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = p.title;
                }
                else
                {
                    chart.HasTitle = false;
                }

                // 清空默认生成的系列
                dynamic seriesCol = chart.SeriesCollection();
                while (seriesCol.Count > 0)
                {
                    seriesCol.Item(1).Delete();
                }

                // 精准绑定类别列与每个数值系列列
                catRange = wsSource.Range[
                    wsSource.Cells[srcRng.Row + dataStartRow - 1, srcRng.Column + p.categoryColIndex - 1],
                    wsSource.Cells[srcRng.Row + dataEndRow - 1, srcRng.Column + p.categoryColIndex - 1]
                ];

                for (int s = 0; s < p.seriesColIndices.Count; s++)
                {
                    int colRelIdx = p.seriesColIndices[s];
                    dynamic series = seriesCol.NewSeries();

                    // 系列名称获取
                    string sName = "";
                    if (p.seriesNames != null && s < p.seriesNames.Count && !string.IsNullOrEmpty(p.seriesNames[s]))
                    {
                        sName = p.seriesNames[s];
                    }
                    else if (p.hasHeaders)
                    {
                        try
                        {
                            object hdrObj = matrix[1, colRelIdx];
                            if (hdrObj != null) sName = hdrObj.ToString();
                        }
                        catch { }
                    }
                    if (string.IsNullOrEmpty(sName))
                    {
                        sName = "系列 " + (s + 1);
                    }
                    series.Name = sName;

                    // 绑定 X 轴类别与数值
                    series.XValues = catRange;

                    dynamic valRange = wsSource.Range[
                        wsSource.Cells[srcRng.Row + dataStartRow - 1, srcRng.Column + colRelIdx - 1],
                        wsSource.Cells[srcRng.Row + dataEndRow - 1, srcRng.Column + colRelIdx - 1]
                    ];
                    series.Values = valRange;
                }
            }
            catch (Exception exExec)
            {
                res.error = "创建或配置 Excel 图表对象异常: " + exExec.Message;
                res.recoveryNotice = string.Format(
                    "执行中发生异常。若目标工作表中已生成未完成的图表残片，用户可按已验证范围恢复目标工作簿（快照 ID: {0}）。",
                    res.snapshotId ?? "无");
                return res;
            }

            // 11. COM 读回客观核验 (系列名称、数量、类型、位置坐标)
            res.failureStage = "readback";
            var readback = new ChartReadbackDto
            {
                chartId = chartId,
                chartName = newChartName,
                sheetName = targetSheetName,
                chartType = normalizedChartType,
                actualChartTypeNum = chartTypeNum,
                isReplaced = string.Equals(p.action, "replace_existing", StringComparison.OrdinalIgnoreCase)
            };

            try
            {
                readback.left = (double)newChartObj.Left;
                readback.top = (double)newChartObj.Top;
                readback.width = (double)newChartObj.Width;
                readback.height = (double)newChartObj.Height;

                dynamic chart = newChartObj.Chart;
                try
                {
                    if ((bool)chart.HasTitle)
                    {
                        readback.title = (string)chart.ChartTitle.Text;
                    }
                }
                catch { }

                dynamic seriesCol = chart.SeriesCollection();
                readback.seriesCount = (int)seriesCol.Count;

                for (int i = 1; i <= readback.seriesCount; i++)
                {
                    dynamic s = seriesCol.Item(i);
                    try { readback.seriesNames.Add((string)s.Name); } catch { }
                    try
                    {
                        // 读回公式或地址表示
                        string formula = (string)s.Formula;
                        readback.valuesAddresses.Add(formula ?? "");
                    }
                    catch { }
                }

                try
                {
                    readback.categoryAddress = (string)catRange.Address;
                }
                catch { }
            }
            catch (Exception exRb)
            {
                res.error = "读回图表属性核验异常: " + exRb.Message;
                res.recoveryNotice = string.Format(
                    "图表可能已创建但在读回核验时异常。用户可按已验证范围恢复目标工作簿（快照 ID: {0}）。",
                    res.snapshotId ?? "无");
                return res;
            }

            res.ok = true;
            res.chartId = chartId;
            res.chartName = newChartName;
            res.readback = readback;
            res.failureStage = "completed";
            res.summaryText = string.Format(
                "已成功在工作表【{0}】{1}图表【{2}】（类型: {3}，包含 {4} 个数值系列，已绑定类别列【{5}】）。",
                targetSheetName,
                readback.isReplaced ? "替换" : "新建",
                !string.IsNullOrEmpty(readback.title) ? readback.title : readback.chartName,
                readback.chartType,
                readback.seriesCount,
                !string.IsNullOrEmpty(p.categoryColName) ? p.categoryColName : ("第" + p.categoryColIndex + "列")
            );
            res.recoveryNotice = string.Format(
                "图表操作已完成；若需撤销，用户可按已验证范围恢复目标工作簿（快照 ID: {0}）。",
                res.snapshotId ?? "无");

            return res;
        }
    }
}
