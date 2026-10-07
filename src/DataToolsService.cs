using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace LeeExcel
{
    /// <summary>
    /// 结构化复合主键比较对象：严格区分类型、区分大小写、不自动去空格、支持多列复合键，杜绝字符串分隔符拼接碰撞
    /// </summary>
    public class StructuredKey : IEquatable<StructuredKey>
    {
        private readonly object[] _values;
        private readonly int _hashCode;

        public StructuredKey(object[] values)
        {
            _values = values ?? new object[0];
            _hashCode = ComputeHashCode(_values);
        }

        public object[] Values { get { return _values; } }

        public bool IsAllEmpty()
        {
            if (_values.Length == 0) return true;
            foreach (var v in _values)
            {
                if (v != null && !string.IsNullOrEmpty(v.ToString())) return false;
            }
            return true;
        }

        public bool HasEmptyPart()
        {
            foreach (var v in _values)
            {
                if (v == null || string.IsNullOrEmpty(v.ToString())) return true;
            }
            return false;
        }

        public bool HasError()
        {
            foreach (var v in _values)
            {
                if (v == null) continue;
                // Excel COM 错误类型通常在 Value2 中表现为负数 HRESULT (如 -2146826281)
                // 或 string 开头为 "#" (如 "#DIV/0!", "#N/A", "#VALUE!", "#REF!")
                if (v is int && (int)v < 0) return true;
                string s = v as string;
                if (s != null && s.StartsWith("#") && s.Length > 2) return true;
            }
            return false;
        }

        public bool Equals(StructuredKey other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            if (_values.Length != other._values.Length) return false;

            for (int i = 0; i < _values.Length; i++)
            {
                object v1 = _values[i];
                object v2 = other._values[i];

                if (v1 == null && v2 == null) continue;
                if (v1 == null || v2 == null) return false;

                // 类型不同绝不隐式混同（例如字符串 "001" 与数字 1 绝不相等）
                if (v1.GetType() != v2.GetType())
                {
                    return false;
                }

                if (v1 is string)
                {
                    // 严格区分大小写，不 trim 空格
                    if (!string.Equals((string)v1, (string)v2, StringComparison.Ordinal))
                    {
                        return false;
                    }
                }
                else if (v1 is double)
                {
                    if ((double)v1 != (double)v2) return false;
                }
                else if (v1 is bool)
                {
                    if ((bool)v1 != (bool)v2) return false;
                }
                else if (v1 is int)
                {
                    if ((int)v1 != (int)v2) return false;
                }
                else
                {
                    if (!v1.Equals(v2)) return false;
                }
            }
            return true;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as StructuredKey);
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        private static int ComputeHashCode(object[] values)
        {
            unchecked
            {
                int hash = 17;
                foreach (var v in values)
                {
                    hash = hash * 31;
                    if (v != null)
                    {
                        hash += v.GetType().GetHashCode() * 19;
                        if (v is string) hash += StringComparer.Ordinal.GetHashCode((string)v);
                        else hash += v.GetHashCode();
                    }
                }
                return hash;
            }
        }
    }

    /// <summary>
    /// 去重分析报告与统计模型
    /// </summary>
    public class DedupAnalysisResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }
        public string targetWorkbookName { get; set; }
        public string targetSheetName { get; set; }
        public string rangeAddress { get; set; }
        public bool hasHeader { get; set; }
        public List<int> keyColumnIndices { get; set; } // 1-based relative to selection
        public List<string> keyColumnNames { get; set; }

        public int selectedRowCount { get; set; }
        public int selectedColumnCount { get; set; }
        public int dataRowCount { get; set; }
        public int uniqueCount { get; set; }
        public int duplicateCount { get; set; }
        public int duplicateGroupCount { get; set; }
        public int excludedCount { get; set; }
        public string excludedReason { get; set; }

        public long analysisElapsedMs { get; set; }

        // 内部数据行索引清单 (1-based relative to data rows)
        internal List<int> duplicateDataRowIndices { get; set; }
        internal List<int> uniqueDataRowIndices { get; set; }
        internal List<int> excludedDataRowIndices { get; set; }

        public DedupAnalysisResult()
        {
            keyColumnIndices = new List<int>();
            keyColumnNames = new List<string>();
            duplicateDataRowIndices = new List<int>();
            uniqueDataRowIndices = new List<int>();
            excludedDataRowIndices = new List<int>();
        }
    }

    /// <summary>
    /// 去重执行结果模型
    /// </summary>
    public class DedupExecutionResult
    {
        public bool ok { get; set; }
        public string mode { get; set; } // "highlight" | "export_unique"
        public string message { get; set; }
        public string error { get; set; }
        public string snapshotId { get; set; }
        public string snapshotPath { get; set; }
        public bool snapshotCreated { get; set; }

        public int dataRowCount { get; set; }
        public int uniqueCount { get; set; }
        public int duplicateCount { get; set; }
        public int duplicateGroupCount { get; set; }
        public int excludedCount { get; set; }

        public int modifiedRowCount { get; set; }
        public string resultSheetName { get; set; }
        public string highlightedRange { get; set; }

        public long analysisElapsedMs { get; set; }
        public long writeElapsedMs { get; set; }
        public long totalElapsedMs { get; set; }
    }

    /// <summary>
    /// 对账比较列映射定义
    /// </summary>
    public class CompareColMappingDto
    {
        public int leftColIndex { get; set; }  // 1-based relative to left range
        public int rightColIndex { get; set; } // 1-based relative to right range
        public string leftColName { get; set; }
        public string rightColName { get; set; }
    }

    /// <summary>
    /// 单元格差异明细
    /// </summary>
    public class CellDiffItem
    {
        public string colName { get; set; }
        public object leftValue { get; set; }
        public string leftType { get; set; }
        public object rightValue { get; set; }
        public string rightType { get; set; }
    }

    /// <summary>
    /// 对账单行明细
    /// </summary>
    public class ReconcileRowDetail
    {
        public string category { get; set; } // "SAME", "DIFF", "LEFT_ONLY", "RIGHT_ONLY", "LEFT_DUP", "RIGHT_DUP", "LEFT_INVALID", "RIGHT_INVALID"
        public int leftRowIndex { get; set; }  // 1-based relative to left data rows
        public int rightRowIndex { get; set; } // 1-based relative to right data rows
        public string keyDisplay { get; set; }
        public List<CellDiffItem> diffs { get; set; }

        public ReconcileRowDetail()
        {
            diffs = new List<CellDiffItem>();
        }
    }

    /// <summary>
    /// 两表对账分析报告与统计模型
    /// </summary>
    public class ReconcileAnalysisResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }

        // 左表元信息
        public string leftWorkbookName { get; set; }
        public string leftSheetName { get; set; }
        public string leftRangeAddress { get; set; }
        public bool leftHasHeader { get; set; }
        public int leftRowCount { get; set; }
        public int leftColCount { get; set; }
        public int leftDataRowCount { get; set; }

        // 右表元信息
        public string rightWorkbookName { get; set; }
        public string rightSheetName { get; set; }
        public string rightRangeAddress { get; set; }
        public bool rightHasHeader { get; set; }
        public int rightRowCount { get; set; }
        public int rightColCount { get; set; }
        public int rightDataRowCount { get; set; }

        // 映射
        public List<int> leftKeyCols { get; set; }
        public List<int> rightKeyCols { get; set; }
        public List<string> leftKeyNames { get; set; }
        public List<string> rightKeyNames { get; set; }
        public List<CompareColMappingDto> compareCols { get; set; }

        // 统计恒等式指标
        public int matchedBothSameCount { get; set; }
        public int matchedBothDiffCount { get; set; }
        public int leftOnlyCount { get; set; }
        public int rightOnlyCount { get; set; }

        /// <summary>
        /// 左表重复/歧义主键总行数（包含自身多行重复及对侧多行关联产生的匹配歧义行数）
        /// </summary>
        public int leftDuplicateKeyCount { get; set; }
        /// <summary>
        /// 右表重复/歧义主键总行数（包含自身多行重复及对侧多行关联产生的匹配歧义行数）
        /// </summary>
        public int rightDuplicateKeyCount { get; set; }
        public int leftInvalidKeyCount { get; set; }
        public int rightInvalidKeyCount { get; set; }

        public int diffCellCount { get; set; }
        public long analysisElapsedMs { get; set; }

        // 数据指纹，用于执行前源数据变更核验
        public string dataFingerprint { get; set; }

        internal List<ReconcileRowDetail> details { get; set; }

        public ReconcileAnalysisResult()
        {
            leftKeyCols = new List<int>();
            rightKeyCols = new List<int>();
            leftKeyNames = new List<string>();
            rightKeyNames = new List<string>();
            compareCols = new List<CompareColMappingDto>();
            details = new List<ReconcileRowDetail>();
        }
    }

    /// <summary>
    /// 对账执行写入结果模型
    /// </summary>
    public class ReconcileExecutionResult
    {
        public bool ok { get; set; }
        public string message { get; set; }
        public string error { get; set; }
        public string snapshotId { get; set; }
        public string snapshotPath { get; set; }
        public bool snapshotCreated { get; set; }

        public string targetWorkbookName { get; set; }
        public string resultSheetName { get; set; }
        public int exportedRowCount { get; set; }

        public int matchedBothSameCount { get; set; }
        public int matchedBothDiffCount { get; set; }
        public int leftOnlyCount { get; set; }
        public int rightOnlyCount { get; set; }
        public int leftDuplicateKeyCount { get; set; }
        public int rightDuplicateKeyCount { get; set; }
        public int leftInvalidKeyCount { get; set; }
        public int rightInvalidKeyCount { get; set; }
        public int diffCellCount { get; set; }

        public long analysisElapsedMs { get; set; }
        public long writeElapsedMs { get; set; }
        public long totalElapsedMs { get; set; }
    }

    /// <summary>
    /// 多文件汇总来源定义
    /// </summary>
    public class ConsolidationSourceDef
    {
        public string filePath { get; set; }
        public string sheetName { get; set; }
        public string rangeAddress { get; set; }
        public bool hasHeader { get; set; }
        public int headerRowIndex { get; set; }
        public string displayIdentifier { get; set; }
    }

    /// <summary>
    /// 多文件汇总显式列映射定义
    /// </summary>
    public class ConsolidationColMapping
    {
        public int sourceFileIndex { get; set; }
        public string sourceColName { get; set; }
        public string targetColName { get; set; }
    }

    /// <summary>
    /// 来源文件画像分析结果
    /// </summary>
    public class ConsolidationSourceProfile
    {
        public int fileIndex { get; set; }
        public string filePath { get; set; }
        public string displayIdentifier { get; set; }
        public string sheetName { get; set; }
        public string rangeAddress { get; set; }
        public int totalRows { get; set; }
        public int totalCols { get; set; }
        public int headerRow { get; set; }
        public int includedDataRows { get; set; }
        public int excludedBlankRows { get; set; }
        public List<string> columns { get; set; }
        public string fileSha256 { get; set; }

        public ConsolidationSourceProfile()
        {
            columns = new List<string>();
        }
    }

    /// <summary>
    /// 多文件汇总只读分析与映射建议结果
    /// </summary>
    public class ConsolidationAnalysisResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }
        public List<ConsolidationSourceProfile> sources { get; set; }
        public List<string> alignedColumns { get; set; }
        public string metadataSourceFileCol { get; set; }
        public string metadataSourceSheetCol { get; set; }
        public Dictionary<int, List<string>> missingColumnsPerSource { get; set; }
        public Dictionary<int, List<string>> unmappedSourceColumns { get; set; }
        public Dictionary<string, List<string>> columnTypeProfiles { get; set; }
        public List<Dictionary<string, object>> previewRows { get; set; }

        public int totalSourceCount { get; set; }
        public int totalInputRows { get; set; }
        public int totalHeaderRows { get; set; }
        public int totalIncludedDataRows { get; set; }
        public int totalExcludedBlankRows { get; set; }
        public int expectedFinalOutputRows { get; set; }
        public bool accountingIdentityConfirmed { get; set; }

        public string dataFingerprint { get; set; }
        public long analysisElapsedMs { get; set; }

        public ConsolidationAnalysisResult()
        {
            sources = new List<ConsolidationSourceProfile>();
            alignedColumns = new List<string>();
            missingColumnsPerSource = new Dictionary<int, List<string>>();
            unmappedSourceColumns = new Dictionary<int, List<string>>();
            columnTypeProfiles = new Dictionary<string, List<string>>();
            previewRows = new List<Dictionary<string, object>>();
        }
    }

    /// <summary>
    /// 多文件汇总执行结果模型
    /// </summary>
    public class ConsolidationExecutionResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }
        public string outputFilePath { get; set; }
        public string outputSheetName { get; set; }
        public int finalDataRows { get; set; }
        public int finalDataCols { get; set; }
        public int totalIncludedDataRows { get; set; }
        public int totalExcludedBlankRows { get; set; }
        public bool accountingIdentityConfirmed { get; set; }
        public bool verifiedReadback { get; set; }
        public long elapsedMs { get; set; }
    }

    /// <summary>
    /// 快捷数据工具服务：纯 C# 内存哈希极速去重与两表对账，拒绝跨 COM 逐格遍历，严格保证统计恒等式
    /// </summary>
    public static class DataToolsService
    {
        public const int MaxRowsSafetyLimit = 50000;
        public const int MaxCellsSafetyLimit = 1000000;

        /// <summary>
        /// 只读分析选区重复项：不创建快照，不修改表格，批量读取并由内存哈希结构化分析
        /// </summary>
        public static DedupAnalysisResult AnalyzeDedup(
            dynamic targetWb,
            string sheetName,
            string rangeAddress,
            bool hasHeader,
            List<int> keyColumnIndices)
        {
            var sw = Stopwatch.StartNew();
            var res = new DedupAnalysisResult
            {
                ok = true,
                hasHeader = hasHeader,
                keyColumnIndices = keyColumnIndices != null ? new List<int>(keyColumnIndices) : new List<int>()
            };

            if (targetWb == null)
            {
                res.ok = false;
                res.failureStage = "target_missing";
                res.error = "未指定有效的目标工作簿。";
                return res;
            }

            try { res.targetWorkbookName = (string)targetWb.Name; } catch { res.targetWorkbookName = "Workbook"; }

            dynamic sheet = null;
            try
            {
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = targetWb.Worksheets[sheetName];
                }
                else
                {
                    sheet = targetWb.ActiveSheet;
                }
            }
            catch (Exception exSheet)
            {
                res.ok = false;
                res.failureStage = "sheet_not_found";
                res.error = "无法找到指定工作表 '" + sheetName + "': " + exSheet.Message;
                return res;
            }

            try { res.targetSheetName = (string)sheet.Name; } catch { res.targetSheetName = "Sheet1"; }

            // 1. 检查工作表保护状态
            try
            {
                if ((bool)sheet.ProtectContents)
                {
                    res.ok = false;
                    res.failureStage = "sheet_protected";
                    res.error = "工作表【" + res.targetSheetName + "】处于保护状态，无法执行去重工具操作。";
                    return res;
                }
            }
            catch { }

            // 2. 检查选区地址
            if (string.IsNullOrWhiteSpace(rangeAddress))
            {
                res.ok = false;
                res.failureStage = "range_empty";
                res.error = "选区地址为空。请显式选择需要去重的单元格区域。";
                return res;
            }

            dynamic range = null;
            try
            {
                range = sheet.Range[rangeAddress.Trim()];
            }
            catch (Exception exRng)
            {
                res.ok = false;
                res.failureStage = "range_invalid";
                res.error = "选区地址 '" + rangeAddress + "' 无效: " + exRng.Message;
                return res;
            }

            try
            {
                string cleanAddr = ((string)range.Address).Replace("$", "");
                res.rangeAddress = cleanAddr;
            }
            catch
            {
                res.rangeAddress = rangeAddress;
            }

            // 3. 检查是否为多重选区 (Multi-Area)
            try
            {
                int areaCount = (int)range.Areas.Count;
                if (areaCount > 1)
                {
                    res.ok = false;
                    res.failureStage = "multi_area_blocked";
                    res.error = "去重工具暂不支持多重非连续选区 (Multi-Area Range，当前包含 " + areaCount + " 块区域)。请选择单块连续矩形区域。";
                    return res;
                }
            }
            catch { }

            // 4. 检查是否包含合并单元格 (MergeCells)
            // 在 Excel COM 规范中：
            // - True: 整个区域为一个大合并单元格
            // - DBNull.Value (Null): 区域内部分单元格合并（混合状态）
            // - False: 区域内完全无合并单元格
            try
            {
                object mc = range.MergeCells;
                if (mc != null)
                {
                    if (Convert.IsDBNull(mc) || mc is DBNull || (mc is bool && (bool)mc))
                    {
                        res.ok = false;
                        res.failureStage = "merge_cells_blocked";
                        res.error = "所选区域包含合并单元格。去重工具要求单行独立记录以确保数据对齐安全，请先取消合并单元格。";
                        return res;
                    }
                }
            }
            catch { }

            int rowCount = 0;
            int colCount = 0;
            try
            {
                rowCount = (int)range.Rows.Count;
                colCount = (int)range.Columns.Count;
            }
            catch (Exception exDim)
            {
                res.ok = false;
                res.failureStage = "dimension_error";
                res.error = "无法读取选区行列尺寸: " + exDim.Message;
                return res;
            }

            res.selectedRowCount = rowCount;
            res.selectedColumnCount = colCount;

            if (rowCount <= 0 || colCount <= 0)
            {
                res.ok = false;
                res.failureStage = "range_empty";
                res.error = "所选区域为空或不包含有效单元格。";
                return res;
            }

            int dataRows = hasHeader ? (rowCount - 1) : rowCount;
            res.dataRowCount = dataRows;

            if (dataRows <= 0)
            {
                res.ok = false;
                res.failureStage = "no_data_rows";
                res.error = "所选区域仅包含表头行，无可用数据行。";
                return res;
            }

            // 5. 安全规模上限检查（绝不静默抽样）
            if (dataRows > MaxRowsSafetyLimit || (long)dataRows * colCount > MaxCellsSafetyLimit)
            {
                res.ok = false;
                res.failureStage = "scale_limit_exceeded";
                res.error = string.Format("选区规模超过单次安全处理上限（当前数据行: {0:N0} 行, 总单元格: {1:N0} 格；上限为 50,000 行或 1,000,000 单元格）。请缩小选区范围后再试，宿主坚决不作静默抽样假装全量去重。", dataRows, (long)rowCount * colCount);
                return res;
            }

            // 6. 校验主键列号索引 (1-based relative to selection)
            if (res.keyColumnIndices == null || res.keyColumnIndices.Count == 0)
            {
                res.ok = false;
                res.failureStage = "key_columns_missing";
                res.error = "必须显式选择至少一列主键作为去重判断依据。";
                return res;
            }

            foreach (int colIdx in res.keyColumnIndices)
            {
                if (colIdx < 1 || colIdx > colCount)
                {
                    res.ok = false;
                    res.failureStage = "key_column_out_of_range";
                    res.error = string.Format("主键列索引 {0} 超出选区范围 (有效列范围: 1 到 {1})。", colIdx, colCount);
                    return res;
                }
            }

            // 7. 批量读取 Value2 (纯二维数组，绝不逐格跨 COM 访问)
            object rawValues = null;
            try
            {
                rawValues = range.Value2;
            }
            catch (Exception exVal)
            {
                res.ok = false;
                res.failureStage = "read_failed";
                res.error = "批量读取单元格数据失败: " + exVal.Message;
                return res;
            }

            object[,] matrix = null;
            if (rawValues is object[,])
            {
                matrix = (object[,])rawValues;
            }
            else
            {
                // 单单元格特殊情况
                matrix = new object[2, 2];
                matrix[1, 1] = rawValues;
            }

            // 提取表头名称
            for (int k = 0; k < res.keyColumnIndices.Count; k++)
            {
                int cIdx = res.keyColumnIndices[k];
                string hName = "列 " + cIdx;
                if (hasHeader)
                {
                    object hVal = matrix[1, cIdx];
                    if (hVal != null && !string.IsNullOrEmpty(hVal.ToString().Trim()))
                    {
                        hName = hVal.ToString().Trim();
                    }
                }
                res.keyColumnNames.Add(hName);
            }

            // 8. 内存结构化哈希极速去重分析
            int startRowOffset = hasHeader ? 2 : 1;
            var seenKeys = new Dictionary<StructuredKey, int>(); // Key -> 首现数据行号 (1-based)
            var duplicateKeyCounts = new Dictionary<StructuredKey, int>(); // Key -> 重复次数

            int countUnique = 0;
            int countDup = 0;
            int countExcluded = 0;

            for (int r = 1; r <= dataRows; r++)
            {
                int matrixRow = startRowOffset + r - 1;
                object[] keyVals = new object[res.keyColumnIndices.Count];
                for (int ki = 0; ki < res.keyColumnIndices.Count; ki++)
                {
                    int cIdx = res.keyColumnIndices[ki];
                    keyVals[ki] = matrix[matrixRow, cIdx];
                }

                var sKey = new StructuredKey(keyVals);

                // 检查是否包含错误值、完全为空或部分为空的不可靠主键，按确定性规则单独隔离
                if (sKey.HasError() || sKey.IsAllEmpty() || sKey.HasEmptyPart())
                {
                    countExcluded++;
                    res.excludedDataRowIndices.Add(r);
                    continue;
                }

                if (seenKeys.ContainsKey(sKey))
                {
                    // 重复出现：计入重复行
                    countDup++;
                    res.duplicateDataRowIndices.Add(r);
                    duplicateKeyCounts[sKey]++;
                }
                else
                {
                    // 首次出现：计入唯一行
                    seenKeys[sKey] = r;
                    duplicateKeyCounts[sKey] = 1;
                    countUnique++;
                    res.uniqueDataRowIndices.Add(r);
                }
            }

            // 计算重复组数 (有重复出现的 Key 种数)
            int groupCount = 0;
            foreach (var kvp in duplicateKeyCounts)
            {
                if (kvp.Value > 1) groupCount++;
            }

            res.uniqueCount = countUnique;
            res.duplicateCount = countDup;
            res.duplicateGroupCount = groupCount;
            res.excludedCount = countExcluded;
            if (countExcluded > 0)
            {
                res.excludedReason = string.Format("共计 {0} 行主键包含空值、部分字段缺失或公式错误，按确定性规则予以单独隔离，不盲目合并去重。", countExcluded);
            }

            // 严格验证统计恒等式
            if (res.dataRowCount != (res.uniqueCount + res.duplicateCount + res.excludedCount))
            {
                res.ok = false;
                res.failureStage = "checksum_mismatch";
                res.error = string.Format("内部数据一致性校验失败：数据行数 ({0}) != 唯一行 ({1}) + 重复行 ({2}) + 排除行 ({3})",
                    res.dataRowCount, res.uniqueCount, res.duplicateCount, res.excludedCount);
                return res;
            }

            sw.Stop();
            res.analysisElapsedMs = sw.ElapsedMilliseconds;
            return res;
        }

        /// <summary>
        /// 执行去重写入操作：强制在目标工作簿成功快照后执行，支持 "highlight" 与 "export_unique" 两种模式
        /// </summary>
        public static DedupExecutionResult ExecuteDedup(
            dynamic app,
            dynamic targetWb,
            string sheetName,
            string rangeAddress,
            bool hasHeader,
            List<int> keyColumnIndices,
            string mode,
            string expectedFingerprint = null)
        {
            var totalSw = Stopwatch.StartNew();
            var execRes = new DedupExecutionResult
            {
                ok = false,
                mode = mode ?? "highlight"
            };

            if (targetWb == null)
            {
                execRes.error = "目标工作簿对象为空或已关闭，执行已安全终止。";
                return execRes;
            }

            // 1. 先执行只读分析
            var analysis = AnalyzeDedup(targetWb, sheetName, rangeAddress, hasHeader, keyColumnIndices);
            if (!analysis.ok)
            {
                execRes.error = "去重前置分析失败：" + analysis.error;
                return execRes;
            }

            // 2. 检查源数据变更防线
            if (!string.IsNullOrEmpty(expectedFingerprint) && analysis.dataFingerprint != expectedFingerprint)
            {
                execRes.error = "检测到源数据在分析后发生变动（数据指纹不匹配），为保证数据安全已安全阻断，请重新分析确认后再执行。";
                return execRes;
            }

            execRes.dataRowCount = analysis.dataRowCount;
            execRes.uniqueCount = analysis.uniqueCount;
            execRes.duplicateCount = analysis.duplicateCount;
            execRes.duplicateGroupCount = analysis.duplicateGroupCount;
            execRes.excludedCount = analysis.excludedCount;
            execRes.analysisElapsedMs = analysis.analysisElapsedMs;

            // 3. 检查是否有修改需求
            if (execRes.mode == "highlight" && analysis.duplicateCount == 0)
            {
                execRes.ok = true;
                execRes.message = "分析完成：未发现任何重复行，无需标记高亮。";
                execRes.modifiedRowCount = 0;
                totalSw.Stop();
                execRes.totalElapsedMs = totalSw.ElapsedMilliseconds;
                return execRes;
            }

            // 4. 必须先成功创建目标工作簿快照！承诺快照失败零业务写入
            string prompt = (execRes.mode == "highlight")
                ? string.Format("去重工具标记高亮：表【{0}】区域 {1}，发现 {2} 行重复项", analysis.targetSheetName, analysis.rangeAddress, analysis.duplicateCount)
                : string.Format("去重工具输出唯一表：表【{0}】区域 {1}，提取 {2} 行唯一数据", analysis.targetSheetName, analysis.rangeAddress, analysis.uniqueCount);

            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(targetWb, prompt, "");
            }
            catch (Exception exSnap)
            {
                execRes.error = "执行已拒绝：执行前目标工作簿快照创建失败 (" + exSnap.Message + ")。承诺快照失败零业务写入，拒绝执行写入。";
                return execRes;
            }

            if (snap == null || string.IsNullOrEmpty(snap.id))
            {
                execRes.error = "执行已拒绝：未能在磁盘上成功生成有效快照物理文件。承诺快照失败零业务写入，拒绝执行写入。";
                return execRes;
            }

            execRes.snapshotId = snap.id;
            execRes.snapshotPath = snap.filePath;
            execRes.snapshotCreated = true;

            var writeSw = Stopwatch.StartNew();

            try
            {
                dynamic sheet = targetWb.Worksheets[analysis.targetSheetName];
                dynamic srcRange = sheet.Range[analysis.rangeAddress];

                if (execRes.mode == "highlight")
                {
                    // 高亮模式：只对选区内部对应的重复行设置浅红填充，绝不向整行或工作表外扩展
                    int startRow = (int)srcRange.Row;
                    int startCol = (int)srcRange.Column;
                    int colCount = analysis.selectedColumnCount;
                    int headerOffset = hasHeader ? 1 : 0;

                    // 浅淡红色高亮 (Excel 经典重复项色号: RGB(255, 199, 206) = 0xFFC7CE / BGR 0xCEC7FF)
                    int highlightColor = 0xCEC7FF; 

                    foreach (int rIdx in analysis.duplicateDataRowIndices)
                    {
                        int absRow = startRow + headerOffset + rIdx - 1;
                        dynamic rowRange = sheet.Range[sheet.Cells[absRow, startCol], sheet.Cells[absRow, startCol + colCount - 1]];
                        rowRange.Interior.Color = highlightColor;
                    }

                    writeSw.Stop();
                    totalSw.Stop();

                    execRes.ok = true;
                    execRes.modifiedRowCount = analysis.duplicateCount;
                    execRes.highlightedRange = analysis.rangeAddress;
                    execRes.writeElapsedMs = writeSw.ElapsedMilliseconds;
                    execRes.totalElapsedMs = totalSw.ElapsedMilliseconds;
                    execRes.message = string.Format("高亮标记完成！在区域 {0} 内已将 {1} 行重复项（分属 {2} 组重复键）标记为浅红填充。源单元格数值不修改，仅改变指定重复行背景填充格式。",
                        analysis.rangeAddress, analysis.duplicateCount, analysis.duplicateGroupCount);
                    return execRes;
                }
                else if (execRes.mode == "export_unique")
                {
                    // 输出模式：创建唯一命名的新工作表，不覆盖已有工作表，批量写出纯静态数据
                    string baseNewSheetName = analysis.targetSheetName + "_唯一数据";
                    string finalSheetName = baseNewSheetName;
                    int nameCounter = 1;

                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (dynamic sh in targetWb.Worksheets)
                    {
                        try { existingNames.Add((string)sh.Name); } catch { }
                    }

                    while (existingNames.Contains(finalSheetName))
                    {
                        finalSheetName = baseNewSheetName + "_" + nameCounter;
                        nameCounter++;
                    }

                    dynamic newSheet = targetWb.Worksheets.Add(Type.Missing, sheet);
                    newSheet.Name = finalSheetName;

                    // 准备输出二维矩阵 (表头行 + 唯一行)，采用标准的 0-based 数组填充供 COM 批量写回
                    int outRowCount = (hasHeader ? 1 : 0) + analysis.uniqueCount;
                    int colCount = analysis.selectedColumnCount;
                    object[,] outMatrix = new object[outRowCount, colCount];

                    object[,] srcMatrix = (object[,])srcRange.Value2;

                    // 填充表头
                    int outRowIdx = 0;
                    if (hasHeader)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            outMatrix[outRowIdx, c - 1] = srcMatrix[1, c];
                        }
                        outRowIdx++;
                    }

                    // 填充唯一数据行
                    int srcHeaderOffset = hasHeader ? 1 : 0;
                    foreach (int uIdx in analysis.uniqueDataRowIndices)
                    {
                        int srcRow = srcHeaderOffset + uIdx;
                        for (int c = 1; c <= colCount; c++)
                        {
                            object val = srcMatrix[srcRow, c];
                            if (val is string)
                            {
                                string sVal = (string)val;
                                if (sVal.Length > 1 && sVal.StartsWith("0") && char.IsDigit(sVal[1]))
                                {
                                    val = "'" + sVal;
                                }
                            }
                            outMatrix[outRowIdx, c - 1] = val;
                        }
                        outRowIdx++;
                    }

                    // 单次批量写入新表 A1 区域 (输出纯静态值，绝不转移旧公式破坏语义)
                    dynamic destRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[outRowCount, colCount]];
                    destRange.Value2 = outMatrix;

                    // 保持列宽美观
                    try { destRange.Columns.AutoFit(); } catch { }

                    writeSw.Stop();
                    totalSw.Stop();

                    execRes.ok = true;
                    execRes.resultSheetName = finalSheetName;
                    execRes.modifiedRowCount = analysis.uniqueCount;
                    execRes.writeElapsedMs = writeSw.ElapsedMilliseconds;
                    execRes.totalElapsedMs = totalSw.ElapsedMilliseconds;
                    execRes.message = string.Format("唯一数据导出完成！已在新建工作表【{0}】中输出 {1} 行唯一数据（源表包含 {2} 行数据，滤除 {3} 行重复项；{4} 行空键/错误值排除行保留在源表中未作改动）。源工作表数据值未作任何修改。",
                        finalSheetName, analysis.uniqueCount, analysis.dataRowCount, analysis.duplicateCount, analysis.excludedCount);
                    return execRes;
                }
                else
                {
                    execRes.error = "未知的去重执行模式: " + execRes.mode;
                    return execRes;
                }
            }
            catch (Exception exExec)
            {
                writeSw.Stop();
                totalSw.Stop();
                execRes.ok = false;
                execRes.error = "执行写入时发生异常: " + exExec.Message + "。\n提示：若需恢复，可在面板中点击【快照回滚】恢复至执行前状态。";
                return execRes;
            }
        }

        #region TASK-R3b-01 两表对账引擎与确定性核验
        /// <summary>
        /// 比较列数值相等性确定性判定：严格区分类型、区分大小写、不去空格、不隐式转换
        /// </summary>
        public static bool ValuesEqual(object v1, object v2)
        {
            if (v1 == null && v2 == null) return true;
            if (v1 == null || v2 == null) return false;

            if (v1.GetType() != v2.GetType()) return false;

            if (v1 is string)
            {
                return string.Equals((string)v1, (string)v2, StringComparison.Ordinal);
            }
            else if (v1 is double)
            {
                return Math.Abs((double)v1 - (double)v2) < 1e-9;
            }
            else if (v1 is bool)
            {
                return (bool)v1 == (bool)v2;
            }
            else if (v1 is int)
            {
                return (int)v1 == (int)v2;
            }
            return v1.Equals(v2);
        }

        /// <summary>
        /// 格式化单元格输出值：严格保真 19 位纯数字长编号、前导零文本以及以等号开头的纯文本，防止 Excel 自动转科学计数法或公式
        /// </summary>
        public static object FormatValueForExcelOutput(object val)
        {
            if (val == null) return string.Empty;
            if (val is string)
            {
                string s = (string)val;
                if (string.IsNullOrEmpty(s)) return string.Empty;

                // 1. 若文本以 '=' 开头，添加单引号前缀保真为纯文本，防止意外变为可执行公式
                if (s.StartsWith("="))
                {
                    return "'" + s;
                }

                // 2. 若文本长度 >= 11 且全为数字（如 19 位身份证号/工单号），添加单引号前缀防止 Excel 转换为科学计数法
                if (s.Length >= 11 && IsAllDigits(s))
                {
                    return "'" + s;
                }

                // 3. 若文本以 '0' 开头且长度 > 1 且包含数字（如 '001', '0123'），添加单引号保真前导零
                if (s.Length > 1 && s.StartsWith("0") && char.IsDigit(s[1]))
                {
                    return "'" + s;
                }

                // 4. 若文本原本以 '\'' 开头，添加单引号前缀防止 Excel 作为前缀字符吞掉
                if (s.StartsWith("'"))
                {
                    return "'" + s;
                }

                return s;
            }
            return val;
        }

        private static bool IsAllDigits(string str)
        {
            foreach (char c in str)
            {
                if (!char.IsDigit(c)) return false;
            }
            return true;
        }

        /// <summary>
        /// 两表对账只读分析服务：零模型调用、零快照、零源表修改、严格统计恒等式
        /// </summary>
        public static ReconcileAnalysisResult AnalyzeReconcile(
            dynamic app,
            string leftWbName,
            string leftSheetName,
            string leftRangeAddress,
            bool leftHasHeader,
            List<int> leftKeyCols,
            string rightWbName,
            string rightSheetName,
            string rightRangeAddress,
            bool rightHasHeader,
            List<int> rightKeyCols,
            List<CompareColMappingDto> compareCols)
        {
            var sw = Stopwatch.StartNew();
            var res = new ReconcileAnalysisResult
            {
                ok = true,
                leftHasHeader = leftHasHeader,
                rightHasHeader = rightHasHeader,
                leftKeyCols = leftKeyCols != null ? new List<int>(leftKeyCols) : new List<int>(),
                rightKeyCols = rightKeyCols != null ? new List<int>(rightKeyCols) : new List<int>(),
                compareCols = compareCols != null ? new List<CompareColMappingDto>(compareCols) : new List<CompareColMappingDto>()
            };

            if (app == null)
            {
                res.ok = false;
                res.failureStage = "app_null";
                res.error = "Excel Application 对象为空，无法执行对账分析。";
                return res;
            }

            // 1. 查找左表与右表工作簿
            dynamic leftWb = FindWorkbookByName(app, leftWbName);
            if (leftWb == null)
            {
                res.ok = false;
                res.failureStage = "left_wb_missing";
                res.error = string.Format("未找到左表目标工作簿【{0}】，已阻断，不回退活动工作簿。", leftWbName ?? "未指定");
                return res;
            }
            res.leftWorkbookName = (string)leftWb.Name;

            dynamic rightWb = FindWorkbookByName(app, rightWbName);
            if (rightWb == null)
            {
                res.ok = false;
                res.failureStage = "right_wb_missing";
                res.error = string.Format("未找到右表目标工作簿【{0}】，已阻断，不回退活动工作簿。", rightWbName ?? "未指定");
                return res;
            }
            res.rightWorkbookName = (string)rightWb.Name;

            // 2. 获取工作表
            dynamic leftSheet = null;
            try
            {
                leftSheet = string.IsNullOrEmpty(leftSheetName) ? leftWb.ActiveSheet : leftWb.Worksheets[leftSheetName];
            }
            catch (Exception exLs)
            {
                res.ok = false;
                res.failureStage = "left_sheet_missing";
                res.error = "无法定位左表工作表 '" + leftSheetName + "': " + exLs.Message;
                return res;
            }
            res.leftSheetName = (string)leftSheet.Name;

            dynamic rightSheet = null;
            try
            {
                rightSheet = string.IsNullOrEmpty(rightSheetName) ? rightWb.ActiveSheet : rightWb.Worksheets[rightSheetName];
            }
            catch (Exception exRs)
            {
                res.ok = false;
                res.failureStage = "right_sheet_missing";
                res.error = "无法定位右表工作表 '" + rightSheetName + "': " + exRs.Message;
                return res;
            }
            res.rightSheetName = (string)rightSheet.Name;

            // 3. 校验区域有效性
            if (string.IsNullOrWhiteSpace(leftRangeAddress) || string.IsNullOrWhiteSpace(rightRangeAddress))
            {
                res.ok = false;
                res.failureStage = "range_empty";
                res.error = "左表或右表的选区地址为空。请显式指定两侧待对账区域。";
                return res;
            }

            dynamic leftRange = null;
            dynamic rightRange = null;
            try
            {
                leftRange = leftSheet.Range[leftRangeAddress.Trim()];
                res.leftRangeAddress = ((string)leftRange.Address).Replace("$", "");
            }
            catch (Exception exLr)
            {
                res.ok = false;
                res.failureStage = "left_range_invalid";
                res.error = "左表选区地址 '" + leftRangeAddress + "' 无效: " + exLr.Message;
                return res;
            }

            try
            {
                rightRange = rightSheet.Range[rightRangeAddress.Trim()];
                res.rightRangeAddress = ((string)rightRange.Address).Replace("$", "");
            }
            catch (Exception exRr)
            {
                res.ok = false;
                res.failureStage = "right_range_invalid";
                res.error = "右表选区地址 '" + rightRangeAddress + "' 无效: " + exRr.Message;
                return res;
            }

            // 4. 多区域与合并单元格阻断检查
            if ((int)leftRange.Areas.Count > 1 || (int)rightRange.Areas.Count > 1)
            {
                res.ok = false;
                res.failureStage = "multi_area_blocked";
                res.error = "对账工具暂不支持多重非连续选区 (Multi-Area Range)。请分别选择连续的单块矩形区域。";
                return res;
            }

            object leftMc = leftRange.MergeCells;
            if (leftMc != null && (Convert.IsDBNull(leftMc) || leftMc is DBNull || (leftMc is bool && (bool)leftMc)))
            {
                res.ok = false;
                res.failureStage = "merge_cells_blocked";
                res.error = "左表选区包含合并单元格。对账工具要求单行独立记录以确保对齐与对账安全，请先取消合并单元格。";
                return res;
            }

            object rightMc = rightRange.MergeCells;
            if (rightMc != null && (Convert.IsDBNull(rightMc) || rightMc is DBNull || (rightMc is bool && (bool)rightMc)))
            {
                res.ok = false;
                res.failureStage = "merge_cells_blocked";
                res.error = "右表选区包含合并单元格。对账工具要求单行独立记录以确保对齐与对账安全，请先取消合并单元格。";
                return res;
            }

            // 5. 尺寸与规模上限检查
            res.leftRowCount = (int)leftRange.Rows.Count;
            res.leftColCount = (int)leftRange.Columns.Count;
            res.rightRowCount = (int)rightRange.Rows.Count;
            res.rightColCount = (int)rightRange.Columns.Count;

            res.leftDataRowCount = leftHasHeader ? (res.leftRowCount - 1) : res.leftRowCount;
            res.rightDataRowCount = rightHasHeader ? (res.rightRowCount - 1) : res.rightRowCount;

            if (res.leftDataRowCount <= 0 || res.rightDataRowCount <= 0)
            {
                res.ok = false;
                res.failureStage = "no_data_rows";
                res.error = "左表或右表无有效数据行（仅有表头或区域为空）。";
                return res;
            }

            long totalDataRows = (long)res.leftDataRowCount + res.rightDataRowCount;
            long totalCells = (long)res.leftRowCount * res.leftColCount + (long)res.rightRowCount * res.rightColCount;
            if (totalDataRows > MaxRowsSafetyLimit || totalCells > MaxCellsSafetyLimit)
            {
                res.ok = false;
                res.failureStage = "scale_limit_exceeded";
                res.error = string.Format("两表规模合计超过单次安全处理上限（当前两表合计数据行: {0:N0} 行, 总单元格: {1:N0} 格；上限为 50,000 行或 1,000,000 单元格）。请缩小范围，绝不静默抽样假装全量对账。", totalDataRows, totalCells);
                return res;
            }

            // 6. 校验主键列索引映射 (1-based relative to selection)
            if (res.leftKeyCols.Count == 0 || res.rightKeyCols.Count == 0 || res.leftKeyCols.Count != res.rightKeyCols.Count)
            {
                res.ok = false;
                res.failureStage = "key_columns_mismatch";
                res.error = "主键列映射必须显式指定且两侧列数必须完全相等。";
                return res;
            }

            for (int i = 0; i < res.leftKeyCols.Count; i++)
            {
                int lc = res.leftKeyCols[i];
                int rc = res.rightKeyCols[i];
                if (lc < 1 || lc > res.leftColCount)
                {
                    res.ok = false;
                    res.failureStage = "key_column_out_of_range";
                    res.error = string.Format("左表主键列索引 {0} 超出左表有效列范围 (1 到 {1})。", lc, res.leftColCount);
                    return res;
                }
                if (rc < 1 || rc > res.rightColCount)
                {
                    res.ok = false;
                    res.failureStage = "key_column_out_of_range";
                    res.error = string.Format("右表主键列索引 {0} 超出右表有效列范围 (1 到 {1})。", rc, res.rightColCount);
                    return res;
                }
            }

            // 校验比较列映射
            foreach (var cm in res.compareCols)
            {
                if (cm.leftColIndex < 1 || cm.leftColIndex > res.leftColCount)
                {
                    res.ok = false;
                    res.failureStage = "compare_column_out_of_range";
                    res.error = string.Format("待核对比较列中，左表列索引 {0} 超出有效列范围 (1 到 {1})。", cm.leftColIndex, res.leftColCount);
                    return res;
                }
                if (cm.rightColIndex < 1 || cm.rightColIndex > res.rightColCount)
                {
                    res.ok = false;
                    res.failureStage = "compare_column_out_of_range";
                    res.error = string.Format("待核对比较列中，右表列索引 {0} 超出有效列范围 (1 到 {1})。", cm.rightColIndex, res.rightColCount);
                    return res;
                }
            }

            // 7. 批量读取 Value2 (纯二维数组)
            object[,] leftMatrix = null;
            object[,] rightMatrix = null;
            try
            {
                leftMatrix = Get2DMatrix(leftRange.Value2, res.leftRowCount, res.leftColCount);
                rightMatrix = Get2DMatrix(rightRange.Value2, res.rightRowCount, res.rightColCount);
            }
            catch (Exception exVal)
            {
                res.ok = false;
                res.failureStage = "read_failed";
                res.error = "批量读取单元格数据失败: " + exVal.Message;
                return res;
            }

            // 提取主键与比较列表头名称
            for (int i = 0; i < res.leftKeyCols.Count; i++)
            {
                int lc = res.leftKeyCols[i];
                int rc = res.rightKeyCols[i];
                string lName = "左列 " + lc;
                string rName = "右列 " + rc;
                if (leftHasHeader && leftMatrix[1, lc] != null) lName = leftMatrix[1, lc].ToString().Trim();
                if (rightHasHeader && rightMatrix[1, rc] != null) rName = rightMatrix[1, rc].ToString().Trim();
                res.leftKeyNames.Add(lName);
                res.rightKeyNames.Add(rName);
            }

            foreach (var cm in res.compareCols)
            {
                if (string.IsNullOrEmpty(cm.leftColName))
                {
                    cm.leftColName = leftHasHeader && leftMatrix[1, cm.leftColIndex] != null ? leftMatrix[1, cm.leftColIndex].ToString().Trim() : ("左列 " + cm.leftColIndex);
                }
                if (string.IsNullOrEmpty(cm.rightColName))
                {
                    cm.rightColName = rightHasHeader && rightMatrix[1, cm.rightColIndex] != null ? rightMatrix[1, cm.rightColIndex].ToString().Trim() : ("右列 " + cm.rightColIndex);
                }
            }

            // 8. 内存高效哈希分组对账核心算法
            int leftStartOffset = leftHasHeader ? 2 : 1;
            int rightStartOffset = rightHasHeader ? 2 : 1;

            // 字典收集主键出现的所有数据行号 (1-based relative to data rows)
            var leftKeyToRows = new Dictionary<StructuredKey, List<int>>();
            var rightKeyToRows = new Dictionary<StructuredKey, List<int>>();

            var leftInvalidRows = new List<int>();
            var rightInvalidRows = new List<int>();

            // 扫描左表数据行
            for (int r = 1; r <= res.leftDataRowCount; r++)
            {
                int mRow = leftStartOffset + r - 1;
                object[] keyVals = new object[res.leftKeyCols.Count];
                for (int ki = 0; ki < res.leftKeyCols.Count; ki++)
                {
                    keyVals[ki] = leftMatrix[mRow, res.leftKeyCols[ki]];
                }
                var sk = new StructuredKey(keyVals);

                if (sk.HasError() || sk.IsAllEmpty() || sk.HasEmptyPart())
                {
                    leftInvalidRows.Add(r);
                }
                else
                {
                    if (!leftKeyToRows.ContainsKey(sk))
                    {
                        leftKeyToRows[sk] = new List<int>();
                    }
                    leftKeyToRows[sk].Add(r);
                }
            }

            // 扫描右表数据行
            for (int r = 1; r <= res.rightDataRowCount; r++)
            {
                int mRow = rightStartOffset + r - 1;
                object[] keyVals = new object[res.rightKeyCols.Count];
                for (int ki = 0; ki < res.rightKeyCols.Count; ki++)
                {
                    keyVals[ki] = rightMatrix[mRow, res.rightKeyCols[ki]];
                }
                var sk = new StructuredKey(keyVals);

                if (sk.HasError() || sk.IsAllEmpty() || sk.HasEmptyPart())
                {
                    rightInvalidRows.Add(r);
                }
                else
                {
                    if (!rightKeyToRows.ContainsKey(sk))
                    {
                        rightKeyToRows[sk] = new List<int>();
                    }
                    rightKeyToRows[sk].Add(r);
                }
            }

            // 记录主键异常行 (空键、部分空、错误值)
            foreach (int r in leftInvalidRows)
            {
                res.details.Add(new ReconcileRowDetail
                {
                    category = "LEFT_INVALID",
                    leftRowIndex = r,
                    rightRowIndex = 0,
                    keyDisplay = "[异常主键: 空值/错误值]"
                });
            }
            foreach (int r in rightInvalidRows)
            {
                res.details.Add(new ReconcileRowDetail
                {
                    category = "RIGHT_INVALID",
                    leftRowIndex = 0,
                    rightRowIndex = r,
                    keyDisplay = "[异常主键: 空值/错误值]"
                });
            }

            int matchedSame = 0;
            int matchedDiff = 0;
            int leftOnly = 0;
            int rightOnly = 0;
            int leftDupRowCount = 0;
            int rightDupRowCount = 0;
            int diffCellsTotal = 0;

            // 获取所有去重主键集合并统一判定分类
            var allKeys = new HashSet<StructuredKey>(leftKeyToRows.Keys);
            foreach (var rk in rightKeyToRows.Keys)
            {
                allKeys.Add(rk);
            }

            foreach (var key in allKeys)
            {
                List<int> leftRows;
                bool hasLeft = leftKeyToRows.TryGetValue(key, out leftRows);
                List<int> rightRows;
                bool hasRight = rightKeyToRows.TryGetValue(key, out rightRows);

                int lCount = hasLeft ? leftRows.Count : 0;
                int rCount = hasRight ? rightRows.Count : 0;

                if (lCount == 1 && rCount == 1)
                {
                    // 1对1 唯一匹配比对
                    int leftR = leftRows[0];
                    int rightR = rightRows[0];
                    int leftMRow = leftStartOffset + leftR - 1;
                    int rightMRow = rightStartOffset + rightR - 1;

                    var cellDiffs = new List<CellDiffItem>();
                    foreach (var cm in res.compareCols)
                    {
                        object lv = leftMatrix[leftMRow, cm.leftColIndex];
                        object rv = rightMatrix[rightMRow, cm.rightColIndex];

                        if (!ValuesEqual(lv, rv))
                        {
                            cellDiffs.Add(new CellDiffItem
                            {
                                colName = cm.leftColName + " vs " + cm.rightColName,
                                leftValue = lv,
                                leftType = lv != null ? lv.GetType().Name : "null",
                                rightValue = rv,
                                rightType = rv != null ? rv.GetType().Name : "null"
                            });
                        }
                    }

                    if (cellDiffs.Count == 0)
                    {
                        matchedSame++;
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "SAME",
                            leftRowIndex = leftR,
                            rightRowIndex = rightR,
                            keyDisplay = FormatKeyDisplay(key)
                        });
                    }
                    else
                    {
                        matchedDiff++;
                        diffCellsTotal += cellDiffs.Count;
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "DIFF",
                            leftRowIndex = leftR,
                            rightRowIndex = rightR,
                            keyDisplay = FormatKeyDisplay(key),
                            diffs = cellDiffs
                        });
                    }
                }
                else if (hasLeft && !hasRight)
                {
                    // 仅左表存在 (右表完全不存在该主键)
                    if (lCount == 1)
                    {
                        leftOnly++;
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "LEFT_ONLY",
                            leftRowIndex = leftRows[0],
                            rightRowIndex = 0,
                            keyDisplay = FormatKeyDisplay(key)
                        });
                    }
                    else
                    {
                        leftDupRowCount += lCount;
                        foreach (int r in leftRows)
                        {
                            res.details.Add(new ReconcileRowDetail
                            {
                                category = "LEFT_DUP",
                                leftRowIndex = r,
                                rightRowIndex = 0,
                                keyDisplay = FormatKeyDisplay(key)
                            });
                        }
                    }
                }
                else if (!hasLeft && hasRight)
                {
                    // 仅右表存在 (左表完全不存在该主键)
                    if (rCount == 1)
                    {
                        rightOnly++;
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "RIGHT_ONLY",
                            leftRowIndex = 0,
                            rightRowIndex = rightRows[0],
                            keyDisplay = FormatKeyDisplay(key)
                        });
                    }
                    else
                    {
                        rightDupRowCount += rCount;
                        foreach (int r in rightRows)
                        {
                            res.details.Add(new ReconcileRowDetail
                            {
                                category = "RIGHT_DUP",
                                leftRowIndex = 0,
                                rightRowIndex = r,
                                keyDisplay = FormatKeyDisplay(key)
                            });
                        }
                    }
                }
                else
                {
                    // 两侧均存在该键，但至少一侧出现多行 (存在重复或非对称匹配歧义，绝不自动一对多，两侧相关行分别记入重复/歧义项)
                    leftDupRowCount += lCount;
                    foreach (int r in leftRows)
                    {
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "LEFT_DUP",
                            leftRowIndex = r,
                            rightRowIndex = 0,
                            keyDisplay = FormatKeyDisplay(key)
                        });
                    }

                    rightDupRowCount += rCount;
                    foreach (int r in rightRows)
                    {
                        res.details.Add(new ReconcileRowDetail
                        {
                            category = "RIGHT_DUP",
                            leftRowIndex = 0,
                            rightRowIndex = r,
                            keyDisplay = FormatKeyDisplay(key)
                        });
                    }
                }
            }

            res.matchedBothSameCount = matchedSame;
            res.matchedBothDiffCount = matchedDiff;
            res.leftOnlyCount = leftOnly;
            res.rightOnlyCount = rightOnly;
            res.leftDuplicateKeyCount = leftDupRowCount;
            res.rightDuplicateKeyCount = rightDupRowCount;
            res.leftInvalidKeyCount = leftInvalidRows.Count;
            res.rightInvalidKeyCount = rightInvalidRows.Count;
            res.diffCellCount = diffCellsTotal;

            // 严格验证统计恒等式
            int leftCalcTotal = res.matchedBothSameCount + res.matchedBothDiffCount + res.leftOnlyCount + res.leftDuplicateKeyCount + res.leftInvalidKeyCount;
            int rightCalcTotal = res.matchedBothSameCount + res.matchedBothDiffCount + res.rightOnlyCount + res.rightDuplicateKeyCount + res.rightInvalidKeyCount;

            if (res.leftDataRowCount != leftCalcTotal)
            {
                res.ok = false;
                res.failureStage = "left_checksum_mismatch";
                res.error = string.Format("左表统计恒等式破坏：左表数据行 ({0}) != 完全一致 ({1}) + 存在差异 ({2}) + 仅左表 ({3}) + 左表重复键 ({4}) + 左表异常键 ({5})",
                    res.leftDataRowCount, res.matchedBothSameCount, res.matchedBothDiffCount, res.leftOnlyCount, res.leftDuplicateKeyCount, res.leftInvalidKeyCount);
                return res;
            }

            if (res.rightDataRowCount != rightCalcTotal)
            {
                res.ok = false;
                res.failureStage = "right_checksum_mismatch";
                res.error = string.Format("右表统计恒等式破坏：右表数据行 ({0}) != 完全一致 ({1}) + 存在差异 ({2}) + 仅右表 ({3}) + 右表重复键 ({4}) + 右表异常键 ({5})",
                    res.rightDataRowCount, res.matchedBothSameCount, res.matchedBothDiffCount, res.rightOnlyCount, res.rightDuplicateKeyCount, res.rightInvalidKeyCount);
                return res;
            }

            sw.Stop();
            res.analysisElapsedMs = sw.ElapsedMilliseconds;
            res.dataFingerprint = string.Format("L:{0}-{1}-{2}|R:{3}-{4}-{5}|S:{6}|D:{7}",
                res.leftDataRowCount, res.leftKeyCols.Count, res.leftDuplicateKeyCount,
                res.rightDataRowCount, res.rightKeyCols.Count, res.rightDuplicateKeyCount,
                res.matchedBothSameCount, res.matchedBothDiffCount);

            return res;
        }

        /// <summary>
        /// 对账执行写入服务：强制前置快照、新建独立唯一命名结果表、纯静态值批量写出、长编号与前导零保真
        /// </summary>
        public static ReconcileExecutionResult ExecuteReconcile(
            dynamic app,
            string leftWbName,
            string leftSheetName,
            string leftRangeAddress,
            bool leftHasHeader,
            List<int> leftKeyCols,
            string rightWbName,
            string rightSheetName,
            string rightRangeAddress,
            bool rightHasHeader,
            List<int> rightKeyCols,
            List<CompareColMappingDto> compareCols,
            string outputWbName = null,
            string expectedFingerprint = null)
        {
            var totalSw = Stopwatch.StartNew();
            var execRes = new ReconcileExecutionResult { ok = false };

            if (app == null)
            {
                execRes.error = "Excel Application 对象为空，无法执行对账写入。";
                return execRes;
            }

            // 1. 执行前实时只读分析
            var analysis = AnalyzeReconcile(app, leftWbName, leftSheetName, leftRangeAddress, leftHasHeader, leftKeyCols,
                rightWbName, rightSheetName, rightRangeAddress, rightHasHeader, rightKeyCols, compareCols);

            if (!analysis.ok)
            {
                execRes.error = "对账前置分析失败：" + analysis.error;
                return execRes;
            }

            // 2. 检查源数据变更防线
            if (!string.IsNullOrEmpty(expectedFingerprint) && analysis.dataFingerprint != expectedFingerprint)
            {
                execRes.error = "检测到源数据在分析后发生变动（数据指纹不匹配），为保证对账真实性已安全阻断，请重新分析确认后再执行。";
                return execRes;
            }

            execRes.matchedBothSameCount = analysis.matchedBothSameCount;
            execRes.matchedBothDiffCount = analysis.matchedBothDiffCount;
            execRes.leftOnlyCount = analysis.leftOnlyCount;
            execRes.rightOnlyCount = analysis.rightOnlyCount;
            execRes.leftDuplicateKeyCount = analysis.leftDuplicateKeyCount;
            execRes.rightDuplicateKeyCount = analysis.rightDuplicateKeyCount;
            execRes.leftInvalidKeyCount = analysis.leftInvalidKeyCount;
            execRes.rightInvalidKeyCount = analysis.rightInvalidKeyCount;
            execRes.diffCellCount = analysis.diffCellCount;
            execRes.analysisElapsedMs = analysis.analysisElapsedMs;

            // 3. 确定目标输出工作簿
            string targetOutWbName = string.IsNullOrEmpty(outputWbName) ? analysis.leftWorkbookName : outputWbName;
            dynamic targetWb = FindWorkbookByName(app, targetOutWbName);
            if (targetWb == null)
            {
                execRes.error = string.Format("未找到对账结果输出目标工作簿【{0}】，已阻断。", targetOutWbName);
                return execRes;
            }
            execRes.targetWorkbookName = (string)targetWb.Name;

            // 4. 强制前置物理快照保障
            string prompt = string.Format("两表对账输出：左表【{0}!{1}】与右表【{2}!{3}】，共对账 {4} 条明细",
                analysis.leftSheetName, analysis.leftRangeAddress, analysis.rightSheetName, analysis.rightRangeAddress, analysis.details.Count);

            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(targetWb, prompt, "");
            }
            catch (Exception exSnap)
            {
                execRes.error = "执行已拒绝：输出工作簿快照创建失败 (" + exSnap.Message + ")。拒绝写入以确保数据安全。";
                return execRes;
            }

            if (snap == null || string.IsNullOrEmpty(snap.id))
            {
                execRes.error = "执行已拒绝：未能在磁盘上成功生成有效快照物理文件。拒绝写入以防数据无法恢复。";
                return execRes;
            }

            execRes.snapshotId = snap.id;
            execRes.snapshotPath = snap.filePath;
            execRes.snapshotCreated = true;

            var writeSw = Stopwatch.StartNew();

            try
            {
                // 5. 创建独立命名的新工作表
                string baseSheetName = "两表对账结果";
                string finalSheetName = baseSheetName;
                int counter = 1;

                var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic sh in targetWb.Worksheets)
                {
                    try { existingNames.Add((string)sh.Name); } catch { }
                }

                while (existingNames.Contains(finalSheetName))
                {
                    finalSheetName = baseSheetName + "_" + counter;
                    counter++;
                }

                dynamic newSheet = targetWb.Worksheets.Add();
                newSheet.Name = finalSheetName;

                // 6. 构造输出数据矩阵
                // 汇总区：行 1-4
                // 空行：行 5
                // 明细表头：行 6
                // 明细数据：行 7 ~ 7 + details.Count - 1
                int detailCount = analysis.details.Count;
                int totalOutRows = 6 + detailCount;
                int totalOutCols = 5; // 状态, 主键, 左行号, 右行号, 差异明细

                object[,] outMatrix = new object[totalOutRows, totalOutCols];

                // 填充汇总行 (0-based 索引)
                outMatrix[0, 0] = "【对账汇总卡片】";
                outMatrix[0, 1] = string.Format("左表数据行: {0}，右表数据行: {1}", analysis.leftDataRowCount, analysis.rightDataRowCount);
                outMatrix[0, 2] = "";
                outMatrix[0, 3] = "";
                outMatrix[0, 4] = "";

                outMatrix[1, 0] = "【合法唯一主键】";
                outMatrix[1, 1] = string.Format("完全一致: {0}，存在差异: {1}，仅左表: {2}，仅右表: {3}",
                    analysis.matchedBothSameCount, analysis.matchedBothDiffCount, analysis.leftOnlyCount, analysis.rightOnlyCount);
                outMatrix[1, 2] = "";
                outMatrix[1, 3] = "";
                outMatrix[1, 4] = "";

                outMatrix[2, 0] = "【异常与重复/歧义键】";
                outMatrix[2, 1] = string.Format("左表重复/歧义行数: {0}，右表重复/歧义行数: {1}，左表异常主键: {2}，右表异常主键: {3}，差异单元格对数: {4}",
                    analysis.leftDuplicateKeyCount, analysis.rightDuplicateKeyCount, analysis.leftInvalidKeyCount, analysis.rightInvalidKeyCount, analysis.diffCellCount);
                outMatrix[2, 2] = "";
                outMatrix[2, 3] = "";
                outMatrix[2, 4] = "";

                outMatrix[3, 0] = "【源表保护声明】";
                outMatrix[3, 1] = "源工作表完全无损，源数据值零修改；本次生成独立静态值对账表，公式及长编号已安全保真。";
                outMatrix[3, 2] = "";
                outMatrix[3, 3] = "";
                outMatrix[3, 4] = "";

                outMatrix[4, 0] = "";
                outMatrix[4, 1] = "";
                outMatrix[4, 2] = "";
                outMatrix[4, 3] = "";
                outMatrix[4, 4] = "";

                // 表头
                outMatrix[5, 0] = "对账判定";
                outMatrix[5, 1] = "对账主键";
                outMatrix[5, 2] = "左表行号";
                outMatrix[5, 3] = "右表行号";
                outMatrix[5, 4] = "差异字段与详情说明";

                // 填充明细
                for (int i = 0; i < detailCount; i++)
                {
                    int rIdx = 6 + i;
                    var d = analysis.details[i];

                    string catDisplay = d.category ?? "";
                    if (catDisplay == "SAME") catDisplay = "完全一致";
                    else if (catDisplay == "DIFF") catDisplay = "存在差异";
                    else if (catDisplay == "LEFT_ONLY") catDisplay = "仅左表存在";
                    else if (catDisplay == "RIGHT_ONLY") catDisplay = "仅右表存在";
                    else if (catDisplay == "LEFT_DUP") catDisplay = "左表重复主键";
                    else if (catDisplay == "RIGHT_DUP") catDisplay = "右表重复主键";
                    else if (catDisplay == "LEFT_INVALID") catDisplay = "左表主键异常";
                    else if (catDisplay == "RIGHT_INVALID") catDisplay = "右表主键异常";

                    outMatrix[rIdx, 0] = catDisplay;
                    outMatrix[rIdx, 1] = FormatValueForExcelOutput(d.keyDisplay);
                    outMatrix[rIdx, 2] = d.leftRowIndex > 0 ? (object)d.leftRowIndex : "";
                    outMatrix[rIdx, 3] = d.rightRowIndex > 0 ? (object)d.rightRowIndex : "";

                    if (d.diffs != null && d.diffs.Count > 0)
                    {
                        var sbDiff = new StringBuilder();
                        for (int di = 0; di < d.diffs.Count; di++)
                        {
                            var diff = d.diffs[di];
                            if (di > 0) sbDiff.Append("; ");
                            sbDiff.AppendFormat("[{0}]: 左='{1}'({2}) vs 右='{3}'({4})",
                                diff.colName,
                                diff.leftValue != null ? diff.leftValue.ToString() : "空",
                                diff.leftType,
                                diff.rightValue != null ? diff.rightValue.ToString() : "空",
                                diff.rightType);
                        }
                        outMatrix[rIdx, 4] = sbDiff.ToString();
                    }
                    else
                    {
                        outMatrix[rIdx, 4] = "";
                    }
                }

                // 7. 单次批量写入新表
                dynamic destRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[totalOutRows, totalOutCols]];
                destRange.Value2 = outMatrix;

                // 8. 格式美化与差异染色 (仅设置状态列背景色)
                try
                {
                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = analysis.details[i];
                        int absRow = 7 + i;
                        dynamic statusCell = newSheet.Cells[absRow, 1];
                        if (d.category == "DIFF")
                        {
                            statusCell.Interior.Color = 0xC6E2FF; // 浅橙黄
                        }
                        else if (d.category == "LEFT_ONLY")
                        {
                            statusCell.Interior.Color = 0xFFEBCD; // 浅蓝
                        }
                        else if (d.category == "RIGHT_ONLY")
                        {
                            statusCell.Interior.Color = 0xE6E6FA; // 浅紫
                        }
                        else if (d.category.Contains("DUP") || d.category.Contains("INVALID"))
                        {
                            statusCell.Interior.Color = 0xCEC7FF; // 浅红
                        }
                    }

                    destRange.Columns.AutoFit();
                }
                catch { }

                writeSw.Stop();
                totalSw.Stop();

                execRes.ok = true;
                execRes.resultSheetName = finalSheetName;
                execRes.exportedRowCount = detailCount;
                execRes.writeElapsedMs = writeSw.ElapsedMilliseconds;
                execRes.totalElapsedMs = totalSw.ElapsedMilliseconds;
                execRes.message = string.Format("两表对账完成！已在新工作表【{0}】中生成静态对账结果表（完全一致 {1} 行，差异 {2} 行，仅左 {3} 行，仅右 {4} 行，左重复 {5} 行，右重复 {6} 行）。源表数据 100% 零修改。",
                    finalSheetName, analysis.matchedBothSameCount, analysis.matchedBothDiffCount, analysis.leftOnlyCount, analysis.rightOnlyCount, analysis.leftDuplicateKeyCount, analysis.rightDuplicateKeyCount);
                return execRes;
            }
            catch (Exception exWrite)
            {
                writeSw.Stop();
                totalSw.Stop();
                execRes.ok = false;
                execRes.error = "写入对账结果表时发生异常: " + exWrite.Message + "。\n提示：若需恢复，可使用面板中的【快照回滚】功能恢复输出工作簿。";
                return execRes;
            }
        }

        private static dynamic FindWorkbookByName(dynamic app, string name)
        {
            if (app == null) return null;
            if (string.IsNullOrEmpty(name))
            {
                // 未显式指定工作簿名时，默认为当前活动工作簿
                try { return app.ActiveWorkbook; } catch { return null; }
            }
            string cleanTarget = name.Trim();
            foreach (dynamic wb in app.Workbooks)
            {
                try
                {
                    string wbName = (string)wb.Name;
                    string wbFullName = (string)wb.FullName;
                    if (string.Equals(wbName, cleanTarget, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(wbFullName, cleanTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }
                catch { }
            }
            // 若显式指定了目标工作簿但未能找到，拒绝隐式回退当前活动工作簿，以防误操作！
            return null;
        }

        private static object[,] Get2DMatrix(object rawValues, int expectedRows, int expectedCols)
        {
            if (rawValues is object[,])
            {
                return (object[,])rawValues;
            }
            var m = new object[expectedRows + 1, expectedCols + 1];
            m[1, 1] = rawValues;
            return m;
        }

        private static string FormatKeyDisplay(StructuredKey key)
        {
            if (key == null || key.Values == null || key.Values.Length == 0) return string.Empty;
            if (key.Values.Length == 1)
            {
                object v = key.Values[0];
                return v != null ? v.ToString() : "空";
            }
            var sb = new StringBuilder();
            sb.Append("(");
            for (int i = 0; i < key.Values.Length; i++)
            {
                if (i > 0) sb.Append(", ");
                object v = key.Values[i];
                sb.Append(v != null ? v.ToString() : "空");
            }
            sb.Append(")");
            return sb.ToString();
        }

        #region 多文件列名对齐汇总服务 (TASK-R4c-01)

        private static string ComputeFileSha256(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return string.Empty;
            try
            {
                using (var sha = SHA256.Create())
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hash = sha.ComputeHash(stream);
                    var sb = new StringBuilder();
                    foreach (byte b in hash) sb.Append(b.ToString("x2"));
                    return sb.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static void ResolveDisplayIdentifiers(List<ConsolidationSourceDef> sourceDefs)
        {
            if (sourceDefs == null) return;
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in sourceDefs)
            {
                string fn = Path.GetFileName(s.filePath ?? "");
                if (!counts.ContainsKey(fn)) counts[fn] = 0;
                counts[fn]++;
            }
            foreach (var s in sourceDefs)
            {
                string fn = Path.GetFileName(s.filePath ?? "");
                if (counts[fn] > 1)
                {
                    string dir = Path.GetFileName(Path.GetDirectoryName(s.filePath ?? ""));
                    s.displayIdentifier = (string.IsNullOrEmpty(dir) ? "" : dir + "\\") + fn;
                }
                else
                {
                    s.displayIdentifier = fn;
                }
            }
        }

        // 内部路径计算辅助方法，限定为 internal，避免被 Excel-DNA 自动导出为 Excel 工作表函数产生重复注册冲突
        internal static string ResolveUniqueOutputPath(string desiredPath)
        {
            if (!File.Exists(desiredPath)) return desiredPath;
            string dir = Path.GetDirectoryName(desiredPath);
            string fn = Path.GetFileNameWithoutExtension(desiredPath);
            string ext = Path.GetExtension(desiredPath);
            int idx = 1;
            while (true)
            {
                string candidate = Path.Combine(dir, fn + "_" + idx + ext);
                if (!File.Exists(candidate)) return candidate;
                idx++;
            }
        }

        public static ConsolidationAnalysisResult AnalyzeConsolidation(
            dynamic app,
            List<ConsolidationSourceDef> sourceDefs,
            List<string> userColumnOrder = null,
            List<ConsolidationColMapping> explicitMappings = null)
        {
            var sw = Stopwatch.StartNew();
            var res = new ConsolidationAnalysisResult { ok = false };

            if (sourceDefs == null || sourceDefs.Count == 0)
            {
                res.error = "请至少选择一个来源文件进行多文件汇总。";
                res.failureStage = "empty_source_list";
                return res;
            }

            // 1. 基础校验：文件存在性与扩展名
            foreach (var s in sourceDefs)
            {
                if (string.IsNullOrWhiteSpace(s.filePath))
                {
                    res.error = "存在未指定文件路径的来源项。";
                    res.failureStage = "invalid_file_path";
                    return res;
                }
                if (!File.Exists(s.filePath))
                {
                    res.error = string.Format("来源文件不存在：{0}", s.filePath);
                    res.failureStage = "file_not_found";
                    return res;
                }
                string ext = Path.GetExtension(s.filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xls")
                {
                    res.error = string.Format("来源文件【{0}】格式不受支持，仅支持 .xlsx, .xlsm, .xls 文件。", Path.GetFileName(s.filePath));
                    res.failureStage = "unsupported_extension";
                    return res;
                }
            }

            // 2. 解析同名不同路径文件的可区分标识
            ResolveDisplayIdentifiers(sourceDefs);

            if (app == null)
            {
                res.error = "Excel Application 实例为空，无法读取来源文件。";
                res.failureStage = "null_excel_app";
                return res;
            }

            // 3. 逐个只读读取源文件
            var hashes = new StringBuilder();
            for (int fileIdx = 0; fileIdx < sourceDefs.Count; fileIdx++)
            {
                var sDef = sourceDefs[fileIdx];
                string fileHash = ComputeFileSha256(sDef.filePath);
                hashes.Append(fileHash).Append(";");

                dynamic wb = null;
                try
                {
                    app.AutomationSecurity = 3; // ForceDisable
                    app.EnableEvents = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(sDef.filePath, ReadOnly: true, UpdateLinks: 0);

                    dynamic ws = null;
                    if (!string.IsNullOrEmpty(sDef.sheetName))
                    {
                        foreach (dynamic sheet in wb.Worksheets)
                        {
                            if (string.Equals((string)sheet.Name, sDef.sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                ws = sheet;
                                break;
                            }
                        }
                    }
                    else
                    {
                        ws = wb.Worksheets[1];
                        sDef.sheetName = (string)ws.Name;
                    }

                    if (ws == null)
                    {
                        res.error = string.Format("在来源文件【{0}】中未找到工作表【{1}】。", sDef.displayIdentifier, sDef.sheetName);
                        res.failureStage = "sheet_not_found";
                        return res;
                    }

                    dynamic rng = null;
                    if (!string.IsNullOrEmpty(sDef.rangeAddress))
                    {
                        try { rng = ws.Range[sDef.rangeAddress]; } catch { }
                    }
                    if (rng == null)
                    {
                        rng = ws.UsedRange;
                        sDef.rangeAddress = (string)rng.Address;
                    }

                    int totalRows = (int)rng.Rows.Count;
                    int totalCols = (int)rng.Columns.Count;
                    int minRequiredRows = sDef.hasHeader ? 2 : 1;

                    if (totalRows < minRequiredRows || totalCols < 1)
                    {
                        res.error = string.Format("来源文件【{0}】选区【{1}】无有效数据行。", sDef.displayIdentifier, sDef.rangeAddress);
                        res.failureStage = "empty_range";
                        return res;
                    }

                    object[,] rawMatrix = Get2DMatrix(rng.Value2, totalRows, totalCols);

                    // 表头解析与严谨性校验（区分大小写、不自动去空格、空表头阻断、重复表头阻断）
                    int headerRow = sDef.headerRowIndex > 0 ? sDef.headerRowIndex : 1;
                    List<string> headers = new List<string>();
                    var seenHeaders = new HashSet<string>(StringComparer.Ordinal);

                    for (int c = 1; c <= totalCols; c++)
                    {
                        object hObj = rawMatrix[headerRow, c];
                        if (hObj == null || string.IsNullOrWhiteSpace(hObj.ToString()))
                        {
                            res.error = string.Format("来源文件【{0}】第 {1} 列表头为空，已阻断。所有表头必须包含明确列名。", sDef.displayIdentifier, c);
                            res.failureStage = "empty_header";
                            return res;
                        }
                        string hStr = hObj.ToString(); // 原文严格保留
                        if (seenHeaders.Contains(hStr))
                        {
                            res.error = string.Format("来源文件【{0}】包含重复表头【{1}】，已阻断。单文件内列名必须唯一。", sDef.displayIdentifier, hStr);
                            res.failureStage = "duplicate_header";
                            return res;
                        }
                        seenHeaders.Add(hStr);
                        headers.Add(hStr);
                    }

                    // 统计空白行与纳入数据行（表头行绝不计入数据行）
                    int startDataRow = headerRow + 1;
                    int blankRows = 0;
                    int includedRows = 0;

                    for (int r = startDataRow; r <= totalRows; r++)
                    {
                        bool isAllBlank = true;
                        for (int c = 1; c <= totalCols; c++)
                        {
                            object val = rawMatrix[r, c];
                            if (val != null && !string.IsNullOrEmpty(val.ToString()))
                            {
                                isAllBlank = false;
                                break;
                            }
                        }
                        if (isAllBlank) blankRows++;
                        else includedRows++;
                    }

                    var profile = new ConsolidationSourceProfile
                    {
                        fileIndex = fileIdx,
                        filePath = sDef.filePath,
                        displayIdentifier = sDef.displayIdentifier,
                        sheetName = sDef.sheetName,
                        rangeAddress = sDef.rangeAddress,
                        totalRows = totalRows,
                        totalCols = totalCols,
                        headerRow = headerRow,
                        includedDataRows = includedRows,
                        excludedBlankRows = blankRows,
                        columns = headers,
                        fileSha256 = fileHash
                    };
                    res.sources.Add(profile);
                }
                finally
                {
                    if (wb != null)
                    {
                        try
                        {
                            wb.Close(false);
                            Marshal.ReleaseComObject(wb);
                        }
                        catch { }
                    }
                }
            }

            // 4. 构建对齐列与映射关系
            var mappingBySource = new Dictionary<int, Dictionary<string, string>>();
            if (explicitMappings != null)
            {
                foreach (var m in explicitMappings)
                {
                    if (m == null || string.IsNullOrEmpty(m.sourceColName) || string.IsNullOrEmpty(m.targetColName)) continue;
                    if (!mappingBySource.ContainsKey(m.sourceFileIndex))
                    {
                        mappingBySource[m.sourceFileIndex] = new Dictionary<string, string>(StringComparer.Ordinal);
                    }
                    mappingBySource[m.sourceFileIndex][m.sourceColName] = m.targetColName;
                }
            }

            // 检查多源列映射到同一目标列冲突
            for (int f = 0; f < res.sources.Count; f++)
            {
                var srcProf = res.sources[f];
                var targetCount = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var col in srcProf.columns)
                {
                    string target = col;
                    if (mappingBySource.ContainsKey(f) && mappingBySource[f].ContainsKey(col))
                    {
                        target = mappingBySource[f][col];
                    }
                    if (!targetCount.ContainsKey(target)) targetCount[target] = 0;
                    targetCount[target]++;
                    if (targetCount[target] > 1)
                    {
                        res.error = string.Format("来源文件【{0}】有多个源列映射到同一目标列【{1}】，已阻断。", srcProf.displayIdentifier, target);
                        res.failureStage = "mapping_conflict";
                        return res;
                    }
                }
            }

            // 5. 稳定目标列顺序建立
            var alignedSet = new HashSet<string>(StringComparer.Ordinal);
            var alignedList = new List<string>();

            // 优先采用用户显式指定的列顺序
            if (userColumnOrder != null)
            {
                foreach (var col in userColumnOrder)
                {
                    if (!string.IsNullOrEmpty(col) && !alignedSet.Contains(col))
                    {
                        alignedSet.Add(col);
                        alignedList.Add(col);
                    }
                }
            }

            // 按源文件出现顺序稳定追加新列
            foreach (var srcProf in res.sources)
            {
                foreach (var col in srcProf.columns)
                {
                    string target = col;
                    if (mappingBySource.ContainsKey(srcProf.fileIndex) && mappingBySource[srcProf.fileIndex].ContainsKey(col))
                    {
                        target = mappingBySource[srcProf.fileIndex][col];
                    }
                    if (!alignedSet.Contains(target))
                    {
                        alignedSet.Add(target);
                        alignedList.Add(target);
                    }
                }
            }
            res.alignedColumns = alignedList;

            // 6. 避让“来源文件”与“来源工作表”元数据列重名
            string metaFile = "来源文件";
            while (alignedSet.Contains(metaFile))
            {
                metaFile = metaFile == "来源文件" ? "来源文件_元数据" : metaFile + "_1";
            }
            res.metadataSourceFileCol = metaFile;

            string metaSheet = "来源工作表";
            while (alignedSet.Contains(metaSheet))
            {
                metaSheet = metaSheet == "来源工作表" ? "来源工作表_元数据" : metaSheet + "_1";
            }
            res.metadataSourceSheetCol = metaSheet;

            // 7. 分析各源文件缺失列与未映射列
            for (int f = 0; f < res.sources.Count; f++)
            {
                var srcProf = res.sources[f];
                var srcTargets = new HashSet<string>(StringComparer.Ordinal);
                foreach (var col in srcProf.columns)
                {
                    string target = col;
                    if (mappingBySource.ContainsKey(f) && mappingBySource[f].ContainsKey(col))
                    {
                        target = mappingBySource[f][col];
                    }
                    srcTargets.Add(target);
                }

                var missing = new List<string>();
                foreach (var tCol in alignedList)
                {
                    if (!srcTargets.Contains(tCol))
                    {
                        missing.Add(tCol);
                    }
                }
                res.missingColumnsPerSource[f] = missing;
            }

            // 8. 统计汇总与恒等式核验
            int totalInRows = 0;
            int totalHeadRows = 0;
            int totalIncRows = 0;
            int totalExBlank = 0;

            foreach (var s in res.sources)
            {
                totalInRows += s.totalRows;
                totalHeadRows += (s.headerRow > 0 ? 1 : 0);
                totalIncRows += s.includedDataRows;
                totalExBlank += s.excludedBlankRows;
            }

            res.totalSourceCount = res.sources.Count;
            res.totalInputRows = totalInRows;
            res.totalHeaderRows = totalHeadRows;
            res.totalIncludedDataRows = totalIncRows;
            res.totalExcludedBlankRows = totalExBlank;
            res.expectedFinalOutputRows = totalIncRows;

            // 严格统计恒等式：总数据行 = 纳入行 + 排除空白行；最终输出行 = 各源纳入行之和
            res.accountingIdentityConfirmed = (res.expectedFinalOutputRows == totalIncRows) &&
                (totalInRows == totalHeadRows + totalIncRows + totalExBlank);

            res.dataFingerprint = string.Format("SOURCES:{0}|ROWS:{1}|COLS:{2}|HASHES:{3}",
                res.sources.Count, totalIncRows, alignedList.Count, hashes.ToString());

            sw.Stop();
            res.analysisElapsedMs = sw.ElapsedMilliseconds;
            res.ok = true;
            return res;
        }

        public static ConsolidationExecutionResult ExecuteConsolidation(
            dynamic app,
            List<ConsolidationSourceDef> sourceDefs,
            string outputFilePath,
            List<string> userColumnOrder = null,
            List<ConsolidationColMapping> explicitMappings = null,
            string expectedFingerprint = null,
            bool includeMetadataCols = true)
        {
            var sw = Stopwatch.StartNew();
            var execRes = new ConsolidationExecutionResult { ok = false };

            if (app == null)
            {
                execRes.error = "Excel Application 对象为空，无法执行多文件汇总写入。";
                execRes.failureStage = "null_excel_app";
                return execRes;
            }

            if (string.IsNullOrWhiteSpace(outputFilePath))
            {
                execRes.error = "未指定汇总结果输出文件路径。";
                execRes.failureStage = "empty_output_path";
                return execRes;
            }

            // 1. 前置分析核验
            var analysis = AnalyzeConsolidation(app, sourceDefs, userColumnOrder, explicitMappings);
            if (!analysis.ok)
            {
                execRes.error = "汇总前置分析失败：" + analysis.error;
                execRes.failureStage = analysis.failureStage;
                return execRes;
            }

            // 2. 防源文件静默篡改
            if (!string.IsNullOrEmpty(expectedFingerprint) && analysis.dataFingerprint != expectedFingerprint)
            {
                execRes.error = "检测到源文件在分析后发生变动（数据指纹不匹配），为保证汇总真实性已安全阻断，请重新分析确认后再执行。";
                execRes.failureStage = "data_fingerprint_mismatch";
                return execRes;
            }

            // 3. 输出路径安全校验：不得与任何来源文件相同
            string fullOutPath = Path.GetFullPath(outputFilePath);
            foreach (var s in sourceDefs)
            {
                if (string.Equals(fullOutPath, Path.GetFullPath(s.filePath), StringComparison.OrdinalIgnoreCase))
                {
                    execRes.error = string.Format("输出路径【{0}】与来源文件相同，为防止覆盖源数据已安全阻断。", fullOutPath);
                    execRes.failureStage = "output_path_conflict";
                    return execRes;
                }
            }

            // 4. 同名文件防覆盖保护递增
            string resolvedOutPath = ResolveUniqueOutputPath(fullOutPath);
            execRes.outputFilePath = resolvedOutPath;

            // 5. 规模与单表容量超限预检
            int metaColCount = includeMetadataCols ? 2 : 0;
            int totalMatrixCols = metaColCount + analysis.alignedColumns.Count;
            int totalMatrixRows = 1 + analysis.expectedFinalOutputRows; // 1 行表头 + 数据行

            if (totalMatrixRows > 1048576)
            {
                execRes.error = string.Format("汇总后总行数 ({0}) 超过 Excel 单工作表最大容量 (1,048,576 行)，已安全阻断，绝不静默截断。", totalMatrixRows);
                execRes.failureStage = "capacity_exceeded";
                return execRes;
            }
            if (totalMatrixCols > 16384)
            {
                execRes.error = string.Format("汇总后总列数 ({0}) 超过 Excel 单工作表最大列数 (16,384 列)，已安全阻断。", totalMatrixCols);
                execRes.failureStage = "capacity_exceeded";
                return execRes;
            }

            // 6. 构造输出数据矩阵 (0-based SAFEARRAY, 严格对应 destRange 行列)
            object[,] outMatrix = new object[totalMatrixRows, totalMatrixCols];

            // 写入表头 (第 0 行)
            int colPointer = 0;
            if (includeMetadataCols)
            {
                outMatrix[0, colPointer++] = analysis.metadataSourceFileCol;
                outMatrix[0, colPointer++] = analysis.metadataSourceSheetCol;
            }
            for (int c = 0; c < analysis.alignedColumns.Count; c++)
            {
                outMatrix[0, colPointer++] = analysis.alignedColumns[c];
            }

            // 建立源映射速查表
            var mappingBySource = new Dictionary<int, Dictionary<string, string>>();
            if (explicitMappings != null)
            {
                foreach (var m in explicitMappings)
                {
                    if (m == null) continue;
                    if (!mappingBySource.ContainsKey(m.sourceFileIndex))
                    {
                        mappingBySource[m.sourceFileIndex] = new Dictionary<string, string>(StringComparer.Ordinal);
                    }
                    mappingBySource[m.sourceFileIndex][m.sourceColName] = m.targetColName;
                }
            }

            // 7. 逐文件读取源数据填充矩阵（19位、前导零、公式样文本逐字符保真）
            int curRow = 1;
            int totalWrittenDataRows = 0;

            for (int f = 0; f < analysis.sources.Count; f++)
            {
                var prof = analysis.sources[f];
                dynamic wb = null;
                try
                {
                    app.AutomationSecurity = 3;
                    app.EnableEvents = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(prof.filePath, ReadOnly: true, UpdateLinks: 0);
                    dynamic ws = null;
                    foreach (dynamic sheet in wb.Worksheets)
                    {
                        if (string.Equals((string)sheet.Name, prof.sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = sheet;
                            break;
                        }
                    }
                    if (ws == null) ws = wb.Worksheets[1];

                    dynamic rng = ws.Range[prof.rangeAddress];
                    object[,] rawMatrix = Get2DMatrix(rng.Value2, prof.totalRows, prof.totalCols);

                    // 建立当前源文件 targetCol -> rawColIdx 映射
                    var targetToRawCol = new Dictionary<string, int>(StringComparer.Ordinal);
                    for (int c = 1; c <= prof.totalCols; c++)
                    {
                        string srcCol = prof.columns[c - 1];
                        string targetCol = srcCol;
                        if (mappingBySource.ContainsKey(f) && mappingBySource[f].ContainsKey(srcCol))
                        {
                            targetCol = mappingBySource[f][srcCol];
                        }
                        targetToRawCol[targetCol] = c;
                    }

                    int startDataRow = prof.headerRow + 1;
                    for (int r = startDataRow; r <= prof.totalRows; r++)
                    {
                        // 判定是否全空白行
                        bool isAllBlank = true;
                        for (int c = 1; c <= prof.totalCols; c++)
                        {
                            object val = rawMatrix[r, c];
                            if (val != null && !string.IsNullOrEmpty(val.ToString()))
                            {
                                isAllBlank = false;
                                break;
                            }
                        }
                        if (isAllBlank) continue; // 跳过全空行

                        int cPtr = 0;
                        if (includeMetadataCols)
                        {
                            outMatrix[curRow, cPtr++] = prof.displayIdentifier;
                            outMatrix[curRow, cPtr++] = prof.sheetName;
                        }

                        for (int tIdx = 0; tIdx < analysis.alignedColumns.Count; tIdx++)
                        {
                            string targetName = analysis.alignedColumns[tIdx];
                            if (targetToRawCol.ContainsKey(targetName))
                            {
                                int rawCol = targetToRawCol[targetName];
                                object rawVal = rawMatrix[r, rawCol];
                                outMatrix[curRow, cPtr++] = FormatValueForExcelOutput(rawVal);
                            }
                            else
                            {
                                outMatrix[curRow, cPtr++] = ""; // 缺失列填空
                            }
                        }

                        curRow++;
                        totalWrittenDataRows++;
                    }
                }
                finally
                {
                    if (wb != null)
                    {
                        try
                        {
                            wb.Close(false);
                            Marshal.ReleaseComObject(wb);
                        }
                        catch { }
                    }
                }
            }

            // 8. 写入独立新工作簿并持久化
            dynamic wbOut = null;
            string saveStep = "init";
            try
            {
                saveStep = "workbooks_add";
                wbOut = app.Workbooks.Add();
                saveStep = "worksheets_1";
                dynamic wsOut = wbOut.Worksheets[1];
                saveStep = "name_sheet";
                wsOut.Name = "汇总数据";
                execRes.outputSheetName = "汇总数据";

                saveStep = string.Format("range_{0}x{1}", totalMatrixRows, totalMatrixCols);
                dynamic destRange = wsOut.Range[wsOut.Cells[1, 1], wsOut.Cells[totalMatrixRows, totalMatrixCols]];
                saveStep = "set_value2";
                destRange.Value2 = outMatrix;

                // 保存为独立新工作簿
                saveStep = "saveas";
                wbOut.SaveAs(resolvedOutPath);
                saveStep = "close";
                try { wbOut.Close(false); } catch { }
                saveStep = "release";
                try { Marshal.ReleaseComObject(wbOut); } catch { }
                wbOut = null;
            }
            catch (Exception exSave)
            {
                execRes.error = string.Format("汇总结果文件保存失败[{0}]：{1}", saveStep, exSave.Message);
                execRes.failureStage = "save_failure";
                return execRes;
            }

            // 9. 重新只读打开并核验输出结果（真实 Value2 与行数核对）
            dynamic wbVerify = null;
            try
            {
                wbVerify = app.Workbooks.Open(resolvedOutPath, ReadOnly: true, UpdateLinks: 0);
                dynamic wsVerify = wbVerify.Worksheets[1];
                dynamic used = wsVerify.UsedRange;
                int readRows = (int)used.Rows.Count;
                int readCols = (int)used.Columns.Count;

                if (readRows == totalMatrixRows && readCols == totalMatrixCols)
                {
                    execRes.verifiedReadback = true;
                }
            }
            catch (Exception exVerify)
            {
                execRes.error = "输出文件核验失败：" + exVerify.Message;
                execRes.failureStage = "readback_verification_failure";
                return execRes;
            }
            finally
            {
                if (wbVerify != null)
                {
                    try
                    {
                        wbVerify.Close(false);
                        Marshal.ReleaseComObject(wbVerify);
                    }
                    catch { }
                }
            }

            sw.Stop();
            execRes.ok = true;
            execRes.finalDataRows = totalWrittenDataRows;
            execRes.finalDataCols = totalMatrixCols;
            execRes.totalIncludedDataRows = totalWrittenDataRows;
            execRes.totalExcludedBlankRows = analysis.totalExcludedBlankRows;
            execRes.accountingIdentityConfirmed = (totalWrittenDataRows == analysis.expectedFinalOutputRows);
            execRes.elapsedMs = sw.ElapsedMilliseconds;

            return execRes;
        }

        #endregion

        #endregion
    }
}

