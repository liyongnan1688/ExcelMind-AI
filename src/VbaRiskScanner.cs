using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public enum VbaRiskLevel
    {
        L0, // 只读查询，不修改工作表内容，直接执行且不创建快照
        L1, // 原地写操作，创建快照并自动执行，支持一键撤销
        L2  // 高危/越界/不可逆，严格阻断执行，返回 needsConfirm 并待二次确认
    }

    public class VbaRiskAnalysis
    {
        public VbaRiskLevel RiskLevel { get; set; }
        public string RiskCode { get; set; }
        public List<string> Reasons { get; set; }
        public string TargetDetail { get; set; }

        public VbaRiskAnalysis()
        {
            Reasons = new List<string>();
        }
    }

    /// <summary>
    /// VBA 代码执行前风险等级静态扫描器 (P5 规范)
    /// </summary>
    public static class VbaRiskScanner
    {
        public static VbaRiskAnalysis Analyze(string vbaCode, string targetWorkbookFullName)
        {
            var res = new VbaRiskAnalysis();

            if (string.IsNullOrWhiteSpace(vbaCode))
            {
                res.RiskLevel = VbaRiskLevel.L0;
                res.Reasons.Add("代码为空");
                return res;
            }

            // 1. 预处理：剔除单行注释和行尾注释，防止注释内容干扰规则匹配
            string[] lines = vbaCode.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> cleanLines = new List<string>();
            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.StartsWith("'") || line.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                int quoteIdx = line.IndexOf('\'');
                if (quoteIdx >= 0)
                {
                    string prefix = line.Substring(0, quoteIdx);
                    int doubleQuotes = 0;
                    for (int i = 0; i < prefix.Length; i++)
                    {
                        if (prefix[i] == '"') doubleQuotes++;
                    }
                    if (doubleQuotes % 2 == 0)
                    {
                        line = prefix.Trim();
                    }
                }
                if (!string.IsNullOrEmpty(line))
                {
                    cleanLines.Add(line);
                }
            }

            string cleanCode = string.Join("\n", cleanLines.ToArray());

            // 2. L2 规则判定 (命中任一即算 L2)
            // 2.1 删除工作表
            if (Regex.IsMatch(cleanCode, @"(?:Worksheets|Sheets)\s*\([^)]*\)\s*\.Delete", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"(?:ActiveSheet|ThisWorkbook\.ActiveSheet)\.Delete", RegexOptions.IgnoreCase))
            {
                Match m = Regex.Match(cleanCode, @"(?:Worksheets|Sheets)\s*\(\s*[""']([^""']+)[""']\s*\)\s*\.Delete", RegexOptions.IgnoreCase);
                string sheetName = m.Success ? m.Groups[1].Value : "工作表";
                res.Reasons.Add("包含删除工作表指令 (将删除工作表《" + sheetName + "》)");
                res.TargetDetail = "将删除工作表《" + sheetName + "》";
                res.RiskCode = "delete_sheet";
            }

            // 2.2 外部工作簿操作 (Open / SaveAs / Close)
            if (Regex.IsMatch(cleanCode, @"\bWorkbooks\s*\.\s*(?:Open|SaveAs|Close)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含外部工作簿打开/另存/关闭操作 (Workbooks.Open/SaveAs/Close)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部工作簿操作";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "external_workbook";
            }

            // 2.3 高危外部 COM 组件 (FileSystemObject / 网络请求等)
            if (Regex.IsMatch(cleanCode, @"CreateObject\s*\(\s*[""'](?:Scripting\.FileSystemObject|MSXML2\.[^""']+|WinHttp\.[^""']+|InternetExplorer\.Application|WScript\.Shell)[""']", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"New\s+FileSystemObject\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含文件系统或外部网络组件调用 (FileSystemObject/WinHttp/Shell)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部系统与文件调用";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "system_com";
            }

            // 2.4 篡改代码工程
            if (Regex.IsMatch(cleanCode, @"\b(?:VBProject|ThisWorkbook\.VBProject)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含对 VBA 代码工程自身的访问与修改 (VBProject)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及代码工程修改";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "modify_vbproject";
            }

            // 2.5 原生文件操作语句 (Kill / FileCopy / Name ... As ...)
            if (Regex.IsMatch(cleanCode, @"\b(?:Kill|FileCopy|RmDir)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bName\b\s+.+\s+\bAs\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bOpen\b\s+[""'][^""']+[""']\s+For\s+(?:Output|Append|Binary)", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含文件删除或外部磁盘写入语句 (Kill/FileCopy/Open For Output)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部文件写入/删除";
                if (string.IsNullOrEmpty(res.RiskCode))
                {
                    res.RiskCode = Regex.IsMatch(cleanCode, @"\bKill\b", RegexOptions.IgnoreCase) ? "delete_file" : "file_io";
                }
            }

            // 若命中任意 L2 规则，直接返回 L2
            if (res.Reasons.Count > 0)
            {
                res.RiskLevel = VbaRiskLevel.L2;
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "不可逆操作";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "dangerous_action";
                return res;
            }

            // 3. L1 规则判定：有写操作
            bool hasWrite =
                Regex.IsMatch(cleanCode, @"\.(?:Value|Value2|Formula|FormulaR1C1|Formula2)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Interior|Font|Borders)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Clear|ClearContents|ClearFormats)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"(?:Rows|Columns)\s*\([^)]*\)\s*\.(?:Insert|Delete)", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:NumberFormat|NumberFormatLocal)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:RowHeight|ColumnWidth)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bRange\s*\([^)]*\)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.Cells\s*\([^)]*\)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:AddComment|WrapText|Merge|UnMerge)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"(?:Charts|ChartObjects|Shapes)\s*\.\s*Add", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:EntireRow|EntireColumn)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Sort|SortSpecial)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Copy|Cut)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Paste|PasteSpecial)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Name)\s*=", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:Group|Ungroup)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:AutoFill|Replace)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\.(?:ListObjects|PivotTables|PivotCaches|QueryTables|Tables)\s*\.\s*(?:Add|Create)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bApplication\s*\.\s*Run\b", RegexOptions.IgnoreCase);

            if (hasWrite)
            {
                res.RiskLevel = VbaRiskLevel.L1;
                res.Reasons.Add("包含普通单元格或工作表修改操作，已创建快照支持一键撤销");
                return res;
            }

            // 4. L0 纯读核验：仅当能严格证明为已知纯读操作，才放行为 L0 (免快照)
            if (IsStrictlyReadOnly(cleanLines))
            {
                res.RiskLevel = VbaRiskLevel.L0;
                res.Reasons.Add("只读操作，不修改工作簿内容");
                return res;
            }

            // 5. 兜底：无法证明是只读 → 按 L1 处理 (fail-safe)
            res.RiskLevel = VbaRiskLevel.L1;
            res.Reasons.Add("未识别到明确写操作，按可回滚写入处理并创建快照");
            return res;
        }

        private static bool HasMultiStatementColon(string line)
        {
            bool inQuote = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inQuote = !inQuote;
                }
                else if (!inQuote && c == ':')
                {
                    // 排除命名参数 :=
                    if (i + 1 < line.Length && line[i + 1] == '=')
                    {
                        i++;
                        continue;
                    }
                    return true;
                }
            }
            return false;
        }

        private static bool IsStrictlyReadOnly(List<string> cleanLines)
        {
            if (cleanLines == null || cleanLines.Count == 0) return true;

            foreach (var rawLine in cleanLines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                // 遇到冒号分隔的多语句行，直接判定为非纯读 (返回 false)
                if (HasMultiStatementColon(line))
                {
                    return false;
                }

                // 允许的过程控制/结构声明语句
                if (Regex.IsMatch(line, @"^(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\b", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^End\s+(?:Sub|Function)\b", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^(?:Dim|Const|ReDim)\b", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^Set\s+[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*(?!.*(?:\.Delete|\.Insert|\.Clear|\.Add|\.Copy|\.Paste|\.Sort|\.Group|\.Ungroup|\.AutoFill|\.Replace))", RegexOptions.IgnoreCase) ||
                    // 控制流语句只允许换行块形式（行尾无后续执行语句），带 Then <语句> 的单行 If 一律不放行
                    Regex.IsMatch(line, @"^If\b.+\bThen\s*$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^ElseIf\b.+\bThen\s*$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^Else\s*$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^End\s+If\s*$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^(?:For\b|Next\b|Do\b|Loop\b|While\b|Wend|Select\s+Case\b|Case\b|End\s+Select|With\b|End\s+With|Exit\s+(?:Sub|Function|For|Do))\b", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                // 允许的纯读输出 / 检查语句
                if (Regex.IsMatch(line, @"^MsgBox\b", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^Debug\.Print\b", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^(?:Application\.)?StatusBar\s*=", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^(?:Range|Cells|ActiveSheet|Worksheets|Sheets|ActiveCell|Selection)\b[^\.]*\.(?:Select|Activate)\b", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                // 允许将工作表只读属性赋值给局部变量 (左值只能是单纯标识符变量名，不能是包含点号的对象属性)
                if (Regex.IsMatch(line, @"^[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*", RegexOptions.IgnoreCase))
                {
                    // 确保等号右边没有写操作或方法调用
                    if (!Regex.IsMatch(line, @"(?:\.Delete|\.Insert|\.Clear|\.Add|\.Copy|\.Paste|\.Sort|\.Group|\.Ungroup|\.AutoFill|\.Replace)\b", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }
                }

                // 一旦存在任何未在纯读白名单内的语句，即无法严格证明是纯读
                return false;
            }

            return true;
        }
    }
}
