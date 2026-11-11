using System;
using System.Collections.Generic;
using System.Text;
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
            // 关键词判断前移除字符串字面量，避免把 "Shell" 等普通文本误判为外部调用；
            // 同时保留代码中的调用位置，确保 Call Shell(...) / Debug.Print Shell(...) 仍能被识别。
            string keywordCode = StripStringLiterals(cleanCode);

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
            if (Regex.IsMatch(cleanCode, @"\b(?:Workbooks\s*\.\s*(?:Open|SaveAs|Close)|(?:ThisWorkbook|ActiveWorkbook|Application\.ActiveWorkbook)\s*\.\s*(?:Close|SaveAs))\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含工作簿打开/另存/关闭操作");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部工作簿操作";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "external_workbook";
            }

            // 2.3 高危外部 COM 组件 (FileSystemObject / 网络请求等)
            if (Regex.IsMatch(keywordCode, @"\b(?:VBA\.)?(?:CreateObject|GetObject|Shell|SendKeys)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(keywordCode, @"\bNew\s+(?:FileSystemObject|ADODB\.|Scripting\.|MSXML2\.|WinHttp\.|WScript\.)", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(keywordCode, @"(?im)^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含外部 COM、Shell、SendKeys、GetObject 或 Win32 API 调用");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部系统与文件调用";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "system_com";
            }

            // 2.4 篡改代码工程
            if (Regex.IsMatch(keywordCode, @"\b(?:VBProject|ThisWorkbook\.VBProject)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含对 VBA 代码工程自身的访问与修改 (VBProject)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及代码工程修改";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "modify_vbproject";
            }

            // 2.5 应用级动作与外部跳转
            if (Regex.IsMatch(keywordCode, @"\bApplication\s*\.\s*(?:Run|Quit)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(keywordCode, @"(?:\.\s*)?FollowHyperlink\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(keywordCode, @"\b(?:ADODB|MSXML2|WinHttp|WScript)\.[A-Za-z_][A-Za-z0-9_]*", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含应用级执行、退出、外部跳转或脚本对象调用");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及应用级或外部系统副作用";
                if (string.IsNullOrEmpty(res.RiskCode)) res.RiskCode = "external_side_effect";
            }

            // 2.6 原生文件操作语句 (Kill / FileCopy / Name ... As ...)
            if (Regex.IsMatch(keywordCode, @"\b(?:Kill|FileCopy|RmDir)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(keywordCode, @"\bName\b\s+.+\s+\bAs\b", RegexOptions.IgnoreCase) ||
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
                Regex.IsMatch(cleanCode, @"\.(?:ListObjects|PivotTables|PivotCaches|QueryTables|Tables)\s*\.\s*(?:Add|Create)\b", RegexOptions.IgnoreCase);

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

        private static string GetAssignmentRightHandSide(string line)
        {
            int idx = line.IndexOf('=');
            return idx >= 0 ? line.Substring(idx + 1).Trim() : "";
        }

        private static bool IsClearlyReadOnlyValueExpression(string line)
        {
            string rhs = GetAssignmentRightHandSide(line);
            if (string.IsNullOrEmpty(rhs)) return false;

            if (Regex.IsMatch(rhs, @"^(?:""(?:""""|[^""])*""|'(?:''|[^'])*'|[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?|True|False|Nothing|Empty|Null)$", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return Regex.IsMatch(rhs, @"^(?:Application\.)?(?:(?:ActiveSheet|ActiveCell|Selection|ThisWorkbook|ActiveWorkbook|Worksheets\s*\([^)]*\)|Sheets\s*\([^)]*\)|Range\s*\(.*\)|Cells\s*\([^)]*\))(?:\.(?:Range|Cells)\s*\(.*\))?\.(?:Value2?|Text|Formula|FormulaR1C1|Address|Row|Column|Count|Name|Path|FullName))$", RegexOptions.IgnoreCase);
        }

        private static bool IsClearlyReadOnlySetExpression(string line)
        {
            string rhs = GetAssignmentRightHandSide(line);
            if (string.IsNullOrEmpty(rhs)) return false;

            return Regex.IsMatch(rhs, @"^(?:Nothing|ActiveSheet|ActiveCell|Selection|ThisWorkbook|ActiveWorkbook|(?:Application\.)?(?:Worksheets|Sheets|Range|Cells)\s*\([^)]*\)|[A-Za-z_][A-Za-z0-9_]*\.(?:Worksheets|Sheets|Range|Cells)\s*\([^)]*\))$", RegexOptions.IgnoreCase);
        }

        private static string StripStringLiterals(string code)
        {
            if (string.IsNullOrEmpty(code)) return "";
            var sb = new StringBuilder(code.Length);
            bool inString = false;
            for (int i = 0; i < code.Length; i++)
            {
                char c = code[i];
                if (c == '"')
                {
                    // VBA 使用两个连续双引号表示字符串内的一个字面量双引号。
                    if (inString && i + 1 < code.Length && code[i + 1] == '"')
                    {
                        i++;
                        sb.Append(' ');
                        continue;
                    }
                    inString = !inString;
                    sb.Append(' ');
                    continue;
                }
                sb.Append(inString ? ' ' : c);
            }
            return sb.ToString();
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

                // 仅接受可证明为读取的简单赋值；任何函数调用或未知表达式都按 fail-safe 处理
                if (Regex.IsMatch(line, @"^[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*", RegexOptions.IgnoreCase) &&
                    IsClearlyReadOnlyValueExpression(line))
                {
                    continue;
                }

                if (Regex.IsMatch(line, @"^Set\s+[a-zA-Z_][a-zA-Z0-9_]*\s*=\s*", RegexOptions.IgnoreCase) &&
                    IsClearlyReadOnlySetExpression(line))
                {
                    continue;
                }

                // 一旦存在任何未在纯读白名单内的语句，即无法严格证明是纯读
                return false;
            }

            return true;
        }
    }
}