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
            }

            // 2.2 外部工作簿操作 (Open / SaveAs / Close)
            if (Regex.IsMatch(cleanCode, @"\bWorkbooks\s*\.\s*(?:Open|SaveAs|Close)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含外部工作簿打开/另存/关闭操作 (Workbooks.Open/SaveAs/Close)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部工作簿操作";
            }

            // 2.3 高危外部 COM 组件 (FileSystemObject / 网络请求等)
            if (Regex.IsMatch(cleanCode, @"CreateObject\s*\(\s*[""'](?:Scripting\.FileSystemObject|MSXML2\.[^""']+|WinHttp\.[^""']+|InternetExplorer\.Application|WScript\.Shell)[""']", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"New\s+FileSystemObject\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含文件系统或外部网络组件调用 (FileSystemObject/WinHttp/Shell)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部系统与文件调用";
            }

            // 2.4 篡改代码工程
            if (Regex.IsMatch(cleanCode, @"\b(?:VBProject|ThisWorkbook\.VBProject)\b", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含对 VBA 代码工程自身的访问与修改 (VBProject)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及代码工程修改";
            }

            // 2.5 原生文件操作语句 (Kill / FileCopy / Name ... As ...)
            if (Regex.IsMatch(cleanCode, @"\b(?:Kill|FileCopy|RmDir)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bName\b\s+.+\s+\bAs\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(cleanCode, @"\bOpen\b\s+[""'][^""']+[""']\s+For\s+(?:Output|Append|Binary)", RegexOptions.IgnoreCase))
            {
                res.Reasons.Add("包含文件删除或外部磁盘写入语句 (Kill/FileCopy/Open For Output)");
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "涉及外部文件写入/删除";
            }

            // 若命中任意 L2 规则，直接返回 L2
            if (res.Reasons.Count > 0)
            {
                res.RiskLevel = VbaRiskLevel.L2;
                if (string.IsNullOrEmpty(res.TargetDetail)) res.TargetDetail = "不可逆操作";
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
                Regex.IsMatch(cleanCode, @"(?:Charts|ChartObjects|Shapes)\s*\.\s*Add", RegexOptions.IgnoreCase);

            if (hasWrite)
            {
                res.RiskLevel = VbaRiskLevel.L1;
                res.Reasons.Add("包含普通单元格或工作表修改操作，已创建快照支持一键撤销");
                return res;
            }

            // 4. L0 规则判定：只读
            res.RiskLevel = VbaRiskLevel.L0;
            res.Reasons.Add("只读操作，不修改工作簿内容");
            return res;
        }
    }
}
