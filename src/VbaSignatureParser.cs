using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    /// <summary>
    /// 单个 VBA 参数的声明与元数据描述
    /// </summary>
    public class ScriptParameterDef
    {
        public string name { get; set; }
        public string type { get; set; }            // 标准化类型: "String", "Long", "Double", "Boolean", "Date", "Worksheet", "Range", "Workbook"
        public string typeName { get { return type; } set { type = value; } } // 兼容前端/元数据 typeName 别名
        public string rawType { get; set; }         // 源码原始类型: 如 "Integer", "Single"
        public string description { get; set; }
        public bool isOptional { get; set; }
        public string defaultValue { get; set; }
        public bool isTargetWorkbook { get; set; }   // 是否为宿主注入的目标工作簿参数 (targetWb As Workbook)
        public bool isSupported { get; set; }
        public string unsupportedReason { get; set; }
    }

    /// <summary>
    /// VBA 过程 (Sub / Function) 的签名分析结果
    /// </summary>
    public class VbaEntryPointInfo
    {
        public string name { get; set; }
        public string kind { get; set; }             // "Sub" | "Function"
        public string visibility { get; set; }       // "Public" | "Private" | "Friend" | "Default"
        public bool isExecutable { get; set; }       // 是否可作为宏直接执行
        public bool isSupported { get { return isExecutable; } set { isExecutable = value; } } // 兼容前端 isSupported 契约
        public string unsupportedReason { get; set; }
        public bool hasNativeWbParam { get; set; }   // 是否包含 targetWb As Workbook
        public List<ScriptParameterDef> parameters { get; set; }
        public string rawSignature { get; set; }
        public int lineIndex { get; set; }

        public VbaEntryPointInfo()
        {
            parameters = new List<ScriptParameterDef>();
            isExecutable = true;
        }
    }

    /// <summary>
    /// 元数据与源码签名比对结果
    /// </summary>
    public class SignatureComparisonResult
    {
        public bool isMatch { get; set; }
        public string error { get; set; }
        public string reason { get { return error; } set { error = value; } }
        public List<ScriptParameterDef> mergedParameters { get; set; }

        public SignatureComparisonResult()
        {
            mergedParameters = new List<ScriptParameterDef>();
        }
    }

    /// <summary>
    /// 严密的 VBA 过程签名与参数解析器
    /// </summary>
    public static class VbaSignatureParser
    {
        // 允许的标准支持类型映射
        private static readonly Dictionary<string, string> SupportedTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "String", "String" },
            { "Long", "Long" },
            { "Integer", "Long" },
            { "Byte", "Long" },
            { "LongLong", "Long" },
            { "LongPtr", "Long" },
            { "Double", "Double" },
            { "Single", "Double" },
            { "Currency", "Double" },
            { "Boolean", "Boolean" },
            { "Date", "Date" },
            { "Worksheet", "Worksheet" },
            { "Range", "Range" },
            { "Workbook", "Workbook" }
        };

        /// <summary>
        /// 解析 VBA 源码中的所有过程签名
        /// </summary>
        public static List<VbaEntryPointInfo> ParseSignatures(string vbaCode)
        {
            var results = new List<VbaEntryPointInfo>();
            if (string.IsNullOrWhiteSpace(vbaCode)) return results;

            // 预处理：合并以下划线 "_" 结尾的续行，保留行号对应
            string[] rawLines = vbaCode.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var consolidated = new List<Tuple<string, int>>();

            for (int i = 0; i < rawLines.Length; i++)
            {
                string line = rawLines[i];
                int startIdx = i;
                // 去掉行尾注释再判断是否有续行符
                string stripped = StripTrailingComment(line).TrimEnd();
                while (stripped.EndsWith("_") && i + 1 < rawLines.Length)
                {
                    stripped = stripped.Substring(0, stripped.Length - 1).TrimEnd();
                    i++;
                    string nextLine = StripTrailingComment(rawLines[i]).Trim();
                    stripped = stripped + " " + nextLine;
                }
                consolidated.Add(Tuple.Create(stripped, startIdx + 1));
            }

            // 过程定义正则：捕获 [Visibility] [Static] Sub/Function Name
            var procRegex = new Regex(
                @"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+([a-zA-Z_\u4e00-\u9fa5][a-zA-Z0-9_\u4e00-\u9fa5]*)",
                RegexOptions.IgnoreCase);

            foreach (var item in consolidated)
            {
                string line = item.Item1;
                int lineNo = item.Item2;

                // 排除以注释开头的行或属性指令
                if (line.StartsWith("'") || line.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase) || line.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Match m = procRegex.Match(line);
                if (m.Success)
                {
                    string vis = m.Groups[1].Success ? m.Groups[1].Value : "Default";
                    string kind = m.Groups[2].Value; // Sub | Function
                    string name = m.Groups[3].Value;
                    string paramStr = "";

                    // 使用括号深度计数提取完整参数列表，防止参数内含括号（如 arr() As Long）被非贪婪正则截断
                    int openParen = line.IndexOf('(', m.Index + m.Length);
                    if (openParen >= 0)
                    {
                        int pDepth = 0;
                        int closeParen = -1;
                        for (int pIdx = openParen; pIdx < line.Length; pIdx++)
                        {
                            if (line[pIdx] == '(') pDepth++;
                            else if (line[pIdx] == ')')
                            {
                                pDepth--;
                                if (pDepth == 0)
                                {
                                    closeParen = pIdx;
                                    break;
                                }
                            }
                        }
                        if (closeParen > openParen)
                        {
                            paramStr = line.Substring(openParen + 1, closeParen - openParen - 1);
                        }
                    }

                    var entry = new VbaEntryPointInfo
                    {
                        name = name,
                        kind = char.ToUpper(kind[0]) + kind.Substring(1).ToLower(),
                        visibility = char.ToUpper(vis[0]) + vis.Substring(1).ToLower(),
                        rawSignature = line.Trim(),
                        lineIndex = lineNo
                    };

                    // 1. Function 过程直接标记为不可执行
                    if (string.Equals(entry.kind, "Function", StringComparison.OrdinalIgnoreCase))
                    {
                        entry.isExecutable = false;
                        entry.unsupportedReason = "Function 过程为有返回值的函数，不能作为独立宏直接运行。请将其包装在 Sub 过程中调用。";
                    }
                    // 2. Private 过程直接标记为不可执行
                    else if (string.Equals(entry.visibility, "Private", StringComparison.OrdinalIgnoreCase))
                    {
                        entry.isExecutable = false;
                        entry.unsupportedReason = "过程 '" + name + "' 声明为 Private 私有过程，无法从功能区或外部宏入口直接调用。";
                    }

                    // 3. 解析参数列表
                    ParseParameters(paramStr, entry);

                    // 4. 若参数存在不支持项且过程当前未被标记错误，置为不可执行
                    if (entry.isExecutable)
                    {
                        var unsupportedParams = new List<string>();
                        foreach (var p in entry.parameters)
                        {
                            if (!p.isSupported)
                            {
                                unsupportedParams.Add(p.name + ": " + p.unsupportedReason);
                            }
                        }
                        if (unsupportedParams.Count > 0)
                        {
                            entry.isExecutable = false;
                            entry.unsupportedReason = "过程包含暂不支持的参数签名: " + string.Join("; ", unsupportedParams.ToArray());
                        }
                    }

                    results.Add(entry);
                }
            }

            return results;
        }

        private static void ParseParameters(string paramStr, VbaEntryPointInfo entry)
        {
            if (string.IsNullOrWhiteSpace(paramStr)) return;

            // 分割参数，注意处理括号可能包含的内容
            string[] rawParams = SplitParameters(paramStr);

            foreach (var rawP in rawParams)
            {
                string p = rawP.Trim();
                if (string.IsNullOrEmpty(p)) continue;

                var paramDef = new ScriptParameterDef
                {
                    isSupported = true
                };

                // 检查 ParamArray
                if (Regex.IsMatch(p, @"^ParamArray\s+", RegexOptions.IgnoreCase))
                {
                    paramDef.name = p.Replace("ParamArray", "").Trim();
                    paramDef.isSupported = false;
                    paramDef.unsupportedReason = "第一切片暂不支持 ParamArray 可变长参数声明。";
                    entry.parameters.Add(paramDef);
                    continue;
                }

                // 检查 Optional
                bool isOpt = false;
                if (Regex.IsMatch(p, @"^Optional\s+", RegexOptions.IgnoreCase))
                {
                    isOpt = true;
                    p = Regex.Replace(p, @"^Optional\s+", "", RegexOptions.IgnoreCase).Trim();
                }
                paramDef.isOptional = isOpt;

                // 检查 ByVal / ByRef
                p = Regex.Replace(p, @"^(?:ByVal|ByRef)\s+", "", RegexOptions.IgnoreCase).Trim();

                // 检查默认值: Name As Type = DefaultValue
                string defVal = null;
                int eqIdx = p.IndexOf('=');
                if (eqIdx >= 0)
                {
                    defVal = p.Substring(eqIdx + 1).Trim();
                    p = p.Substring(0, eqIdx).Trim();
                }
                paramDef.defaultValue = defVal;

                // 解析 名称 与 As Type
                // 如: targetWb As Workbook, wsName As String, arr() As Long
                var asMatch = Regex.Match(p, @"^([a-zA-Z_\u4e00-\u9fa5][a-zA-Z0-9_\u4e00-\u9fa5]*)(\s*\(\s*\))?\s*(?:As\s+(?:New\s+)?([a-zA-Z0-9_\u4e00-\u9fa5\.]+))?$", RegexOptions.IgnoreCase);
                if (asMatch.Success)
                {
                    paramDef.name = asMatch.Groups[1].Value;
                    bool isArray = asMatch.Groups[2].Success && !string.IsNullOrWhiteSpace(asMatch.Groups[2].Value);
                    string typeName = asMatch.Groups[3].Success ? asMatch.Groups[3].Value : null;

                    if (isArray)
                    {
                        paramDef.isSupported = false;
                        paramDef.unsupportedReason = "第一切片暂不支持数组参数 (Array())。";
                    }
                    else if (string.IsNullOrEmpty(typeName))
                    {
                        // 缺少 As Type，属于隐式 Variant
                        paramDef.isSupported = false;
                        paramDef.unsupportedReason = "参数未声明显式类型 (缺少 As Type)。宿主严格拒绝隐式 Variant 参数以确保类型安全。";
                    }
                    else
                    {
                        paramDef.rawType = typeName;
                        if (SupportedTypeMap.ContainsKey(typeName))
                        {
                            string standardType = SupportedTypeMap[typeName];
                            paramDef.type = standardType;

                            // 检查是否为 targetWb As Workbook
                            if (string.Equals(standardType, "Workbook", StringComparison.OrdinalIgnoreCase))
                            {
                                if (Regex.IsMatch(paramDef.name, @"^(?:targetWb|wb|workbook)$", RegexOptions.IgnoreCase))
                                {
                                    paramDef.isTargetWorkbook = true;
                                    entry.hasNativeWbParam = true;
                                }
                                else
                                {
                                    // 其他名字的 Workbook 参数，第一切片也直接绑定到目标工作簿
                                    paramDef.isTargetWorkbook = true;
                                    entry.hasNativeWbParam = true;
                                }
                            }
                        }
                        else
                        {
                            paramDef.isSupported = false;
                            paramDef.unsupportedReason = "暂不支持的参数类型 '" + typeName + "'。第一切片支持: String, Long, Double, Boolean, Date, Worksheet, Range。";
                        }
                    }
                }
                else
                {
                    paramDef.name = p;
                    paramDef.isSupported = false;
                    paramDef.unsupportedReason = "无法可靠解析的参数声明格式: '" + p + "'";
                }

                entry.parameters.Add(paramDef);
            }
        }

        private static string[] SplitParameters(string paramStr)
        {
            var list = new List<string>();
            var sb = new StringBuilder();
            int parenDepth = 0;

            for (int i = 0; i < paramStr.Length; i++)
            {
                char c = paramStr[i];
                if (c == '(') parenDepth++;
                else if (c == ')') parenDepth = Math.Max(0, parenDepth - 1);

                if (c == ',' && parenDepth == 0)
                {
                    list.Add(sb.ToString());
                    sb.Length = 0;
                }
                else
                {
                    sb.Append(c);
                }
            }
            if (sb.Length > 0)
            {
                list.Add(sb.ToString());
            }

            return list.ToArray();
        }

        private static string StripTrailingComment(string line)
        {
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == '\'' && !inQuotes)
                {
                    return line.Substring(0, i);
                }
            }
            return line;
        }

        /// <summary>
        /// 比对元数据中的参数声明与源码过程签名，发现冲突立即阻断
        /// </summary>
        /// <summary>
        /// 比对元数据中的参数声明与源码过程签名，发现冲突立即阻断
        /// 支持传入 List<ScriptParameterDef>、IEnumerable 或 null
        /// </summary>
        public static SignatureComparisonResult CompareWithMetadata(VbaEntryPointInfo proc, object metaParamsObj)
        {
            var res = new SignatureComparisonResult();

            // 过滤掉源码中的 targetWb As Workbook 参数（由宿主隐式绑定）
            var codeExplicitParams = new List<ScriptParameterDef>();
            foreach (var p in proc.parameters)
            {
                if (!p.isTargetWorkbook)
                {
                    codeExplicitParams.Add(p);
                }
            }

            // 标准化 metaParams 为 List<ScriptParameterDef>
            List<ScriptParameterDef> metaParams = null;
            if (metaParamsObj is List<ScriptParameterDef>)
            {
                metaParams = (List<ScriptParameterDef>)metaParamsObj;
            }
            else if (metaParamsObj is System.Collections.IEnumerable)
            {
                metaParams = new List<ScriptParameterDef>();
                foreach (var item in (System.Collections.IEnumerable)metaParamsObj)
                {
                    if (item is ScriptParameterDef)
                    {
                        metaParams.Add((ScriptParameterDef)item);
                    }
                    else if (item is IDictionary<string, object>)
                    {
                        var d = (IDictionary<string, object>)item;
                        metaParams.Add(new ScriptParameterDef
                        {
                            name = d.ContainsKey("name") && d["name"] != null ? d["name"].ToString() : "",
                            type = d.ContainsKey("type") && d["type"] != null ? d["type"].ToString() : "",
                            description = d.ContainsKey("description") && d["description"] != null ? d["description"].ToString() : ""
                        });
                    }
                    else if (item is IDictionary<string, string>)
                    {
                        var d = (IDictionary<string, string>)item;
                        metaParams.Add(new ScriptParameterDef
                        {
                            name = d.ContainsKey("name") ? d["name"] : "",
                            type = d.ContainsKey("type") ? d["type"] : "",
                            description = d.ContainsKey("description") ? d["description"] : ""
                        });
                    }
                }
            }

            // 如果元数据没有声明参数 (null 或 0 项)
            if (metaParams == null || metaParams.Count == 0)
            {
                // 若源码有显式参数，则直接使用源码中可靠识别的显式参数生成表单
                res.isMatch = true;
                res.mergedParameters = codeExplicitParams;
                return res;
            }

            // 如果元数据有声明参数，必须严格核对数量、顺序与类型
            if (metaParams.Count != codeExplicitParams.Count)
            {
                res.isMatch = false;
                res.error = "参数数量不一致：.meta.json 声明了 " + metaParams.Count + " 个参数，而源码过程 '" + proc.name + "' 包含 " + codeExplicitParams.Count + " 个参数。宿主拒绝猜测执行。";
                return res;
            }

            for (int i = 0; i < metaParams.Count; i++)
            {
                var mP = metaParams[i];
                var cP = codeExplicitParams[i];

                // 核对参数名称与顺序
                if (!string.IsNullOrEmpty(mP.name) && !string.Equals(mP.name, cP.name, StringComparison.OrdinalIgnoreCase))
                {
                    res.isMatch = false;
                    res.error = "参数名称或顺序不匹配：第 " + (i + 1) + " 个参数 .meta.json 声明为 '" + mP.name + "'，而源码过程为 '" + cP.name + "'。宿主严格拒绝乱序绑定。";
                    return res;
                }

                // 核对类型 (按标准化类型)
                string normMetaType = SupportedTypeMap.ContainsKey(mP.type ?? "") ? SupportedTypeMap[mP.type] : mP.type;
                if (!string.Equals(normMetaType, cP.type, StringComparison.OrdinalIgnoreCase))
                {
                    res.isMatch = false;
                    res.error = "第 " + (i + 1) + " 个参数类型冲突：.meta.json 声明为 '" + mP.type + "'，源码过程实际为 '" + cP.type + "' (参数名: " + cP.name + ")。";
                    return res;
                }

                // 融合元数据中的描述与约束
                var merged = new ScriptParameterDef
                {
                    name = cP.name, // 优先以源码实名为准
                    type = cP.type,
                    rawType = cP.rawType,
                    description = !string.IsNullOrEmpty(mP.description) ? mP.description : cP.description,
                    isOptional = cP.isOptional,
                    defaultValue = !string.IsNullOrEmpty(mP.defaultValue) ? mP.defaultValue : cP.defaultValue,
                    isSupported = cP.isSupported,
                    unsupportedReason = cP.unsupportedReason
                };
                res.mergedParameters.Add(merged);
            }

            res.isMatch = true;
            return res;
        }

        /// <summary>
        /// 将用户输入的纯文本安全转义为 VBA 字符串表达式，彻底杜绝语句逃逸注入
        /// </summary>
        public static string EscapeVbaString(string input)
        {
            if (input == null) return "\"\"";

            // 统一换行符
            string normalized = input.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split('\n');
            var parts = new List<string>();

            foreach (var line in lines)
            {
                // 双引号转义为 ""
                parts.Add("\"" + line.Replace("\"", "\"\"") + "\"");
            }

            if (parts.Count == 1) return parts[0];
            return string.Join(" & vbCrLf & ", parts.ToArray());
        }

        /// <summary>
        /// 对参数输入值生成脱敏摘要，避免敏感业务数据与完整文本进入运行历史
        /// </summary>
        public static string SanitizeParameterValue(string type, object value)
        {
            if (value == null) return "<null>";
            string strVal = value.ToString();

            if (string.Equals(type, "String", StringComparison.OrdinalIgnoreCase))
            {
                if (strVal.Length > 25)
                {
                    return "\"" + strVal.Substring(0, 20) + "...\" (长度: " + strVal.Length + ")";
                }
                return "\"" + strVal + "\"";
            }
            else if (string.Equals(type, "Worksheet", StringComparison.OrdinalIgnoreCase))
            {
                return "Worksheet(" + strVal + ")";
            }
            else if (string.Equals(type, "Range", StringComparison.OrdinalIgnoreCase))
            {
                return "Range(" + strVal + ")";
            }

            return strVal;
        }
    }
}
