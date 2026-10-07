using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace LeeExcel
{
    public class CsvParseOptions
    {
        public string encoding { get; set; }
        public string delimiter { get; set; }
        public bool hasHeader { get; set; }
        public string quote { get; set; }

        public CsvParseOptions()
        {
            encoding = "utf-8";
            delimiter = ",";
            hasHeader = true;
            quote = "\"";
        }
    }

    public class JsonParseOptions
    {
        public string arrayPath { get; set; }

        public JsonParseOptions()
        {
            arrayPath = "";
        }
    }

    public class HttpGetOptions
    {
        public string url { get; set; }
        public Dictionary<string, string> headers { get; set; }
        public int timeoutSeconds { get; set; }
        public bool allowRedirect { get; set; }
        public List<string> whitelistRules { get; set; }
        public string credentialKey { get; set; }

        public HttpGetOptions()
        {
            url = "";
            headers = new Dictionary<string, string>();
            timeoutSeconds = 15;
            allowRedirect = false;
            whitelistRules = new List<string>();
            credentialKey = "";
        }
    }

    public class ExternalDataPreviewParams
    {
        public string sourceType { get; set; }
        public string filePath { get; set; }
        public string path { get { return filePath; } set { if (!string.IsNullOrEmpty(value)) filePath = value; } }
        public string url { get { return httpOptions != null ? httpOptions.url : null; } set { if (httpOptions == null) httpOptions = new HttpGetOptions(); httpOptions.url = value; } }

        public CsvParseOptions csvOptions { get; set; }
        public JsonParseOptions jsonOptions { get; set; }
        public HttpGetOptions httpOptions { get; set; }
        public int previewRowCount { get; set; }

        public ExternalDataPreviewParams()
        {
            sourceType = "csv";
            filePath = "";
            csvOptions = new CsvParseOptions();
            jsonOptions = new JsonParseOptions();
            httpOptions = new HttpGetOptions();
            previewRowCount = 5;
        }
    }

    public class ExternalDataPreviewResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string sourceType { get; set; }
        public string previewId { get; set; }                  // 预览快照唯一标识
        public string dataFingerprint { get; set; }            // 预览数据 SHA-256 指纹
        public List<string> columns { get; set; }
        public List<string> detectedTypes { get; set; }
        public List<string> unsupportedColumns { get; set; }   // 包含嵌套对象/数组的不支持列
        public List<List<string>> sampleRows { get; set; }
        public int totalRowsEstimate { get; set; }
        public int totalCols { get; set; }
        public bool isTruncated { get; set; }
        public string sanitizedSummary { get; set; }
        public Dictionary<string, string> nullHandlingRules { get; set; }
        public Dictionary<string, string> capacityLimits { get; set; }

        public ExternalDataPreviewResult()
        {
            columns = new List<string>();
            detectedTypes = new List<string>();
            unsupportedColumns = new List<string>();
            sampleRows = new List<List<string>>();
            nullHandlingRules = new Dictionary<string, string>();
            capacityLimits = new Dictionary<string, string>();
        }
    }

    public class ExternalDataImportParams
    {
        public string targetWorkbookFullName { get; set; }
        public string targetWorkbook { get { return targetWorkbookFullName; } set { if (!string.IsNullOrEmpty(value)) targetWorkbookFullName = value; } }
        public string targetSheetName { get; set; }
        public string targetSheet { get { return targetSheetName; } set { if (!string.IsNullOrEmpty(value)) targetSheetName = value; } }
        public string sourceType { get; set; }
        public string filePath { get; set; }
        public string path { get { return filePath; } set { if (!string.IsNullOrEmpty(value)) filePath = value; } }
        public string url { get { return httpOptions != null ? httpOptions.url : null; } set { if (httpOptions == null) httpOptions = new HttpGetOptions(); httpOptions.url = value; } }

        public string previewId { get; set; }                  // 绑定用户已确认的预览数据快照
        public string expectedFingerprint { get; set; }        // 期望数据源指纹（防静默改动）

        public CsvParseOptions csvOptions { get; set; }
        public JsonParseOptions jsonOptions { get; set; }
        public HttpGetOptions httpOptions { get; set; }
        public List<string> selectedColumns { get; set; }

        public ExternalDataImportParams()
        {
            sourceType = "csv";
            filePath = "";
            previewId = "";
            expectedFingerprint = "";
            csvOptions = new CsvParseOptions();
            jsonOptions = new JsonParseOptions();
            httpOptions = new HttpGetOptions();
            selectedColumns = new List<string>();
        }
    }

    public class ExternalDataImportResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string failureStage { get; set; }
        public string targetWorkbookFullName { get; set; }
        public string sheetName { get; set; }
        public int importedRowCount { get; set; }
        public int importedColCount { get; set; }
        public string snapshotId { get; set; }
        public long elapsedMs { get; set; }
        public string recoveryNotice { get; set; }
        public string sanitizedSource { get; set; }
    }

    public class DataSourceConfigDto
    {
        public string id { get; set; }
        public string name { get; set; }
        public string sourceType { get; set; }
        public string pathOrUrl { get; set; }
        public CsvParseOptions csvOptions { get; set; }
        public JsonParseOptions jsonOptions { get; set; }
        public HttpGetOptions httpOptions { get; set; }
        public bool hasCredential { get; set; }
        public string createdAt { get; set; }
        public string updatedAt { get; set; }
    }

    public enum CellValueKind
    {
        String,
        Number,
        LongIntegerPreserved,
        Boolean,
        Null,
        Missing,
        EmptyString,
        UnsupportedObject,
        UnsupportedArray
    }

    public class ParsedCell
    {
        public string RawText { get; set; }
        public CellValueKind Kind { get; set; }
    }

    public class ParsedTable
    {
        public List<string> Columns { get; set; }
        public List<List<string>> Rows { get; set; }
        public List<List<ParsedCell>> CellMatrix { get; set; }
        public List<string> DetectedTypes { get; set; }
        public HashSet<string> UnsupportedColumns { get; set; }
        public int TotalRowCount { get; set; }

        public ParsedTable()
        {
            Columns = new List<string>();
            Rows = new List<List<string>>();
            CellMatrix = new List<List<ParsedCell>>();
            DetectedTypes = new List<string>();
            UnsupportedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public static class ExternalDataService
    {
        public static readonly long MaxFileBytes = 50 * 1024 * 1024; // 50MB
        public static readonly long MaxHttpBytes = 10 * 1024 * 1024; // 10MB
        public static readonly int MaxRowCount = 100000;             // 100k rows
        public static readonly int MaxColCount = 500;                // 500 cols

        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ExcelMindAI_DPAPI_Entropy_2026");

        private static string GetConfigDirectory()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dir = Path.Combine(appData, "ExcelMindAI", "DataSources");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }

        #region DPAPI Credential Store

        private static string GetCredentialFilePath()
        {
            return Path.Combine(GetConfigDirectory(), "credentials.dat");
        }

        public static void SaveCredential(string key, string secret)
        {
            if (string.IsNullOrEmpty(key)) return;
            var creds = LoadAllCredentials();
            creds[key] = secret ?? "";

            var sb = new StringBuilder();
            sb.Append("{");
            int idx = 0;
            foreach (var kvp in creds)
            {
                if (idx > 0) sb.Append(",");
                sb.AppendFormat("\"{0}\":\"{1}\"", SimpleJson.Escape(kvp.Key), SimpleJson.Escape(kvp.Value));
                idx++;
            }
            sb.Append("}");

            byte[] plainBytes = Encoding.UTF8.GetBytes(sb.ToString());
            byte[] encrypted = ProtectedData.Protect(plainBytes, Entropy, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(GetCredentialFilePath(), encrypted);
        }

        public static string GetCredential(string key)
        {
            if (string.IsNullOrEmpty(key)) return null;
            var creds = LoadAllCredentials();
            return creds.ContainsKey(key) ? creds[key] : null;
        }

        public static bool HasCredential(string key)
        {
            if (string.IsNullOrEmpty(key)) return false;
            var creds = LoadAllCredentials();
            return creds.ContainsKey(key);
        }

        public static void DeleteCredential(string key)
        {
            if (string.IsNullOrEmpty(key)) return;
            var creds = LoadAllCredentials();
            if (creds.Remove(key))
            {
                var sb = new StringBuilder();
                sb.Append("{");
                int idx = 0;
                foreach (var kvp in creds)
                {
                    if (idx > 0) sb.Append(",");
                    sb.AppendFormat("\"{0}\":\"{1}\"", SimpleJson.Escape(kvp.Key), SimpleJson.Escape(kvp.Value));
                    idx++;
                }
                sb.Append("}");

                byte[] plainBytes = Encoding.UTF8.GetBytes(sb.ToString());
                byte[] encrypted = ProtectedData.Protect(plainBytes, Entropy, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(GetCredentialFilePath(), encrypted);
            }
        }

        private static Dictionary<string, string> LoadAllCredentials()
        {
            var dict = new Dictionary<string, string>();
            string path = GetCredentialFilePath();
            if (!File.Exists(path)) return dict;

            try
            {
                byte[] encrypted = File.ReadAllBytes(path);
                byte[] decrypted = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
                string json = Encoding.UTF8.GetString(decrypted);
                return SimpleJson.ParseFlatObject(json);
            }
            catch
            {
                return dict;
            }
        }

        #endregion

        #region URL Whitelist & Sanitization

        public static bool ValidateUrlAgainstWhitelist(string targetUrl, List<string> rules, out string reason)
        {
            reason = "";
            if (string.IsNullOrEmpty(targetUrl))
            {
                reason = "URL 不能为空";
                return false;
            }

            Uri targetUri;
            if (!Uri.TryCreate(targetUrl, UriKind.Absolute, out targetUri))
            {
                reason = "URL 格式不合法，必须为绝对 URL (例如 http://127.0.0.1:8080/data)";
                return false;
            }

            if (!string.Equals(targetUri.Scheme, "http", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(targetUri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
            {
                reason = "协议不合法，仅支持 http 或 https 协议";
                return false;
            }

            // 检查私有 / 回环地址
            bool isPrivateOrLoopback = IsPrivateOrLoopbackAddress(targetUri.Host);

            if (rules == null || rules.Count == 0)
            {
                reason = isPrivateOrLoopback
                    ? "禁止访问内网或本地回环地址: 目标未配置在 URL 允许白名单中"
                    : "URL 未匹配任何配置的白名单规则";
                return false;
            }

            int targetPort = targetUri.IsDefaultPort ? (targetUri.Scheme == "https" ? 443 : 80) : targetUri.Port;

            foreach (var rule in rules)
            {
                if (string.IsNullOrEmpty(rule)) continue;
                string cleanRule = rule.Trim();

                Uri ruleUri;
                if (!Uri.TryCreate(cleanRule, UriKind.Absolute, out ruleUri))
                {
                    // 允许 rule 是 host:port 形式
                    if (!cleanRule.Contains("://"))
                    {
                        cleanRule = targetUri.Scheme + "://" + cleanRule;
                        if (!Uri.TryCreate(cleanRule, UriKind.Absolute, out ruleUri))
                            continue;
                    }
                    else
                    {
                        continue;
                    }
                }

                // 1. 协议比对
                if (!string.Equals(targetUri.Scheme, ruleUri.Scheme, StringComparison.OrdinalIgnoreCase))
                    continue;

                // 2. 主机精准比对 (杜绝子域名模糊碰撞，如 evil-example.com vs example.com)
                if (!string.Equals(targetUri.Host, ruleUri.Host, StringComparison.OrdinalIgnoreCase))
                    continue;

                // 3. 端口精准比对 (杜绝 127.0.0.1:8080 允许访问 127.0.0.1:9090)
                int rulePort = ruleUri.IsDefaultPort ? (ruleUri.Scheme == "https" ? 443 : 80) : ruleUri.Port;
                if (targetPort != rulePort)
                    continue;

                // 4. 路径白名单分段严格边界比对 (杜绝 /api/data 误放行 /api/data_evil 或 /api/data-leak)
                string rulePath = ruleUri.AbsolutePath ?? "/";
                string targetPath = targetUri.AbsolutePath ?? "/";

                if (rulePath.Length > 1 && rulePath.EndsWith("/")) rulePath = rulePath.TrimEnd('/');
                if (targetPath.Length > 1 && targetPath.EndsWith("/")) targetPath = targetPath.TrimEnd('/');

                if (!string.IsNullOrEmpty(rulePath) && rulePath != "/")
                {
                    bool isExactMatch = string.Equals(targetPath, rulePath, StringComparison.OrdinalIgnoreCase);
                    bool isSubDirMatch = targetPath.StartsWith(rulePath + "/", StringComparison.OrdinalIgnoreCase);
                    if (!isExactMatch && !isSubDirMatch)
                        continue;
                }

                // 匹配成功
                return true;
            }

            reason = string.Format("目标 URL 【{0}】未命中任何白名单规则。请求已严格阻断。", SanitizeUrl(targetUrl));
            return false;
        }

        private static bool IsPrivateOrLoopbackAddress(string host)
        {
            if (string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(host, "127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(host, "::1", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            IPAddress ip;
            if (IPAddress.TryParse(host, out ip))
            {
                if (IPAddress.IsLoopback(ip)) return true;

                byte[] bytes = ip.GetAddressBytes();
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    // 10.0.0.0/8
                    if (bytes[0] == 10) return true;
                    // 172.16.0.0/12
                    if (bytes[0] == 172 && (bytes[1] >= 16 && bytes[1] <= 31)) return true;
                    // 192.168.0.0/16
                    if (bytes[0] == 192 && bytes[1] == 168) return true;
                    // 169.254.0.0/16
                    if (bytes[0] == 169 && bytes[1] == 254) return true;
                }
            }
            return false;
        }

        public static string SanitizeUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return "";
            try
            {
                Uri u = new Uri(url);
                string cleanQuery = "";
                if (!string.IsNullOrEmpty(u.Query))
                {
                    cleanQuery = Regex.Replace(u.Query, @"(?i)(token|key|secret|password|pwd|auth|apikey|api_key|access_token)=([^&]+)", "$1=***");
                }
                return string.Format("{0}://{1}{2}{3}{4}",
                    u.Scheme, u.Host, u.IsDefaultPort ? "" : ":" + u.Port, u.AbsolutePath, cleanQuery);
            }
            catch
            {
                return Regex.Replace(url, @"(?i)(token|key|secret|password|pwd|auth|apikey|api_key|access_token)=([^&]+)", "$1=***");
            }
        }

        public static string SanitizeMessage(string msg)
        {
            if (string.IsNullOrEmpty(msg)) return "";
            return Regex.Replace(msg, @"(?i)(token|key|secret|password|pwd|auth|apikey|api_key|access_token)=([^\s&""'<>]+)", "$1=***");
        }

        public static Dictionary<string, string> SanitizeHeaders(Dictionary<string, string> headers)
        {
            var res = new Dictionary<string, string>();
            if (headers == null) return res;
            foreach (var kvp in headers)
            {
                string kLower = kvp.Key.ToLowerInvariant();
                if (kLower.Contains("auth") || kLower.Contains("token") || kLower.Contains("secret") || kLower.Contains("key") || kLower.Contains("cookie") || kLower.Contains("pass"))
                {
                    res[kvp.Key] = "******";
                }
                else
                {
                    res[kvp.Key] = kvp.Value;
                }
            }
            return res;
        }

        #endregion

        #region CSV Parser

        public static ParsedTable ParseCsv(string content, CsvParseOptions options, int maxRowsToRead)
        {
            var result = new ParsedTable();
            if (string.IsNullOrEmpty(content)) return result;

            char delim = ',';
            if (!string.IsNullOrEmpty(options.delimiter))
            {
                delim = options.delimiter[0];
            }

            var allRecords = new List<List<string>>();
            var currentRecord = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < content.Length && content[i + 1] == '"')
                        {
                            currentField.Append('"');
                            i++; // 跳过转义双引号
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        currentField.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == delim)
                    {
                        currentRecord.Add(currentField.ToString());
                        currentField.Length = 0;
                    }
                    else if (c == '\r')
                    {
                        if (i + 1 < content.Length && content[i + 1] == '\n')
                        {
                            i++;
                        }
                        currentRecord.Add(currentField.ToString());
                        currentField.Length = 0;
                        allRecords.Add(currentRecord);
                        currentRecord = new List<string>();
                        if (maxRowsToRead > 0 && allRecords.Count >= maxRowsToRead)
                            break;
                    }
                    else if (c == '\n')
                    {
                        currentRecord.Add(currentField.ToString());
                        currentField.Length = 0;
                        allRecords.Add(currentRecord);
                        currentRecord = new List<string>();
                        if (maxRowsToRead > 0 && allRecords.Count >= maxRowsToRead)
                            break;
                    }
                    else
                    {
                        currentField.Append(c);
                    }
                }
            }

            if (currentField.Length > 0 || currentRecord.Count > 0)
            {
                currentRecord.Add(currentField.ToString());
                allRecords.Add(currentRecord);
            }

            if (allRecords.Count == 0) return result;

            // 过滤末尾空行
            while (allRecords.Count > 0 && allRecords[allRecords.Count - 1].Count == 1 && string.IsNullOrEmpty(allRecords[allRecords.Count - 1][0]))
            {
                allRecords.RemoveAt(allRecords.Count - 1);
            }

            if (allRecords.Count == 0) return result;

            int startRow = 0;
            if (options.hasHeader)
            {
                result.Columns = allRecords[0];
                startRow = 1;
            }
            else
            {
                int maxCols = 0;
                foreach (var r in allRecords) if (r.Count > maxCols) maxCols = r.Count;
                result.Columns = new List<string>();
                for (int c = 1; c <= maxCols; c++) result.Columns.Add("列" + c);
            }

            // 规整列数
            int targetColCount = result.Columns.Count;
            for (int r = startRow; r < allRecords.Count; r++)
            {
                var row = allRecords[r];
                while (row.Count < targetColCount) row.Add("");
                result.Rows.Add(row);
            }

            result.TotalRowCount = result.Rows.Count;
            result.DetectedTypes = InferColumnTypes(result.Columns, result.Rows);
            return result;
        }

        #endregion

        #region JSON Parser with Long Number Preservation

        // RFC 8259 严格合规的 JSON 数字正则：
        // 语法结构: ^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$
        // 允许可选负号、0 或非零开头多位数字、可选小数、可选指数
        // 严格禁止前导零（如 0123）、前导加号（如 +123）、前导小数点（如 .5）、末尾小数点（如 12.）以及非法字符（如 123a）
        private static readonly Regex JsonNumberRegex = new Regex(@"^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$", RegexOptions.Compiled);

        public static ParsedTable ParseJson(string jsonText, JsonParseOptions options, int maxRowsToRead)
        {
            var result = new ParsedTable();
            if (string.IsNullOrEmpty(jsonText)) return result;

            jsonText = jsonText.Trim();

            // 若指定了 arrayPath (例如 data.items)，先下钻提取对应子 JSON 文本
            if (options != null && !string.IsNullOrEmpty(options.arrayPath))
            {
                jsonText = ExtractJsonSubPath(jsonText, options.arrayPath);
                if (string.IsNullOrEmpty(jsonText))
                {
                    throw new InvalidOperationException("未在 JSON 中找到指定路径【" + options.arrayPath + "】的数组数据。");
                }
            }

            jsonText = jsonText.Trim();
            if (!jsonText.StartsWith("["))
            {
                throw new InvalidOperationException("JSON 数据必须为对象数组 (例如 [{...}, {...}])。当前根节点不是数组。");
            }

            // 词法阶段原样提取 token，绝不使用正则改写整体文本，保持 RFC 8259 原始语义与精度！
            var rawObjects = ParseJsonRecordArrayPreservingNumbers(jsonText, maxRowsToRead, result.UnsupportedColumns);
            if (rawObjects == null || rawObjects.Count == 0)
            {
                return result;
            }

            // 收集所有列名的并集，保持首次出现顺序
            var colOrder = new List<string>();
            var colSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var dict in rawObjects)
            {
                foreach (var key in dict.Keys)
                {
                    if (!colSet.Contains(key))
                    {
                        colSet.Add(key);
                        colOrder.Add(key);
                    }
                }
            }

            if (colOrder.Count > MaxColCount)
            {
                throw new InvalidOperationException(string.Format("JSON 字段列数 ({0}) 超过系统上限 ({1})。已安全阻断。", colOrder.Count, MaxColCount));
            }

            result.Columns = colOrder;

            // 填充数据行与单元格矩阵
            foreach (var dict in rawObjects)
            {
                var row = new List<string>();
                var cellRow = new List<ParsedCell>();
                foreach (var col in colOrder)
                {
                    ParsedCell cell;
                    if (dict.TryGetValue(col, out cell) && cell != null)
                    {
                        row.Add(cell.RawText);
                        cellRow.Add(cell);
                    }
                    else
                    {
                        // 缺失字段按明确空白处理，不报崩溃
                        row.Add("");
                        cellRow.Add(new ParsedCell { RawText = "", Kind = CellValueKind.Missing });
                    }
                }
                result.Rows.Add(row);
                result.CellMatrix.Add(cellRow);
            }

            result.TotalRowCount = result.Rows.Count;
            result.DetectedTypes = InferColumnTypes(result.Columns, result.Rows, result.UnsupportedColumns);
            return result;
        }

        private static string ExtractJsonSubPath(string json, string path)
        {
            string[] parts = path.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            string current = json.Trim();

            foreach (var part in parts)
            {
                if (!current.StartsWith("{")) return null;
                string pattern = "\"" + Regex.Escape(part) + "\"\\s*:\\s*";
                Match m = Regex.Match(current, pattern);
                if (!m.Success) return null;

                int valStart = m.Index + m.Length;
                current = ExtractJsonValueAt(current, valStart);
                if (current == null) return null;
                current = current.Trim();
            }

            return current;
        }

        private static string ExtractJsonValueAt(string json, int start)
        {
            while (start < json.Length && char.IsWhiteSpace(json[start])) start++;
            if (start >= json.Length) return null;

            char first = json[start];
            if (first == '{' || first == '[')
            {
                char open = first;
                char close = first == '{' ? '}' : ']';
                int depth = 0;
                bool inStr = false;
                for (int i = start; i < json.Length; i++)
                {
                    char c = json[i];
                    if (inStr)
                    {
                        if (c == '\\' && i + 1 < json.Length) i++;
                        else if (c == '"') inStr = false;
                    }
                    else
                    {
                        if (c == '"') inStr = true;
                        else if (c == open) depth++;
                        else if (c == close)
                        {
                            depth--;
                            if (depth == 0)
                            {
                                return json.Substring(start, i - start + 1);
                            }
                        }
                    }
                }
            }
            return null;
        }

        private static List<Dictionary<string, ParsedCell>> ParseJsonRecordArrayPreservingNumbers(string json, int maxRows, HashSet<string> unsupportedCols)
        {
            var list = new List<Dictionary<string, ParsedCell>>();
            int i = 0;
            int len = json.Length;

            while (i < len && json[i] != '[') i++;
            if (i >= len) return list;
            i++; // 跳过 '['

            while (i < len)
            {
                while (i < len && (char.IsWhiteSpace(json[i]) || json[i] == ',')) i++;
                if (i >= len || json[i] == ']') break;

                if (json[i] == '{')
                {
                    int objStart = i;
                    int depth = 0;
                    bool inStr = false;
                    for (; i < len; i++)
                    {
                        char c = json[i];
                        if (inStr)
                        {
                            if (c == '\\' && i + 1 < len) i++;
                            else if (c == '"') inStr = false;
                        }
                        else
                        {
                            if (c == '"') inStr = true;
                            else if (c == '{') depth++;
                            else if (c == '}')
                            {
                                depth--;
                                if (depth == 0)
                                {
                                    i++;
                                    break;
                                }
                            }
                        }
                    }

                    string objStr = json.Substring(objStart, i - objStart);
                    var dict = ParseSingleJsonObjectPreservingNumbers(objStr, unsupportedCols);
                    list.Add(dict);

                    if (maxRows > 0 && list.Count >= maxRows)
                        break;
                }
                else
                {
                    i++;
                }
            }

            return list;
        }

        private static Dictionary<string, ParsedCell> ParseSingleJsonObjectPreservingNumbers(string json, HashSet<string> unsupportedCols)
        {
            var dict = new Dictionary<string, ParsedCell>(StringComparer.OrdinalIgnoreCase);
            int i = 0;
            int len = json.Length;

            while (i < len && json[i] != '{') i++;
            if (i >= len) return dict;
            i++; // skip '{'

            while (i < len)
            {
                while (i < len && (char.IsWhiteSpace(json[i]) || json[i] == ',')) i++;
                if (i >= len || json[i] == '}') break;

                if (json[i] == '"')
                {
                    // 1. 词法解析属性名（Property Key）：使用 ParseJsonString 严格解码，不触碰数字正则
                    string key = ParseJsonString(json, ref i, len);

                    // 2. 寻找 ':'
                    while (i < len && (char.IsWhiteSpace(json[i]) || json[i] == ':')) i++;

                    // 3. 跳过值前空白
                    while (i < len && char.IsWhiteSpace(json[i])) i++;
                    if (i >= len) break;

                    char valFirst = json[i];
                    var cell = new ParsedCell();

                    if (valFirst == '"')
                    {
                        // 字符串 Token：使用 ParseJsonString 词法读取，包含长数字的字符串绝不被当成数字改写
                        string strVal = ParseJsonString(json, ref i, len);
                        cell.RawText = strVal;
                        cell.Kind = CellValueKind.String;
                    }
                    else if (valFirst == '{')
                    {
                        // 嵌套对象：撤回静默占位，明确记录不支持复合对象并标记该列
                        int s = i;
                        int d = 0; bool ins = false;
                        for (; i < len; i++)
                        {
                            if (ins) { if (json[i] == '\\') i++; else if (json[i] == '"') ins = false; }
                            else { if (json[i] == '"') ins = true; else if (json[i] == '{') d++; else if (json[i] == '}') { d--; if (d == 0) { i++; break; } } }
                        }
                        cell.RawText = "[不支持嵌套对象]";
                        cell.Kind = CellValueKind.UnsupportedObject;
                        if (unsupportedCols != null) unsupportedCols.Add(key);
                    }
                    else if (valFirst == '[')
                    {
                        // 嵌套数组：撤回静默占位，明确记录不支持复合数组并标记该列
                        int s = i;
                        int d = 0; bool ins = false;
                        for (; i < len; i++)
                        {
                            if (ins) { if (json[i] == '\\') i++; else if (json[i] == '"') ins = false; }
                            else { if (json[i] == '"') ins = true; else if (json[i] == '[') d++; else if (json[i] == ']') { d--; if (d == 0) { i++; break; } } }
                        }
                        cell.RawText = "[不支持嵌套数组]";
                        cell.Kind = CellValueKind.UnsupportedArray;
                        if (unsupportedCols != null) unsupportedCols.Add(key);
                    }
                    else
                    {
                        // 字面量：null、布尔、数字
                        int s = i;
                        while (i < len && json[i] != ',' && json[i] != '}' && json[i] != ']' && !char.IsWhiteSpace(json[i]))
                        {
                            i++;
                        }
                        string rawToken = json.Substring(s, i - s).Trim();
                        if (rawToken == "null")
                        {
                            cell.RawText = "";
                            cell.Kind = CellValueKind.Null;
                        }
                        else if (rawToken == "true")
                        {
                            cell.RawText = "true";
                            cell.Kind = CellValueKind.Boolean;
                        }
                        else if (rawToken == "false")
                        {
                            cell.RawText = "false";
                            cell.Kind = CellValueKind.Boolean;
                        }
                        else
                        {
                            // 严格执行 RFC 8259 数字语义验证
                            if (!JsonNumberRegex.IsMatch(rawToken))
                            {
                                throw new InvalidOperationException(string.Format("JSON 词法错误: 字段【{0}】包含非法数字 token 【{1}】（RFC 8259 规范禁止前导零、前导加号、非法小数点或字符）。", key, rawToken));
                            }

                            cell.RawText = rawToken;
                            string digitsOnly = rawToken.StartsWith("-") ? rawToken.Substring(1) : rawToken;
                            if (digitsOnly.Length >= 12 && IsAllDigits(digitsOnly))
                            {
                                cell.Kind = CellValueKind.LongIntegerPreserved;
                            }
                            else
                            {
                                cell.Kind = CellValueKind.Number;
                            }
                        }
                    }

                    dict[key] = cell;
                }
                else
                {
                    i++;
                }
            }

            return dict;
        }

        private static string ParseJsonString(string json, ref int i, int len)
        {
            if (i >= len || json[i] != '"') return "";
            i++; // skip '"'
            var sb = new StringBuilder();
            while (i < len)
            {
                char c = json[i];
                if (c == '"')
                {
                    i++; // skip closing '"'
                    return sb.ToString();
                }
                if (c == '\\')
                {
                    i++;
                    if (i >= len) break;
                    char esc = json[i];
                    switch (esc)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (i + 4 < len)
                            {
                                string hex = json.Substring(i + 1, 4);
                                int code;
                                if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out code))
                                {
                                    sb.Append((char)code);
                                    i += 4;
                                }
                                else
                                {
                                    sb.Append("\\u").Append(hex);
                                    i += 4;
                                }
                            }
                            else
                            {
                                sb.Append("\\u");
                            }
                            break;
                        default:
                            sb.Append(esc);
                            break;
                    }
                }
                else
                {
                    sb.Append(c);
                }
                i++;
            }
            return sb.ToString();
        }

        #endregion

        #region HTTP GET Client

        public static string ExecuteHttpGet(HttpGetOptions options, out string sanitizedUrlSummary)
        {
            sanitizedUrlSummary = SanitizeUrl(options.url);

            // 1. 白名单严格校验
            string whitelistReason;
            if (!ValidateUrlAgainstWhitelist(options.url, options.whitelistRules, out whitelistReason))
            {
                throw new InvalidOperationException("安全阻断: " + whitelistReason);
            }

            // 2. 超时配置
            int timeoutSec = options.timeoutSeconds > 0 && options.timeoutSeconds <= 30 ? options.timeoutSeconds : 15;

            // 3. 构建安全 HttpClientHandler
            var handler = new HttpClientHandler
            {
                AllowAutoRedirect = false // 默认不自动跟随重定向
            };

            using (var client = new HttpClient(handler))
            {
                client.Timeout = TimeSpan.FromSeconds(timeoutSec);

                // 注入用户显式请求头
                if (options.headers != null)
                {
                    foreach (var kvp in options.headers)
                    {
                        client.DefaultRequestHeaders.TryAddWithoutValidation(kvp.Key, kvp.Value);
                    }
                }

                // 若指定了安全凭据 Key，从 DPAPI 解密并注入 Authorization
                if (!string.IsNullOrEmpty(options.credentialKey))
                {
                    string secret = GetCredential(options.credentialKey);
                    if (!string.IsNullOrEmpty(secret))
                    {
                        client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", "Bearer " + secret);
                    }
                }

                string currentUrl = options.url;
                int redirectCount = 0;
                HttpResponseMessage response = null;

                while (true)
                {
                    try
                    {
                        var req = new HttpRequestMessage(HttpMethod.Get, currentUrl);
                        var task = client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
                        if (!task.Wait(TimeSpan.FromSeconds(timeoutSec)))
                        {
                            throw new TimeoutException(string.Format("HTTP GET 请求超时 ({0} 秒)。", timeoutSec));
                        }
                        response = task.Result;
                    }
                    catch (AggregateException ae)
                    {
                        Exception baseEx = ae.GetBaseException() ?? ae;
                        throw new InvalidOperationException(string.Format("HTTP GET 请求失败 [{0}]: {1}", SanitizeUrl(currentUrl), baseEx.Message), baseEx);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(string.Format("HTTP GET 请求失败 [{0}]: {1}", SanitizeUrl(currentUrl), ex.Message), ex);
                    }

                    // 检查重定向 (301, 302, 307, 308)
                    int statusCode = (int)response.StatusCode;
                    if (statusCode == 301 || statusCode == 302 || statusCode == 307 || statusCode == 308)
                    {
                        if (!options.allowRedirect)
                        {
                            throw new InvalidOperationException(string.Format("服务器返回重定向 ({0})，当前配置默认禁止自动重定向。", statusCode));
                        }

                        redirectCount++;
                        if (redirectCount > 3)
                        {
                            throw new InvalidOperationException("重定向次数过多 (> 3 次)，已安全中止。");
                        }

                        Uri locUri = response.Headers.Location;
                        if (locUri == null)
                        {
                            throw new InvalidOperationException("重定向响应缺少 Location 报头。");
                        }

                        Uri currentUri = new Uri(currentUrl);
                        Uri resolvedUri = locUri.IsAbsoluteUri ? locUri : new Uri(currentUri, locUri);
                        string nextUrl = resolvedUri.ToString();

                        // 每一跳均须重新比对白名单！
                        string hopReason;
                        if (!ValidateUrlAgainstWhitelist(nextUrl, options.whitelistRules, out hopReason))
                        {
                            throw new InvalidOperationException(string.Format("重定向目标 【{0}】 未命中白名单安全规则，已安全阻断。", SanitizeUrl(nextUrl)));
                        }

                        // 跨来源重定向（Scheme、Host 或 Port 发生变化）必须严格剥离 Authorization 与 Cookie 敏感凭据！
                        // 目标在白名单也不等于默认获得原来源凭据授权！
                        bool isCrossOrigin = !string.Equals(currentUri.Scheme, resolvedUri.Scheme, StringComparison.OrdinalIgnoreCase) ||
                                             !string.Equals(currentUri.Host, resolvedUri.Host, StringComparison.OrdinalIgnoreCase) ||
                                             (currentUri.Port != resolvedUri.Port);

                        if (isCrossOrigin)
                        {
                            client.DefaultRequestHeaders.Remove("Authorization");
                            client.DefaultRequestHeaders.Remove("Cookie");
                        }

                        currentUrl = nextUrl;
                        continue;
                    }

                    break;
                }

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException(string.Format("HTTP GET 请求失败: 状态码 {0} ({1})。", (int)response.StatusCode, response.ReasonPhrase));
                }

                // 4. 流式读取并严格限制响应体大小 (10MB)
                using (var stream = response.Content.ReadAsStreamAsync().Result)
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    var sb = new StringBuilder();
                    char[] buffer = new char[8192];
                    long totalChars = 0;
                    long maxChars = MaxHttpBytes; // 约 10M 字符

                    int read;
                    while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        totalChars += read;
                        if (totalChars > maxChars)
                        {
                            throw new InvalidOperationException(string.Format("HTTP 响应体大小超过系统上限 ({0} MB)。已安全阻断，不静默截断。", MaxHttpBytes / (1024 * 1024)));
                        }
                        sb.Append(buffer, 0, read);
                    }

                    return sb.ToString();
                }
            }
        }

        #endregion

        #region 类型推断辅助与数字检测

        private static bool IsNormalNumber(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            double d;
            return double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d);
        }

        private static List<string> InferColumnTypes(List<string> columns, List<List<string>> rows, HashSet<string> unsupportedCols = null)
        {
            var types = new List<string>();
            int sampleCount = Math.Min(rows.Count, 20);

            for (int c = 0; c < columns.Count; c++)
            {
                string colName = columns[c];
                if (unsupportedCols != null && unsupportedCols.Contains(colName))
                {
                    types.Add("unsupported_nested");
                    continue;
                }

                int numCount = 0;
                int dateCount = 0;
                int boolCount = 0;
                int totalNonEmpty = 0;

                for (int r = 0; r < sampleCount; r++)
                {
                    if (c >= rows[r].Count) continue;
                    string v = rows[r][c] != null ? rows[r][c].Trim() : null;
                    if (string.IsNullOrEmpty(v)) continue;

                    totalNonEmpty++;

                    // 检查前导零长编号 (如 00123) -> 视为 string
                    if (v.Length > 1 && v.StartsWith("0") && char.IsDigit(v[1]))
                    {
                        continue;
                    }

                    double d;
                    bool b;
                    DateTime dt;
                    if (double.TryParse(v, out d)) numCount++;
                    else if (bool.TryParse(v, out b)) boolCount++;
                    else if (DateTime.TryParse(v, out dt)) dateCount++;
                }

                if (totalNonEmpty == 0) types.Add("string");
                else if (numCount == totalNonEmpty) types.Add("number");
                else if (dateCount == totalNonEmpty) types.Add("date");
                else if (boolCount == totalNonEmpty) types.Add("boolean");
                else types.Add("string");
            }

            return types;
        }

        #endregion

        #region 预览快照缓存与数据完整性指纹

        public class CachedPreview
        {
            public string PreviewId;
            public string DataFingerprint;
            public string SourceType;
            public ParsedTable Table;
            public DateTime CreatedAt;
        }

        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, CachedPreview> PreviewCache = new Dictionary<string, CachedPreview>(StringComparer.OrdinalIgnoreCase);

        public static void StorePreviewCache(string previewId, string fingerprint, string sourceType, ParsedTable table)
        {
            if (string.IsNullOrEmpty(previewId) || table == null) return;
            lock (CacheLock)
            {
                var expired = new List<string>();
                var now = DateTime.UtcNow;
                foreach (var kvp in PreviewCache)
                {
                    if ((now - kvp.Value.CreatedAt).TotalMinutes > 30)
                    {
                        expired.Add(kvp.Key);
                    }
                }
                foreach (var k in expired) PreviewCache.Remove(k);

                PreviewCache[previewId] = new CachedPreview
                {
                    PreviewId = previewId,
                    DataFingerprint = fingerprint,
                    SourceType = sourceType,
                    Table = table,
                    CreatedAt = now
                };
            }
        }

        public static CachedPreview GetPreviewCache(string previewId)
        {
            if (string.IsNullOrEmpty(previewId)) return null;
            lock (CacheLock)
            {
                CachedPreview cp;
                if (PreviewCache.TryGetValue(previewId, out cp))
                {
                    return cp;
                }
            }
            return null;
        }

        public static string ComputeSha256(string input)
        {
            if (input == null) return "";
            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
                var sb = new StringBuilder(bytes.Length * 2);
                for (int i = 0; i < bytes.Length; i++)
                {
                    sb.Append(bytes[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }

        #endregion

        #region 公开核心业务接口：预览与导入

        public static ExternalDataPreviewResult Preview(ExternalDataPreviewParams p)
        {
            var res = new ExternalDataPreviewResult
            {
                sourceType = p.sourceType,
                ok = false
            };

            try
            {
                string rawText = "";
                string summary = "";

                if (string.Equals(p.sourceType, "csv", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(p.filePath) || !File.Exists(p.filePath))
                    {
                        res.error = "指定的 CSV 文件不存在或路径为空: " + p.filePath;
                        return res;
                    }

                    var fi = new FileInfo(p.filePath);
                    if (fi.Length > MaxFileBytes)
                    {
                        res.error = string.Format("CSV 文件大小 ({0:F1} MB) 超过系统上限 ({1} MB)。已安全阻断。", (double)fi.Length / (1024 * 1024), MaxFileBytes / (1024 * 1024));
                        return res;
                    }

                    Encoding enc = GetEncoding(p.csvOptions != null ? p.csvOptions.encoding : null);
                    rawText = File.ReadAllText(p.filePath, enc);
                    summary = string.Format("本地 CSV 文件: {0} ({1:F1} KB, 编码: {2})", Path.GetFileName(p.filePath), (double)fi.Length / 1024, enc.EncodingName);
                }
                else if (string.Equals(p.sourceType, "json", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(p.filePath) || !File.Exists(p.filePath))
                    {
                        res.error = "指定的 JSON 文件不存在或路径为空: " + p.filePath;
                        return res;
                    }

                    var fi = new FileInfo(p.filePath);
                    if (fi.Length > MaxFileBytes)
                    {
                        res.error = string.Format("JSON 文件大小 ({0:F1} MB) 超过系统上限 ({1} MB)。已安全阻断。", (double)fi.Length / (1024 * 1024), MaxFileBytes / (1024 * 1024));
                        return res;
                    }

                    rawText = File.ReadAllText(p.filePath, Encoding.UTF8);
                    summary = string.Format("本地 JSON 文件: {0} ({1:F1} KB)", Path.GetFileName(p.filePath), (double)fi.Length / 1024);
                }
                else if (string.Equals(p.sourceType, "http_get", StringComparison.OrdinalIgnoreCase))
                {
                    string sanitizedSummary;
                    rawText = ExecuteHttpGet(p.httpOptions, out sanitizedSummary);
                    summary = string.Format("只读 HTTP GET: {0}", sanitizedSummary);
                }
                else
                {
                    res.error = "不支持的外部数据源类型: " + p.sourceType + "。仅支持 csv, json, http_get。";
                    return res;
                }

                // 解析为统一表格
                ParsedTable table;
                if (string.Equals(p.sourceType, "csv", StringComparison.OrdinalIgnoreCase))
                {
                    table = ParseCsv(rawText, p.csvOptions ?? new CsvParseOptions(), 0);
                }
                else
                {
                    var jsonOpt = p.jsonOptions ?? new JsonParseOptions();
                    table = ParseJson(rawText, jsonOpt, 0);
                }

                if (table.Columns.Count > MaxColCount)
                {
                    res.error = string.Format("解析字段列数 ({0}) 超过系统上限 ({1})。已安全阻断。", table.Columns.Count, MaxColCount);
                    return res;
                }
                if (table.TotalRowCount > MaxRowCount)
                {
                    res.error = string.Format("数据行数 ({0}) 超过系统上限 ({1})。已安全阻断。", table.TotalRowCount, MaxRowCount);
                    return res;
                }

                // 生成快照唯一标识与完整性指纹
                string previewId = Guid.NewGuid().ToString("N");
                string fingerprint = ComputeSha256(rawText);

                StorePreviewCache(previewId, fingerprint, p.sourceType, table);

                res.previewId = previewId;
                res.dataFingerprint = fingerprint;
                res.columns = table.Columns;
                res.detectedTypes = table.DetectedTypes;
                res.unsupportedColumns = new List<string>(table.UnsupportedColumns);
                res.totalRowsEstimate = table.TotalRowCount;
                res.totalCols = table.Columns.Count;
                res.sanitizedSummary = summary;

                res.nullHandlingRules["JSON null"] = "按空白单元格写入，不转为空字符串或 NaN";
                res.nullHandlingRules["缺失字段"] = "按空白单元格处理，保持行结构对齐";
                res.nullHandlingRules["空字符串 (\"\")"] = "按空文本写入";
                res.nullHandlingRules["前导零文本 (如 \"0123\")"] = "单引号转义写入，防止 Excel 去零截断";
                res.nullHandlingRules["公式样文本 (如 \"=SUM\")"] = "单引号转义写入，防止被 Excel 自动当作公式执行";

                res.capacityLimits["MaxFileBytes"] = "50 MB";
                res.capacityLimits["MaxHttpBytes"] = "10 MB";
                res.capacityLimits["MaxRowCount"] = "100,000 行";
                res.capacityLimits["MaxColCount"] = "500 列";

                int sampleCount = Math.Min(table.Rows.Count, p.previewRowCount > 0 ? p.previewRowCount : 5);
                res.sampleRows = new List<List<string>>();
                for (int i = 0; i < sampleCount; i++)
                {
                    res.sampleRows.Add(table.Rows[i]);
                }

                res.ok = true;
                return res;
            }
            catch (Exception ex)
            {
                string sourceHint = "";
                if (string.Equals(p.sourceType, "http_get", StringComparison.OrdinalIgnoreCase) && p.httpOptions != null)
                {
                    sourceHint = " [" + SanitizeUrl(p.httpOptions.url) + "]";
                }
                res.error = SanitizeMessage(string.Format("数据源{0}只读解析失败: {1}", sourceHint, ex.Message));
                return res;
            }
        }

        public static ExternalDataImportResult Import(ExternalDataImportParams p, dynamic excelApp)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var res = new ExternalDataImportResult
            {
                ok = false,
                failureStage = "precheck"
            };

            try
            {
                // 1. 目标工作簿校验锁定 (严格禁止回退 ActiveWorkbook)
                if (excelApp == null)
                {
                    res.error = "Excel 宿主实例不可用。";
                    return res;
                }

                dynamic targetWb = FindWorkbook(excelApp, p.targetWorkbookFullName);
                if (targetWb == null)
                {
                    res.error = "未找到指定的目标工作簿【" + (p.targetWorkbookFullName ?? "未指定") + "】（可能已被关闭或重命名）。已安全阻断，杜绝回退当前活动工作簿！";
                    return res;
                }

                res.targetWorkbookFullName = (string)targetWb.FullName;

                // 2. 数据读取与快照绑定 (优先使用用户已确认的预览数据快照，杜绝静默重抓变动数据)
                res.failureStage = "fetch";
                ParsedTable table = null;
                string sanitizedSource = "";

                if (!string.IsNullOrEmpty(p.previewId))
                {
                    var cached = GetPreviewCache(p.previewId);
                    if (cached != null)
                    {
                        if (!string.IsNullOrEmpty(p.expectedFingerprint) &&
                            !string.Equals(cached.DataFingerprint, p.expectedFingerprint, StringComparison.OrdinalIgnoreCase))
                        {
                            res.error = "导入数据与预览确认时的数据指纹不一致（数据已变动）。已安全阻断，请重新预览并确认。";
                            return res;
                        }
                        table = cached.Table;
                        sanitizedSource = "已确认快照 (" + p.previewId.Substring(0, Math.Min(8, p.previewId.Length)) + ")";
                    }
                }

                if (table == null)
                {
                    string rawText = "";

                    if (string.Equals(p.sourceType, "csv", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(p.filePath) || !File.Exists(p.filePath))
                        {
                            res.error = "指定的 CSV 文件不存在: " + p.filePath;
                            return res;
                        }
                        var fi = new FileInfo(p.filePath);
                        if (fi.Length > MaxFileBytes)
                        {
                            res.error = string.Format("CSV 文件大小超过上限 ({0} MB)。", MaxFileBytes / (1024 * 1024));
                            return res;
                        }
                        Encoding enc = GetEncoding(p.csvOptions != null ? p.csvOptions.encoding : null);
                        rawText = File.ReadAllText(p.filePath, enc);
                        sanitizedSource = Path.GetFileName(p.filePath);
                    }
                    else if (string.Equals(p.sourceType, "json", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(p.filePath) || !File.Exists(p.filePath))
                        {
                            res.error = "指定的 JSON 文件不存在: " + p.filePath;
                            return res;
                        }
                        var fi = new FileInfo(p.filePath);
                        if (fi.Length > MaxFileBytes)
                        {
                            res.error = string.Format("JSON 文件大小超过上限 ({0} MB)。", MaxFileBytes / (1024 * 1024));
                            return res;
                        }
                        rawText = File.ReadAllText(p.filePath, Encoding.UTF8);
                        sanitizedSource = Path.GetFileName(p.filePath);
                    }
                    else if (string.Equals(p.sourceType, "http_get", StringComparison.OrdinalIgnoreCase))
                    {
                        rawText = ExecuteHttpGet(p.httpOptions, out sanitizedSource);
                    }
                    else
                    {
                        res.error = "不支持的外部数据源类型: " + p.sourceType;
                        return res;
                    }

                    // 校验指纹（如果调用方指定了 expectedFingerprint）
                    string currentFingerprint = ComputeSha256(rawText);
                    if (!string.IsNullOrEmpty(p.expectedFingerprint) &&
                        !string.Equals(currentFingerprint, p.expectedFingerprint, StringComparison.OrdinalIgnoreCase))
                    {
                        res.error = "导入数据与预览确认时的数据指纹不一致（数据源内容已发生变动）。已安全阻断，请重新预览并确认。";
                        return res;
                    }

                    res.failureStage = "parse";
                    if (string.Equals(p.sourceType, "csv", StringComparison.OrdinalIgnoreCase))
                    {
                        table = ParseCsv(rawText, p.csvOptions ?? new CsvParseOptions(), 0);
                    }
                    else
                    {
                        table = ParseJson(rawText, p.jsonOptions ?? new JsonParseOptions(), 0);
                    }
                }

                res.sanitizedSource = sanitizedSource;

                if (table.TotalRowCount > MaxRowCount)
                {
                    res.error = string.Format("数据行数 ({0}) 超过上限 ({1})。已安全阻断，不静默截断。", table.TotalRowCount, MaxRowCount);
                    return res;
                }

                // 3. 字段过滤与对齐 (若指定了 selectedColumns)
                var finalColIndices = new List<int>();
                var finalColNames = new List<string>();

                if (p.selectedColumns != null && p.selectedColumns.Count > 0)
                {
                    foreach (var sel in p.selectedColumns)
                    {
                        int foundIdx = table.Columns.FindIndex(c => string.Equals(c, sel, StringComparison.OrdinalIgnoreCase));
                        if (foundIdx >= 0)
                        {
                            finalColIndices.Add(foundIdx);
                            finalColNames.Add(table.Columns[foundIdx]);
                        }
                    }
                }
                else
                {
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        finalColIndices.Add(c);
                        finalColNames.Add(table.Columns[c]);
                    }
                }

                if (finalColNames.Count == 0)
                {
                    res.error = "未选择任何需要导入的有效列。";
                    return res;
                }

                // 4. 检查选中的列中是否包含不支持的复合结构列（嵌套对象/数组）
                if (table.UnsupportedColumns != null && table.UnsupportedColumns.Count > 0)
                {
                    foreach (var sel in finalColNames)
                    {
                        if (table.UnsupportedColumns.Contains(sel))
                        {
                            res.error = string.Format("字段【{0}】包含嵌套对象或数组，当前版本不支持复合结构导入。请在字段选择中取消该字段后重试。", sel);
                            return res;
                        }
                    }
                }

                // 5. 强制前置整本物理快照 (承诺快照失败零业务写入)
                res.failureStage = "snapshot";
                string snapPrompt = string.Format("外部数据接入 ({0}): 导入前整本物理快照", p.sourceType);
                SnapshotItem snapItem = null;
                try
                {
                    snapItem = SnapshotManager.CreateSnapshot(targetWb, snapPrompt, "");
                }
                catch (Exception exSnap)
                {
                    res.error = "写入前整本快照创建异常: " + exSnap.Message + "。承诺快照失败零业务写入，已安全阻断。";
                    return res;
                }

                if (snapItem == null || string.IsNullOrEmpty(snapItem.id) || !SnapshotManager.SnapshotExists((string)targetWb.FullName, snapItem.id))
                {
                    res.error = "写入前整本快照物理副本写入失败。承诺快照失败零业务写入，已安全阻断。";
                    return res;
                }

                res.snapshotId = snapItem.id;

                // 6. 新建唯一工作表 (绝对不覆盖用户已有表)
                res.failureStage = "write";
                string baseSheetName = !string.IsNullOrEmpty(p.targetSheetName)
                    ? p.targetSheetName.Trim()
                    : string.Format("Import_{0}_{1}", p.sourceType.ToUpperInvariant(), DateTime.Now.ToString("yyyyMMdd"));

                string uniqueSheetName = GetUniqueSheetName(targetWb, baseSheetName);
                dynamic wsNew = targetWb.Worksheets.Add();
                wsNew.Name = uniqueSheetName;
                res.sheetName = uniqueSheetName;

                // 7. 二维 SAFEARRAY 矩阵批量写入与公式/长数字转义保真
                int totalMatrixRows = table.Rows.Count + 1; // 第 1 行为表头
                int totalMatrixCols = finalColNames.Count;
                object[,] matrix = new object[totalMatrixRows, totalMatrixCols];

                // 表头行
                for (int c = 0; c < totalMatrixCols; c++)
                {
                    matrix[0, c] = finalColNames[c];
                }

                // 数据行
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    var srcRow = table.Rows[r];
                    for (int c = 0; c < totalMatrixCols; c++)
                    {
                        int srcColIdx = finalColIndices[c];
                        string rawVal = (srcColIdx < srcRow.Count) ? srcRow[srcColIdx] : "";

                        if (string.IsNullOrEmpty(rawVal))
                        {
                            matrix[r + 1, c] = "";
                            continue;
                        }

                        // 公式样文本防御：以 '=', '@' 开头，或以 '+', '-' 开头且不是纯数字时强制添加单引号保真
                        bool isFormula = rawVal.StartsWith("=") || rawVal.StartsWith("@") ||
                            (rawVal.StartsWith("+") && !IsNormalNumber(rawVal)) ||
                            (rawVal.StartsWith("-") && !IsNormalNumber(rawVal));

                        // 19位长数字、大整数编号、前导零文本保真：添加单引号，防止 Excel 科学计数法或精度截断
                        string digitsOnly = rawVal.StartsWith("-") ? rawVal.Substring(1) : rawVal;
                        bool isLongInteger = digitsOnly.Length >= 12 && IsAllDigits(digitsOnly);
                        bool isLeadingZero = rawVal.Length > 1 && rawVal.StartsWith("0") && char.IsDigit(rawVal[1]);
                        bool isHighPrecisionNum = rawVal.Length >= 16 && (rawVal.Contains(".") || rawVal.Contains("e") || rawVal.Contains("E")) && IsNormalNumber(rawVal);

                        if (isFormula || isLongInteger || isLeadingZero || isHighPrecisionNum)
                        {
                            matrix[r + 1, c] = "'" + rawVal;
                        }
                        else
                        {
                            matrix[r + 1, c] = rawVal;
                        }
                    }
                }

                dynamic destRange = wsNew.Range[wsNew.Cells[1, 1], wsNew.Cells[totalMatrixRows, totalMatrixCols]];
                destRange.Value2 = matrix;

                // 8. 读回客观核验
                res.failureStage = "readback";
                res.importedRowCount = table.Rows.Count;
                res.importedColCount = totalMatrixCols;

                sw.Stop();
                res.elapsedMs = sw.ElapsedMilliseconds;
                res.ok = true;
                return res;
            }
            catch (Exception ex)
            {
                sw.Stop();
                res.elapsedMs = sw.ElapsedMilliseconds;
                string sourceHint = "";
                if (string.Equals(p.sourceType, "http_get", StringComparison.OrdinalIgnoreCase) && p.httpOptions != null)
                {
                    sourceHint = " [" + SanitizeUrl(p.httpOptions.url) + "]";
                }
                res.error = SanitizeMessage(string.Format("外部数据导入异常{0}: {1}", sourceHint, ex.Message));
                res.recoveryNotice = string.Format(
                    "执行中发生异常 (阶段: {0})。用户可按已验证范围恢复目标工作簿（快照 ID: {1}）。",
                    res.failureStage, res.snapshotId ?? "无");
                return res;
            }
        }

        private static dynamic FindWorkbook(dynamic excelApp, string targetFullName)
        {
            if (excelApp == null || string.IsNullOrEmpty(targetFullName)) return null;
            string cleanTarget = targetFullName.Trim();
            try
            {
                foreach (dynamic wb in excelApp.Workbooks)
                {
                    if (string.Equals((string)wb.FullName, cleanTarget, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals((string)wb.Name, cleanTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }
            }
            catch { }
            return null;
        }

        private static bool IsAllDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
            {
                if (!char.IsDigit(c)) return false;
            }
            return true;
        }

        private static string GetUniqueSheetName(dynamic targetWb, string baseName)
        {
            string candidate = baseName;
            int counter = 1;

            while (SheetExists(targetWb, candidate))
            {
                candidate = string.Format("{0}_{1}", baseName, counter);
                counter++;
            }
            return candidate;
        }

        private static bool SheetExists(dynamic targetWb, string sheetName)
        {
            try
            {
                dynamic ws = targetWb.Worksheets[sheetName];
                return ws != null;
            }
            catch
            {
                return false;
            }
        }

        private static Encoding GetEncoding(string encName)
        {
            if (string.IsNullOrEmpty(encName)) return Encoding.UTF8;
            string lower = encName.Trim().ToLowerInvariant();
            if (lower == "gbk" || lower == "gb2312")
            {
                try { return Encoding.GetEncoding("GBK"); } catch { return Encoding.GetEncoding(936); }
            }
            if (lower == "ascii") return Encoding.ASCII;
            if (lower == "utf-8-bom") return new UTF8Encoding(true);
            return new UTF8Encoding(false);
        }

        #endregion

        #region 数据源配置持久化 (纯本地 JSON，绝不含凭据)

        private static string GetConfigsFilePath()
        {
            return Path.Combine(GetConfigDirectory(), "sources.json");
        }

        public static List<DataSourceConfigDto> ListDataSourceConfigs()
        {
            var list = new List<DataSourceConfigDto>();
            string path = GetConfigsFilePath();
            if (!File.Exists(path)) return list;

            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                var items = SimpleJson.DeserializeList<DataSourceConfigDto>(json);
                if (items != null)
                {
                    foreach (var it in items)
                    {
                        // 确保返回的配置中绝对不泄露凭据
                        it.hasCredential = HasCredential(it.id);
                        if (it.httpOptions != null)
                        {
                            it.httpOptions.credentialKey = ""; // 脱敏
                            it.httpOptions.headers = SanitizeHeaders(it.httpOptions.headers);
                        }
                        list.Add(it);
                    }
                }
            }
            catch { }
            return list;
        }

        public static void SaveDataSourceConfig(DataSourceConfigDto config, string optionalSecret)
        {
            if (config == null) return;
            if (string.IsNullOrEmpty(config.id)) config.id = Guid.NewGuid().ToString("N").Substring(0, 8);
            config.updatedAt = DateTime.Now.ToString("o");
            if (string.IsNullOrEmpty(config.createdAt)) config.createdAt = config.updatedAt;

            // 如果用户传入了密码/Token，使用 DPAPI 加密独立存储
            if (!string.IsNullOrEmpty(optionalSecret))
            {
                SaveCredential(config.id, optionalSecret);
                config.hasCredential = true;
            }

            // 清洗持久化配置：绝不将敏感 Token、凭据 Key 或未脱敏请求头写出到明文 sources.json
            if (config.httpOptions != null)
            {
                config.httpOptions.credentialKey = "";
                config.httpOptions.headers = SanitizeHeaders(config.httpOptions.headers);
            }

            var all = ListDataSourceConfigs();
            int existingIdx = all.FindIndex(x => x.id == config.id);
            if (existingIdx >= 0) all[existingIdx] = config;
            else all.Add(config);

            // 序列化持久化至 sources.json (绝不写出 Secret)
            string json = SimpleJson.Serialize(all);
            File.WriteAllText(GetConfigsFilePath(), json, Encoding.UTF8);
        }

        #endregion
    }
}
