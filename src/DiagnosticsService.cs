using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace LeeExcel
{
    public class DiagnosticsEnvironmentSummary
    {
        public string appName { get; set; }
        public string appVersion { get; set; }
        public string buildCommit { get; set; }
        public string clrVersion { get; set; }
        public string osVersion { get; set; }
        public string osArchitecture { get; set; }
        public string processArchitecture { get; set; }
        public string excelVersion { get; set; }
        public string excelBitness { get; set; }
        public string webView2Version { get; set; }
        public string taskPaneStatus { get; set; }
        public bool hasActiveWorkbook { get; set; }
        public string activeWorkbookMaskedName { get; set; }
        public string lastFailureStage { get; set; }
        public string lastErrorSummary { get; set; }
        public string generatedAtUtc { get; set; }

        public DiagnosticsEnvironmentSummary()
        {
            appName = "ExcelMind AI";
            appVersion = "v1.2.0";
            buildCommit = "231ae3a";
            clrVersion = Environment.Version.ToString();
            osVersion = Environment.OSVersion.ToString();
            osArchitecture = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit";
            processArchitecture = Environment.Is64BitProcess ? "64-bit" : "32-bit";
            excelVersion = "Unknown";
            excelBitness = Environment.Is64BitProcess ? "64-bit" : "32-bit";
            webView2Version = "Unknown";
            taskPaneStatus = "loaded";
            hasActiveWorkbook = false;
            activeWorkbookMaskedName = "None";
            lastFailureStage = "None";
            lastErrorSummary = "None";
            generatedAtUtc = DateTime.UtcNow.ToString("o");
        }
    }

    public class DiagnosticsPreviewResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public DiagnosticsEnvironmentSummary summary { get; set; }
        public List<string> includedFiles { get; set; }
        public List<string> includedCategories { get; set; }
        public List<string> excludedCategories { get; set; }
        public List<string> sanitizedLogPreview { get; set; }
        public long estimatedTotalBytes { get; set; }
        public int sanitizedLogLinesCount { get; set; }
        public int omittedSensitiveLinesCount { get; set; }

        public DiagnosticsPreviewResult()
        {
            includedFiles = new List<string>();
            includedCategories = new List<string>();
            excludedCategories = new List<string>();
            sanitizedLogPreview = new List<string>();
        }
    }

    public class DiagnosticsExportResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string zipFilePath { get; set; }
        public long zipSizeBytes { get; set; }
        public string sha256 { get; set; }
        public List<string> includedFiles { get; set; }
        public int totalLogLines { get; set; }
        public int omittedSensitiveLinesCount { get; set; }

        public DiagnosticsExportResult()
        {
            includedFiles = new List<string>();
        }
    }

    /// <summary>
    /// TASK-R6c-01: 脱敏诊断导出服务
    /// 遵循严格白名单收集、默认排除全部业务与凭据数据、路径与工作簿及 Query 脱敏、正则扫描第二道防线
    /// </summary>
    public static class DiagnosticsService
    {
        private static readonly object _logLock = new object();
        private static readonly List<string> _recentLogs = new List<string>();
        private const int MaxLogLines = 200;

        public static string LastRecordedFailureStage = "None";
        public static string LastRecordedErrorSummary = "None";

        public static void RecordLog(string message, string level = "INFO")
        {
            if (string.IsNullOrEmpty(message)) return;
            lock (_logLock)
            {
                string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string entry = string.Format("[{0}] [{1}] {2}", timestamp, level, message.Trim());
                _recentLogs.Add(entry);
                if (_recentLogs.Count > MaxLogLines)
                {
                    _recentLogs.RemoveAt(0);
                }
            }
        }

        public static void RecordFailure(string stage, string summary)
        {
            LastRecordedFailureStage = string.IsNullOrEmpty(stage) ? "None" : stage;
            LastRecordedErrorSummary = string.IsNullOrEmpty(summary) ? "None" : summary;
            RecordLog(string.Format("FailureRecorded: stage={0}, summary={1}", LastRecordedFailureStage, LastRecordedErrorSummary), "WARN");
        }

        public static string DetectWebView2Version()
        {
            try
            {
                string[] regPaths = new string[]
                {
                    @"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}",
                    @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}"
                };

                foreach (var p in regPaths)
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(p))
                    {
                        if (key != null)
                        {
                            object pv = key.GetValue("pv");
                            if (pv != null && !string.IsNullOrEmpty(pv.ToString()) && pv.ToString() != "0.0.0.0")
                            {
                                return pv.ToString();
                            }
                        }
                    }
                    using (var key = Registry.CurrentUser.OpenSubKey(p))
                    {
                        if (key != null)
                        {
                            object pv = key.GetValue("pv");
                            if (pv != null && !string.IsNullOrEmpty(pv.ToString()) && pv.ToString() != "0.0.0.0")
                            {
                                return pv.ToString();
                            }
                        }
                    }
                }
            }
            catch { }
            return "Not Detected / Embedded";
        }

        public static DiagnosticsEnvironmentSummary BuildEnvironmentSummary(dynamic app)
        {
            var summary = new DiagnosticsEnvironmentSummary();
            summary.webView2Version = DetectWebView2Version();
            summary.lastFailureStage = LastRecordedFailureStage;
            summary.lastErrorSummary = LastRecordedErrorSummary;

            if (app != null)
            {
                try
                {
                    summary.excelVersion = (string)app.Version;
                }
                catch { summary.excelVersion = "COM Available"; }

                try
                {
                    dynamic wb = app.ActiveWorkbook;
                    if (wb != null)
                    {
                        summary.hasActiveWorkbook = true;
                        string wbName = (string)wb.Name;
                        string ext = Path.GetExtension(wbName);
                        summary.activeWorkbookMaskedName = "workbook_***" + (string.IsNullOrEmpty(ext) ? ".xlsx" : ext);
                    }
                }
                catch
                {
                    summary.hasActiveWorkbook = false;
                    summary.activeWorkbookMaskedName = "None";
                }
            }

            return summary;
        }

        public static List<string> GetStandardIncludedCategories()
        {
            return new List<string>
            {
                "SystemEnvironment (操作系统、体系架构、.NET CLR 版本)",
                "HostStatus (Excel 宿主版本、WebView2 运行时状态、任务窗格就绪)",
                "SanitizedDiagnosticsLog (最近过滤脱敏日志，限定 200 行)"
            };
        }

        public static List<string> GetStandardExcludedCategories()
        {
            return new List<string>
            {
                "CredentialsAndApiKeys (大模型 API Key、DPAPI 凭据、私钥)",
                "HttpAuthorizationAndTokens (HTTP 请求头、Bearer Token、Cookie)",
                "ChatHistoryAndPrompts (用户会话聊天记录、提示词正文)",
                "MacroSourceCode (宏库源码 .bas 正文及内部代码)",
                "MacroParametersAndPayloads (宏运行实际参数值、流水线数据载荷)",
                "WorkbookAndCellData (工作簿内容、单元格数据、选区样本)",
                "SnapshotBackups (工作簿物理历史快照备份)",
                "ExternalDataResponses (外部接口原始响应体、CSV/JSON 业务数据)"
            };
        }

        public static string SanitizeLogLine(string line, string userName, ref int omittedCount)
        {
            if (string.IsNullOrEmpty(line)) return "";

            // 第二道防线：检测无法可靠脱敏的高危敏感凭据行，直接排除
            string[] highRiskPatterns = new string[]
            {
                @"-----BEGIN\s+[A-Z\s]+PRIVATE\s+KEY-----",
                @"(?i)(?:password|passwd|pwd)\s*(?::=|=|:)\s*[""'][^""'\r\n]{4,}[""']",
                @"(?i)(?:secret_key|client_secret)\s*(?::=|=|:)\s*[""'][^""'\r\n]{6,}[""']"
            };

            foreach (var hr in highRiskPatterns)
            {
                if (Regex.IsMatch(line, hr))
                {
                    omittedCount++;
                    return "[REDACTED_SENSITIVE_LINE: 包含潜在高危凭据已自动剔除]";
                }
            }

            string sanitized = line;

            // 1. 用户名脱敏
            if (!string.IsNullOrEmpty(userName) && userName.Length > 1)
            {
                sanitized = Regex.Replace(sanitized, Regex.Escape(userName), "<REDACTED_USER>", RegexOptions.IgnoreCase);
            }
            sanitized = Regex.Replace(sanitized, @"[A-Za-z]:\\Users\\[^\s\\/""']+", @"C:\Users\<REDACTED_USER>", RegexOptions.IgnoreCase);

            // 2. 大模型 API Key 脱敏 (如 sk-...)
            sanitized = Regex.Replace(sanitized, @"(?i)sk-[A-Za-z0-9_\-]{20,}", "sk-***");

            // 3. Authorization Bearer 脱敏
            sanitized = Regex.Replace(sanitized, @"(?i)Bearer\s+[A-Za-z0-9_\-\.]{10,}", "Bearer ***");

            // 4. URL 敏感查询参数脱敏 (?token=... 或 &key=...)
            sanitized = Regex.Replace(sanitized, @"(?i)([?&](?:token|key|secret|password|auth|access_token|api_key)=)[^&\s""'<>]+", "$1***");

            // 5. 工作簿真实文件名脱敏
            sanitized = Regex.Replace(sanitized, @"(?i)\b([A-Za-z0-9_\-\u4e00-\u9fa5]{3,})\.(xlsx|xlsm|xlsb|xls)\b", "workbook_***.$2");

            return sanitized;
        }

        public static List<string> PrepareSanitizedLogs(string customLogContent, out int omittedCount)
        {
            omittedCount = 0;
            string currentUserName = Environment.UserName;
            var rawLines = new List<string>();

            lock (_logLock)
            {
                rawLines.AddRange(_recentLogs);
            }

            if (!string.IsNullOrEmpty(customLogContent))
            {
                string[] extraLines = customLogContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                rawLines.AddRange(extraLines);
            }

            // 限制最多最近 200 行
            if (rawLines.Count > MaxLogLines)
            {
                rawLines = rawLines.GetRange(rawLines.Count - MaxLogLines, MaxLogLines);
            }

            var result = new List<string>();
            foreach (var line in rawLines)
            {
                string s = SanitizeLogLine(line, currentUserName, ref omittedCount);
                if (!string.IsNullOrEmpty(s))
                {
                    result.Add(s);
                }
            }

            return result;
        }

        public static DiagnosticsPreviewResult PreviewDiagnostics(dynamic app, string customLogContent = null)
        {
            var res = new DiagnosticsPreviewResult { ok = true };
            try
            {
                res.summary = BuildEnvironmentSummary(app);
                res.includedCategories = GetStandardIncludedCategories();
                res.excludedCategories = GetStandardExcludedCategories();
                res.includedFiles = new List<string>
                {
                    "diagnostics_summary.json",
                    "diagnostics.log",
                    "manifest.json"
                };

                int omittedCount;
                var sanitizedLogs = PrepareSanitizedLogs(customLogContent, out omittedCount);
                res.sanitizedLogLinesCount = sanitizedLogs.Count;
                res.omittedSensitiveLinesCount = omittedCount;

                // 预览取前 10 行
                int previewCount = Math.Min(10, sanitizedLogs.Count);
                for (int i = 0; i < previewCount; i++)
                {
                    res.sanitizedLogPreview.Add(sanitizedLogs[i]);
                }

                // 粗略预估大小（字节）
                long summaryBytes = 1024;
                long logBytes = sanitizedLogs.Count * 80;
                res.estimatedTotalBytes = summaryBytes + logBytes + 512;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.error = "生成诊断预览失败: " + ex.Message;
            }
            return res;
        }

        private static string ComputeFileSha256(string filePath)
        {
            if (!File.Exists(filePath)) return "";
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                var sb = new StringBuilder();
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }

        public static DiagnosticsExportResult ExportDiagnostics(dynamic app, string targetZipPath, string customLogContent = null)
        {
            var res = new DiagnosticsExportResult { ok = false };

            string tempDir = null;
            try
            {
                // 1. 目标文件路径确认与安全防线
                if (string.IsNullOrEmpty(targetZipPath))
                {
                    string diagDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelMindAI", "Diagnostics");
                    if (!Directory.Exists(diagDir)) Directory.CreateDirectory(diagDir);
                    string fileName = string.Format("ExcelMindAI_Diagnostics_{0}.zip", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                    targetZipPath = Path.Combine(diagDir, fileName);
                }

                // 防目录穿越与非法字符
                string fullTargetZipPath = Path.GetFullPath(targetZipPath);
                string targetDir = Path.GetDirectoryName(fullTargetZipPath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                // 2. 创建隔离临时目录
                string stagingBase = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelMindAI", "Temp");
                if (!Directory.Exists(stagingBase)) Directory.CreateDirectory(stagingBase);
                tempDir = Path.Combine(stagingBase, "_diag_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);

                // 3. 构建环境与系统摘要 (diagnostics_summary.json)
                var summary = BuildEnvironmentSummary(app);
                string summaryJson = SimpleJson.Serialize(summary);
                string summaryFile = Path.Combine(tempDir, "diagnostics_summary.json");
                File.WriteAllText(summaryFile, summaryJson, Encoding.UTF8);

                // 4. 构建脱敏日志文本 (diagnostics.log)
                int omittedCount;
                var sanitizedLogs = PrepareSanitizedLogs(customLogContent, out omittedCount);
                string logFile = Path.Combine(tempDir, "diagnostics.log");
                File.WriteAllLines(logFile, sanitizedLogs.ToArray(), Encoding.UTF8);

                // 5. 构建包清单 (manifest.json)
                var manifestData = new Dictionary<string, object>
                {
                    { "schemaVersion", "1.0" },
                    { "packageType", "ExcelMindAI_Diagnostics" },
                    { "generatedAt", DateTime.UtcNow.ToString("o") },
                    { "totalLogLines", sanitizedLogs.Count },
                    { "omittedSensitiveLinesCount", omittedCount },
                    { "includedCategories", GetStandardIncludedCategories() },
                    { "excludedCategories", GetStandardExcludedCategories() },
                    { "files", new List<string> { "diagnostics_summary.json", "diagnostics.log", "manifest.json" } },
                    { "notice", "本诊断包已遵循白名单过滤与脱敏规则，不包含 API Key、会话历史、工作簿数据或宏源码。不自动联网上传，不包含可执行代码。" }
                };
                string manifestJson = SimpleJson.Serialize(manifestData);
                string manifestFile = Path.Combine(tempDir, "manifest.json");
                File.WriteAllText(manifestFile, manifestJson, Encoding.UTF8);

                // 6. 打包为 ZIP 文件
                if (File.Exists(fullTargetZipPath))
                {
                    File.Delete(fullTargetZipPath);
                }
                ZipFile.CreateFromDirectory(tempDir, fullTargetZipPath, CompressionLevel.Optimal, false);

                // 7. 填写导出结果
                var fi = new FileInfo(fullTargetZipPath);
                res.ok = true;
                res.zipFilePath = fullTargetZipPath;
                res.zipSizeBytes = fi.Length;
                res.sha256 = ComputeFileSha256(fullTargetZipPath);
                res.includedFiles = new List<string> { "diagnostics_summary.json", "diagnostics.log", "manifest.json" };
                res.totalLogLines = sanitizedLogs.Count;
                res.omittedSensitiveLinesCount = omittedCount;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.error = "导出诊断包失败: " + ex.Message;
            }
            finally
            {
                // 清理隔离临时目录
                if (!string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
                {
                    try
                    {
                        Directory.Delete(tempDir, true);
                    }
                    catch { }
                }
            }

            return res;
        }
    }
}
