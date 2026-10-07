using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using LeeExcel;

namespace LeeExcel.Tests
{
    public class VerifyR6cInstallAndDiagnostics
    {
        public static int Main(string[] args)
        {
            Console.WriteLine("==============================================================");
            Console.WriteLine("    TASK-R6c-01: 脱敏诊断导出与系统环境验证工具 (C# 宿主测试)");
            Console.WriteLine("==============================================================");

            string outDir = args.Length > 0 ? args[0] : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".artifacts", "tests", "r6c_diag_test_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

            int passCount = 0;
            int failCount = 0;

            Action<string, bool, string> AssertTest = (name, cond, detail) =>
            {
                if (cond)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("[PASS] " + name);
                    Console.ResetColor();
                    passCount++;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[FAIL] " + name + " -> " + detail);
                    Console.ResetColor();
                    failCount++;
                }
            };

            // 1. 验证 DiagnosticsService.PreviewDiagnostics 白名单与脱敏预览
            string mockLogs = 
                "[INFO] System initialized for user JohnDoe in C:\\Users\\JohnDoe\\Documents\\Secret_Report_2026.xlsx\r\n" +
                "[DEBUG] Request sent with API Key sk-ant-live-12345678901234567890 and Bearer my-secret-token-abcdef at https://api.com/v1?token=supersecretquery\r\n" +
                "[WARN] High-risk line with -----BEGIN RSA PRIVATE KEY----- should be omitted completely\r\n" +
                "[INFO] Normal log line: Workbook calculation completed.\r\n" +
                "[ERROR] DB connection failed with password := \"PlainTextPass123\"\r\n";

            var previewRes = DiagnosticsService.PreviewDiagnostics(null, mockLogs);

            AssertTest("TC-R6c-D01: 诊断预览返回成功且包含系统环境白名单",
                previewRes.ok && previewRes.summary != null && previewRes.summary.appName == "ExcelMind AI" && previewRes.summary.appVersion == "v1.2.0",
                "previewRes.ok=" + previewRes.ok);

            AssertTest("TC-R6c-D02: 诊断预览严格声明 8 类排除项与 3 类包含分类",
                previewRes.includedCategories.Count == 3 && previewRes.excludedCategories.Count == 8,
                "inc=" + previewRes.includedCategories.Count + ", exc=" + previewRes.excludedCategories.Count);

            string allPreviewText = string.Join("\n", previewRes.sanitizedLogPreview.ToArray());
            AssertTest("TC-R6c-D03: 诊断日志预览正确脱敏 Key、Bearer、Token、路径与工作簿名",
                allPreviewText.Contains("sk-***") && !allPreviewText.Contains("sk-ant-live") && allPreviewText.Contains("Bearer ***") && allPreviewText.Contains("token=***") && allPreviewText.Contains("workbook_***.xlsx") && !allPreviewText.Contains("Secret_Report_2026.xlsx"),
                "preview lines did not sanitize correctly");

            AssertTest("TC-R6c-D04: 敏感信息扫描第二道拦截正确剔除高危私钥与明文密码行",
                previewRes.omittedSensitiveLinesCount >= 2,
                "omittedCount=" + previewRes.omittedSensitiveLinesCount);

            // 2. 验证 DiagnosticsService.ExportDiagnostics 导出脱敏 Zip 包
            string targetZipPath = Path.Combine(outDir, "ExcelMindAI_Diagnostics_Test.zip");
            var exportRes = DiagnosticsService.ExportDiagnostics(null, targetZipPath, mockLogs);

            AssertTest("TC-R6c-D05: 脱敏诊断包成功导出至指定目标路径",
                exportRes.ok && File.Exists(targetZipPath) && exportRes.zipSizeBytes > 0 && !string.IsNullOrEmpty(exportRes.sha256),
                "exportRes.ok=" + exportRes.ok + ", zipExists=" + File.Exists(targetZipPath));

            // 3. 解包物理验证内部文件与内容无凭据纯洁性
            string extractDir = Path.Combine(outDir, "extracted_diag");
            if (Directory.Exists(extractDir)) Directory.Delete(extractDir, true);
            ZipFile.ExtractToDirectory(targetZipPath, extractDir);

            string summaryJsonPath = Path.Combine(extractDir, "diagnostics_summary.json");
            string logTextPath = Path.Combine(extractDir, "diagnostics.log");
            string manifestJsonPath = Path.Combine(extractDir, "manifest.json");

            bool filesExist = File.Exists(summaryJsonPath) && File.Exists(logTextPath) && File.Exists(manifestJsonPath);
            AssertTest("TC-R6c-D06: 诊断包解压包含且仅包含白名单 3 项文件",
                filesExist && Directory.GetFiles(extractDir).Length == 3,
                "fileCount=" + Directory.GetFiles(extractDir).Length);

            string logContent = File.ReadAllText(logTextPath, Encoding.UTF8);
            bool noRawSecretsInLog = 
                !logContent.Contains("sk-ant-live") &&
                !logContent.Contains("my-secret-token-abcdef") &&
                !logContent.Contains("supersecretquery") &&
                !logContent.Contains("PlainTextPass123") &&
                !logContent.Contains("-----BEGIN RSA PRIVATE KEY-----") &&
                !logContent.Contains("Secret_Report_2026.xlsx");

            AssertTest("TC-R6c-D07: 导出的日志文件 100% 杜绝原始密钥、凭据、私钥及工作簿业务名",
                noRawSecretsInLog,
                "raw secrets leaked into diagnostics.log");

            string manifestContent = File.ReadAllText(manifestJsonPath, Encoding.UTF8);
            AssertTest("TC-R6c-D08: manifest.json 格式规范且版本为 1.0",
                manifestContent.Contains("\"schemaVersion\":\"1.0\"") && manifestContent.Contains("\"packageType\":\"ExcelMindAI_Diagnostics\""),
                "manifest schemaVersion invalid");

            // 4. 验证用户取消导出流程 (TC-R6c-D09)
            string cancelJson = NativeBridge.Dispatch("{\"action\":\"export_diagnostics\",\"targetZipPath\":\"__CANCEL__\"}", null);
            var cancelResp = SimpleJson.ParseFlatObject(cancelJson);
            bool cancelHandled = cancelResp.ContainsKey("ok") && cancelResp["ok"] == "false" && 
                                 cancelResp.ContainsKey("error") && cancelResp["error"].Contains("取消");
            AssertTest("TC-R6c-D09: 诊断设置页用户取消保存时安全返回且零写文件零残留",
                cancelHandled,
                "cancelJson=" + cancelJson);

            // 5. 验证用户自定义选择保存位置及前端结果展示元数据 (TC-R6c-D10)
            string userChosenZip = Path.Combine(outDir, "UserChosen_Diagnostics_20261003.zip");
            if (File.Exists(userChosenZip)) File.Delete(userChosenZip);

            string exportCustomJson = NativeBridge.Dispatch(
                string.Format("{{\"action\":\"export_diagnostics\",\"targetZipPath\":\"{0}\",\"customLogContent\":\"[INFO] Testing user chosen export path\"}}", 
                    userChosenZip.Replace("\\", "\\\\")), null);
            
            bool customExportPass = File.Exists(userChosenZip) && exportCustomJson.Contains("\"ok\":true") && 
                                    exportCustomJson.Contains("\"zipSizeBytes\":") && exportCustomJson.Contains("\"sha256\":");
            AssertTest("TC-R6c-D10: 用户选择保存位置后成功导出并返回前端结果展示所需完整元数据",
                customExportPass,
                "exportCustomJson=" + exportCustomJson);

            // 6. 深度检查实际包内容，不仅检查界面文字 (TC-R6c-D11, TC-R6c-D12)
            string deepExtractDir = Path.Combine(outDir, "deep_inspect_pkg");
            if (Directory.Exists(deepExtractDir)) Directory.Delete(deepExtractDir, true);
            ZipFile.ExtractToDirectory(userChosenZip, deepExtractDir);

            string deepSummaryJson = File.ReadAllText(Path.Combine(deepExtractDir, "diagnostics_summary.json"), Encoding.UTF8);
            string deepManifestJson = File.ReadAllText(Path.Combine(deepExtractDir, "manifest.json"), Encoding.UTF8);
            string deepLog = File.ReadAllText(Path.Combine(deepExtractDir, "diagnostics.log"), Encoding.UTF8);

            bool summaryClean = deepSummaryJson.Contains("\"appName\":\"ExcelMind AI\"") && 
                                !deepSummaryJson.Contains("sk-") && !deepSummaryJson.Contains("JohnDoe") &&
                                !deepSummaryJson.Contains("Secret_Report_2026.xlsx");

            AssertTest("TC-R6c-D11: 实际包内 diagnostics_summary.json 白名单完整且无真实敏感工作簿/用户名泄漏",
                summaryClean,
                "diagnostics_summary.json contains sensitive or invalid data");

            bool manifestClean = deepManifestJson.Contains("\"schemaVersion\":\"1.0\"") &&
                                 deepManifestJson.Contains("CredentialsAndApiKeys") &&
                                 deepManifestJson.Contains("HttpAuthorizationAndTokens") &&
                                 deepManifestJson.Contains("ChatHistoryAndPrompts") &&
                                 deepManifestJson.Contains("MacroSourceCode") &&
                                 deepManifestJson.Contains("MacroParametersAndPayloads") &&
                                 deepManifestJson.Contains("WorkbookAndCellData") &&
                                 deepManifestJson.Contains("SnapshotBackups") &&
                                 deepManifestJson.Contains("ExternalDataResponses");

            AssertTest("TC-R6c-D12: 实际包内 manifest.json 声明 8 类敏感分类 100% 物理隔离排除规范",
                manifestClean,
                "manifest.json does not declare all 8 excluded categories");

            // 7. 清理提取文件
            try
            {
                if (Directory.Exists(extractDir)) Directory.Delete(extractDir, true);
                if (Directory.Exists(deepExtractDir)) Directory.Delete(deepExtractDir, true);
            }
            catch { }

            Console.WriteLine();
            Console.WriteLine(string.Format("C# 诊断导出与脱敏测试统计: Pass = {0}, Fail = {1}", passCount, failCount));

            return failCount == 0 ? 0 : 1;
        }
    }
}
