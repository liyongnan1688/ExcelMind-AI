using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class MacroPackageEntry
    {
        public string macroId { get; set; }
        public string packageRelativePath { get; set; }
        public string displayName { get; set; }
        public string category { get; set; }
        public string description { get; set; }
        public string entryPoint { get; set; }
        public List<ScriptParameterDef> parameterDefs { get; set; }
        public long sourceByteLength { get; set; }
        public string sha256 { get; set; }

        public MacroPackageEntry()
        {
            macroId = "";
            packageRelativePath = "";
            displayName = "";
            category = "";
            description = "";
            entryPoint = "";
            parameterDefs = new List<ScriptParameterDef>();
            sourceByteLength = 0;
            sha256 = "";
        }
    }

    public class MacroPackageManifest
    {
        public string schemaVersion { get; set; }
        public string packageId { get; set; }
        public string name { get; set; }
        public string version { get; set; }
        public string description { get; set; }
        public string exportedAt { get; set; }
        public string exportedBy { get; set; }
        public List<MacroPackageEntry> entries { get; set; }

        public MacroPackageManifest()
        {
            schemaVersion = "1.0";
            packageId = Guid.NewGuid().ToString("N");
            name = "";
            version = "1.0.0";
            description = "";
            exportedAt = DateTime.UtcNow.ToString("o");
            exportedBy = "ExcelMind AI";
            entries = new List<MacroPackageEntry>();
        }
    }

    public class SensitiveScanWarning
    {
        public string macroDisplayName { get; set; }
        public string fieldName { get; set; }
        public string matchedPattern { get; set; }
        public string snippet { get; set; }
    }

    public class MacroPackageExportParams
    {
        public List<string> macroIds { get; set; }
        public string packageName { get; set; }
        public string packageVersion { get; set; }
        public string packageDescription { get; set; }
        public string targetFilePath { get; set; }
        public string outputPath { get { return targetFilePath; } set { targetFilePath = value; } }
        public bool ignoreSensitiveWarnings { get; set; }
        public bool ignoreWarnings { get { return ignoreSensitiveWarnings; } set { ignoreSensitiveWarnings = value; } }

        public MacroPackageExportParams()
        {
            macroIds = new List<string>();
            packageName = "";
            packageVersion = "1.0.0";
            packageDescription = "";
            targetFilePath = "";
            ignoreSensitiveWarnings = false;
        }
    }

    public class MacroPackageExportResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string packageFilePath { get; set; }
        public string outputPath { get { return packageFilePath; } set { packageFilePath = value; } }
        public int exportedCount { get; set; }
        public long totalBytes { get; set; }
        public long packageSize { get { return totalBytes; } set { totalBytes = value; } }
        public List<SensitiveScanWarning> sensitiveWarnings { get; set; }
        public List<SensitiveScanWarning> warnings { get { return sensitiveWarnings; } set { sensitiveWarnings = value; } }
        public bool hasSensitiveWarnings { get; set; }
        public MacroPackageManifest manifest { get; set; }

        public MacroPackageExportResult()
        {
            sensitiveWarnings = new List<SensitiveScanWarning>();
        }
    }

    public class MacroPackagePreviewResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public string packageFilePath { get; set; }
        public string packagePath { get { return packageFilePath; } set { packageFilePath = value; } }
        public MacroPackageManifest manifest { get; set; }
        public int macroCount { get; set; }
        public int entryCount { get { return macroCount; } set { macroCount = value; } }
        public long totalUncompressedBytes { get; set; }
        public long totalUncompressedSize { get { return totalUncompressedBytes; } set { totalUncompressedBytes = value; } }
        public List<SensitiveScanWarning> sensitiveWarnings { get; set; }
        public List<SensitiveScanWarning> warnings { get { return sensitiveWarnings; } set { sensitiveWarnings = value; } }
        public bool hasSensitiveWarnings { get; set; }
        public List<string> conflictingNames { get; set; }
        public List<string> nameConflicts { get { return conflictingNames; } set { conflictingNames = value; } }
        public bool isTamperedOrCorrupt { get; set; }

        public MacroPackagePreviewResult()
        {
            sensitiveWarnings = new List<SensitiveScanWarning>();
            conflictingNames = new List<string>();
        }
    }

    public class MacroPackageImportParams
    {
        public string packageFilePath { get; set; }
        public string packagePath { get { return packageFilePath; } set { packageFilePath = value; } }
        public List<string> selectedMacroIds { get; set; }
        public string conflictResolution { get; set; }
        public string nameConflictResolution { get { return conflictResolution; } set { conflictResolution = value; } }
        public bool ignoreWarnings { get; set; }

        public MacroPackageImportParams()
        {
            packageFilePath = "";
            selectedMacroIds = new List<string>();
            conflictResolution = "rename_both";
            ignoreWarnings = false;
        }
    }

    public class MacroPackageImportResult
    {
        public bool ok { get; set; }
        public string error { get; set; }
        public int importedCount { get; set; }
        public List<ScriptInfo> importedMacros { get; set; }
        public List<string> importedIds { get; set; }
        public List<string> renamedMacros { get; set; }
        public Dictionary<string, string> nameMapping { get; set; }
        public string summary { get; set; }
        public List<SensitiveScanWarning> warnings { get; set; }

        public MacroPackageImportResult()
        {
            importedMacros = new List<ScriptInfo>();
            importedIds = new List<string>();
            renamedMacros = new List<string>();
            nameMapping = new Dictionary<string, string>();
            warnings = new List<SensitiveScanWarning>();
            summary = "";
        }
    }

    public static class MacroPackageManager
    {
        public const int MaxEntries = 50;
        public const long MaxSingleFileBytes = 5 * 1024 * 1024;    // 5MB
        public const long MaxManifestBytes = 1 * 1024 * 1024;      // 1MB
        public const long MaxTotalUncompressedBytes = 20 * 1024 * 1024; // 20MB
        public const double MaxCompressionRatio = 20.0;            // 20倍压缩比上限

        private static string GetTempExtractionBaseDir()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string baseDir = Path.Combine(appData, "ExcelMindAI", "Temp", "MacroPackages");
            if (!Directory.Exists(baseDir))
            {
                Directory.CreateDirectory(baseDir);
            }
            return baseDir;
        }

        internal static string ComputeSha256(byte[] data)
        {
            if (data == null || data.Length == 0) return "";
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(data);
                var sb = new StringBuilder();
                for (int i = 0; i < hash.Length; i++)
                {
                    sb.Append(hash[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }

        public static List<SensitiveScanWarning> ScanSensitiveContent(string code, string desc, List<ScriptParameterDef> parameters, string macroName)
        {
            var warnings = new List<SensitiveScanWarning>();

            var secretPatterns = new Dictionary<string, string>
            {
                { "API_KEY_OR_TOKEN", @"(?i)(?:api[_-]?key|access[_-]?token|secret[_-]?key|bearer\s+[a-z0-9_\-\.]{12,})\s*[:=]\s*[""']?[a-z0-9_\-\.]{8,}[""']?" },
                { "PASSWORD_IN_CODE", @"(?i)(?:password|passwd|pwd)\s*[:=]\s*[""'][^""'\r\n]{4,}[""']" },
                { "PRIVATE_KEY_MARKER", @"(?i)-----BEGIN\s+[A-Z\s]+PRIVATE\s+KEY-----" },
                { "AUTH_QUERY_URL", @"(?i)https?://[^\s""'<>]+?(?:token|key|secret|password|auth)=[a-z0-9_\-\.]{6,}" }
            };

            Action<string, string> CheckText = delegate(string text, string field)
            {
                if (string.IsNullOrEmpty(text)) return;
                foreach (var kvp in secretPatterns)
                {
                    var m = Regex.Match(text, kvp.Value);
                    if (m.Success)
                    {
                        string snip = m.Value;
                        if (snip.Length > 60) snip = snip.Substring(0, 57) + "...";
                        warnings.Add(new SensitiveScanWarning
                        {
                            macroDisplayName = macroName ?? "",
                            fieldName = field,
                            matchedPattern = kvp.Key,
                            snippet = snip
                        });
                    }
                }
            };

            CheckText(code, "代码正文 (code)");
            CheckText(desc, "宏描述 (description)");
            if (parameters != null)
            {
                foreach (var p in parameters)
                {
                    if (p != null)
                    {
                        CheckText(p.defaultValue, string.Format("参数 [{0}] 默认值", p.name));
                        CheckText(p.description, string.Format("参数 [{0}] 描述", p.name));
                    }
                }
            }

            return warnings;
        }

        private static string SanitizePackageFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "macro";
            var invalids = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in name)
            {
                if (Array.IndexOf(invalids, c) >= 0 || c == ' ' || c == '\t')
                {
                    sb.Append('_');
                }
                else
                {
                    sb.Append(c);
                }
            }
            string res = sb.ToString().Trim('_');
            return string.IsNullOrEmpty(res) ? "macro" : res;
        }

        public static MacroPackageExportResult ExportPackage(MacroPackageExportParams p)
        {
            var res = new MacroPackageExportResult { ok = false };

            if (p == null || p.macroIds == null || p.macroIds.Count == 0)
            {
                res.error = "请至少选择一个待导出的宏。";
                return res;
            }

            if (p.macroIds.Count > MaxEntries)
            {
                res.error = string.Format("单次导出宏数量 ({0}) 超过上限 ({1})。", p.macroIds.Count, MaxEntries);
                return res;
            }

            if (string.IsNullOrEmpty(p.targetFilePath))
            {
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string safePkgName = string.IsNullOrEmpty(p.packageName) ? "macros" : SanitizePackageFileName(p.packageName);
                p.targetFilePath = Path.Combine(docs, string.Format("{0}_{1:yyyyMMdd_HHmmss}.exmpack", safePkgName, DateTime.Now));
            }

            try
            {
                string targetDir = Path.GetDirectoryName(p.targetFilePath);
                if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                var manifest = new MacroPackageManifest
                {
                    name = string.IsNullOrEmpty(p.packageName) ? "ExcelMind_MacroPackage" : p.packageName.Trim(),
                    version = string.IsNullOrEmpty(p.packageVersion) ? "1.0.0" : p.packageVersion.Trim(),
                    description = p.packageDescription ?? "",
                    exportedAt = DateTime.UtcNow.ToString("o")
                };

                var macroSourceBytes = new Dictionary<string, byte[]>();
                var allWarnings = new List<SensitiveScanWarning>();
                var seenRelPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var macroId in p.macroIds)
                {
                    var script = ScriptManager.GetScriptById(macroId);
                    if (script == null)
                    {
                        res.error = string.Format("未找到待导出的宏 (ID: {0})，导出已中止。", macroId);
                        return res;
                    }

                    if (string.IsNullOrEmpty(script.filePath) || !File.Exists(script.filePath))
                    {
                        res.error = string.Format("宏 [{0}] 的源码文件不存在: {1}", script.displayName, script.filePath);
                        return res;
                    }

                    // 1. 严格读取原始二进制字节，不改变编码、BOM、换行或字符
                    byte[] rawBytes = File.ReadAllBytes(script.filePath);
                    if (rawBytes.Length > MaxSingleFileBytes)
                    {
                        res.error = string.Format("宏 [{0}] 的源码大小 ({1:F1} MB) 超过单文件上限 ({2} MB)。", script.displayName, (double)rawBytes.Length / (1024 * 1024), MaxSingleFileBytes / (1024 * 1024));
                        return res;
                    }

                    string sha256 = ComputeSha256(rawBytes);

                    // 2. 敏感内容预检（检出提醒，绝不自动删改或打码源码）
                    string codeText = Encoding.UTF8.GetString(rawBytes);
                    var warnings = ScanSensitiveContent(codeText, script.description, script.parameters, script.displayName);
                    allWarnings.AddRange(warnings);

                    // 3. 构建包内相对路径
                    string safeBaseName = SanitizePackageFileName(script.displayName);
                    string relPath = "scripts/" + safeBaseName + ".bas";
                    int dupIdx = 2;
                    while (seenRelPaths.Contains(relPath))
                    {
                        relPath = string.Format("scripts/{0}_{1}.bas", safeBaseName, dupIdx);
                        dupIdx++;
                    }
                    seenRelPaths.Add(relPath);

                    // 4. 白名单元数据提取（严格不导出 runHistory, tags, 凭据, 工作簿路径, 快照）
                    var entry = new MacroPackageEntry
                    {
                        macroId = script.id,
                        packageRelativePath = relPath,
                        displayName = script.displayName,
                        category = string.IsNullOrEmpty(script.category) ? "通用" : script.category,
                        description = script.description ?? "",
                        entryPoint = script.entryPoint ?? "",
                        parameterDefs = script.parameters != null ? new List<ScriptParameterDef>(script.parameters) : new List<ScriptParameterDef>(),
                        sourceByteLength = rawBytes.Length,
                        sha256 = sha256
                    };

                    manifest.entries.Add(entry);
                    macroSourceBytes[relPath] = rawBytes;
                }

                res.sensitiveWarnings = allWarnings;
                res.hasSensitiveWarnings = allWarnings.Count > 0;

                // 若发现敏感内容且调用方未显式确认忽略，则阻断并返回告警
                if (res.hasSensitiveWarnings && !p.ignoreSensitiveWarnings)
                {
                    res.error = string.Format("在待导出的宏中检测到 {0} 项疑似敏感凭据/密钥信息。导出已暂停，请核对警告并确认后继续。", allWarnings.Count);
                    res.manifest = manifest;
                    return res;
                }

                // 5. 写入标准 ZIP 包 (.exmpack)
                if (File.Exists(p.targetFilePath))
                {
                    File.Delete(p.targetFilePath);
                }

                using (var zipStream = new FileStream(p.targetFilePath, FileMode.CreateNew, FileAccess.Write))
                using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
                {
                    // 写入 manifest.json
                    string manifestJson = SimpleJson.Serialize(manifest);
                    byte[] manifestBytes = Encoding.UTF8.GetBytes(manifestJson);
                    var mEntry = archive.CreateEntry("manifest.json", CompressionLevel.Optimal);
                    using (var entryStream = mEntry.Open())
                    {
                        entryStream.Write(manifestBytes, 0, manifestBytes.Length);
                    }

                    // 写入各宏源码原始字节
                    foreach (var kvp in macroSourceBytes)
                    {
                        var sEntry = archive.CreateEntry(kvp.Key, CompressionLevel.Optimal);
                        using (var entryStream = sEntry.Open())
                        {
                            entryStream.Write(kvp.Value, 0, kvp.Value.Length);
                        }
                    }
                }

                var outFi = new FileInfo(p.targetFilePath);
                res.ok = true;
                res.packageFilePath = p.targetFilePath;
                res.exportedCount = manifest.entries.Count;
                res.totalBytes = outFi.Length;
                res.manifest = manifest;
                return res;
            }
            catch (Exception ex)
            {
                res.error = "导出宏包失败: " + ex.Message;
                return res;
            }
        }

        private static void SafeDeleteDirectory(string dir)
        {
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return;
            try
            {
                Directory.Delete(dir, true);
            }
            catch { }
        }

        public static MacroPackagePreviewResult PreviewPackage(string packageFilePath)
        {
            var res = new MacroPackagePreviewResult
            {
                ok = false,
                packageFilePath = packageFilePath
            };

            if (string.IsNullOrEmpty(packageFilePath) || !File.Exists(packageFilePath))
            {
                res.error = "指定的宏包文件不存在: " + packageFilePath;
                return res;
            }

            string extractDir = Path.Combine(GetTempExtractionBaseDir(), "_pkg_temp_" + Guid.NewGuid().ToString("N"));

            try
            {
                // 1. 结构与容量预检解包（解压至隔离临时解包目录）
                MacroPackageManifest manifest;
                long totalBytes = ValidateAndExtractZip(packageFilePath, extractDir, out manifest);

                // 2. 检查现有宏库同名冲突
                var existingScripts = ScriptManager.ListScripts();
                var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var s in existingScripts)
                {
                    existingNames.Add(s.displayName);
                }

                var conflictList = new List<string>();
                foreach (var e in manifest.entries)
                {
                    if (existingNames.Contains(e.displayName))
                    {
                        conflictList.Add(e.displayName);
                    }
                }

                // 3. 敏感信息探测
                var allWarnings = new List<SensitiveScanWarning>();
                foreach (var e in manifest.entries)
                {
                    string filePath = Path.Combine(extractDir, e.packageRelativePath);
                    if (File.Exists(filePath))
                    {
                        string code = File.ReadAllText(filePath, Encoding.UTF8);
                        var warnings = ScanSensitiveContent(code, e.description, e.parameterDefs, e.displayName);
                        allWarnings.AddRange(warnings);
                    }
                }

                res.ok = true;
                res.manifest = manifest;
                res.macroCount = manifest.entries.Count;
                res.totalUncompressedBytes = totalBytes;
                res.conflictingNames = conflictList;
                res.sensitiveWarnings = allWarnings;
                res.hasSensitiveWarnings = allWarnings.Count > 0;
                res.isTamperedOrCorrupt = false;
                return res;
            }
            catch (Exception ex)
            {
                res.error = "宏包解析失败: " + ex.Message;
                res.isTamperedOrCorrupt = true;
                return res;
            }
            finally
            {
                SafeDeleteDirectory(extractDir);
            }
        }

        public static MacroPackageImportResult ImportPackage(MacroPackageImportParams p)
        {
            var res = new MacroPackageImportResult { ok = false };

            if (p == null || string.IsNullOrEmpty(p.packageFilePath) || !File.Exists(p.packageFilePath))
            {
                res.error = "指定的宏包文件不存在或参数为空。";
                return res;
            }

            string extractDir = Path.Combine(GetTempExtractionBaseDir(), "_pkg_temp_" + Guid.NewGuid().ToString("N"));
            var createdFiles = new List<string>();

            try
            {
                // 1. 隔离解包与全量哈希完整性校验
                MacroPackageManifest manifest;
                ValidateAndExtractZip(p.packageFilePath, extractDir, out manifest);

                // 2. 筛选待导入宏清单
                var entriesToImport = new List<MacroPackageEntry>();
                if (p.selectedMacroIds != null && p.selectedMacroIds.Count > 0)
                {
                    var idSet = new HashSet<string>(p.selectedMacroIds, StringComparer.OrdinalIgnoreCase);
                    foreach (var e in manifest.entries)
                    {
                        if (idSet.Contains(e.macroId) || idSet.Contains(e.displayName))
                        {
                            entriesToImport.Add(e);
                        }
                    }
                }
                else
                {
                    entriesToImport.AddRange(manifest.entries);
                }

                if (entriesToImport.Count == 0)
                {
                    res.error = "未选中任何待导入的宏。";
                    return res;
                }

                // 3. 准备内存中的待写入项（同名冲突保留双方、新本地 ID、R2c 参数校验）
                var existingScripts = ScriptManager.ListScripts();
                var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var existingFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var s in existingScripts)
                {
                    existingNames.Add(s.displayName);
                    existingFileNames.Add(s.fileName);
                }

                var importPlans = new List<Tuple<string, byte[], Dictionary<string, object>, string>>(); // destBas, bytes, meta, displayName

                foreach (var entry in entriesToImport)
                {
                    string srcFile = Path.Combine(extractDir, entry.packageRelativePath);
                    if (!File.Exists(srcFile))
                    {
                        throw new InvalidOperationException("宏包内缺失源码文件: " + entry.packageRelativePath);
                    }

                    byte[] codeBytes = File.ReadAllBytes(srcFile);

                    // 参数契约校验 (R2c 严格类型与签名核对)
                    if (entry.parameterDefs != null && entry.parameterDefs.Count > 0)
                    {
                        string codeText = Encoding.UTF8.GetString(codeBytes);
                        var sigs = VbaSignatureParser.ParseSignatures(codeText);
                        var proc = sigs.Find(s => string.Equals(s.name, entry.entryPoint, StringComparison.OrdinalIgnoreCase));
                        if (proc != null)
                        {
                            var cmp = VbaSignatureParser.CompareWithMetadata(proc, entry.parameterDefs);
                            if (!cmp.isMatch)
                            {
                                throw new InvalidOperationException(string.Format("宏 [{0}] 参数契约与源码签名不一致: {1}", entry.displayName, cmp.error));
                            }
                        }
                    }

                    // 同名冲突处理：第一切片默认保留双方，显式重命名导入宏，绝不静默覆盖已有宏
                    string targetDispName = entry.displayName;
                    string safeBaseName = SanitizePackageFileName(entry.displayName);
                    string targetFileName = safeBaseName + ".bas";

                    int dupIdx = 2;
                    while (existingNames.Contains(targetDispName) || existingFileNames.Contains(targetFileName) || File.Exists(Path.Combine(ScriptManager.ScriptsDir, targetFileName)))
                    {
                        targetDispName = string.Format("{0} (导入{1})", entry.displayName, dupIdx == 2 ? "" : " " + dupIdx);
                        targetFileName = string.Format("{0}_imported{1}.bas", safeBaseName, dupIdx == 2 ? "" : "_" + dupIdx);
                        dupIdx++;
                    }

                    existingNames.Add(targetDispName);
                    existingFileNames.Add(targetFileName);

                    string destBasPath = Path.Combine(ScriptManager.ScriptsDir, targetFileName);
                    string destMetaPath = Path.Combine(ScriptManager.ScriptsDir, Path.GetFileNameWithoutExtension(targetFileName) + ".meta.json");

                    // 生成全新的本地稳定 ID（包内 ID 绝不覆盖本地已有宏）
                    string newId = Guid.NewGuid().ToString("N");

                    var meta = new Dictionary<string, object>
                    {
                        { "id", newId },
                        { "name", targetDispName },
                        { "displayName", targetDispName },
                        { "description", entry.description ?? "" },
                        { "category", string.IsNullOrEmpty(entry.category) ? "已导入" : entry.category },
                        { "sourceType", "package_import" },
                        { "originalFileName", Path.GetFileName(entry.packageRelativePath) },
                        { "encoding", "UTF-8" },
                        { "originalCodeHash", entry.sha256 },
                        { "entryPoint", entry.entryPoint ?? "" },
                        { "createdAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
                        { "updatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
                        { "lastExecutionResult", "未运行" },
                        { "isFavorite", false },
                        { "tags", new List<string> { "已导入" } },
                        { "runHistory", new List<RunRecordDto>() } // 强行清空，绝不夹带历史
                    };

                    if (entry.parameterDefs != null && entry.parameterDefs.Count > 0)
                    {
                        meta["parameters"] = entry.parameterDefs;
                    }

                    importPlans.Add(new Tuple<string, byte[], Dictionary<string, object>, string>(
                        destBasPath, codeBytes, meta, targetDispName));
                }

                // 4. 原子化写入宏库（失败时自动回滚清理本次新增文件，绝不污染存量宏）
                foreach (var plan in importPlans)
                {
                    string basPath = plan.Item1;
                    byte[] bytes = plan.Item2;
                    var metaDict = plan.Item3;
                    string metaPath = Path.Combine(Path.GetDirectoryName(basPath), Path.GetFileNameWithoutExtension(basPath) + ".meta.json");

                    File.WriteAllBytes(basPath, bytes);
                    createdFiles.Add(basPath);

                    string metaJson = SimpleJson.Serialize(metaDict);
                    File.WriteAllText(metaPath, metaJson, Encoding.UTF8);
                    createdFiles.Add(metaPath);
                }

                // 5. 组装返回结果（纯展示结果，零宏执行，不自动加入工作流或收藏）
                foreach (var plan in importPlans)
                {
                    var metaDict = plan.Item3;
                    res.importedMacros.Add(new ScriptInfo
                    {
                        id = (string)metaDict["id"],
                        name = (string)metaDict["name"],
                        displayName = (string)metaDict["displayName"],
                        fileName = Path.GetFileName(plan.Item1),
                        filePath = plan.Item1,
                        category = (string)metaDict["category"],
                        description = (string)metaDict["description"],
                        entryPoint = (string)metaDict["entryPoint"],
                        originalCodeHash = (string)metaDict["originalCodeHash"]
                    });

                    res.importedIds.Add((string)metaDict["id"]);
                    res.nameMapping[plan.Item4] = (string)metaDict["displayName"];

                    if (plan.Item4 != plan.Item3["name"].ToString() || plan.Item4.Contains("(导入"))
                    {
                        res.renamedMacros.Add(string.Format("{0} -> {1}", plan.Item4, plan.Item3["displayName"]));
                    }
                }

                res.ok = true;
                res.importedCount = importPlans.Count;
                res.summary = string.Format("已安全导入 {0} 个宏到宏库（零自动执行，元数据白名单生效）。", importPlans.Count);
                return res;
            }
            catch (Exception ex)
            {
                // 原子回滚清理本次新增的所有文件
                foreach (var file in createdFiles)
                {
                    try
                    {
                        if (File.Exists(file)) File.Delete(file);
                    }
                    catch { }
                }

                res.error = "宏包导入失败已回滚: " + ex.Message;
                return res;
            }
            finally
            {
                SafeDeleteDirectory(extractDir);
            }
        }

        private static long ValidateAndExtractZip(string zipPath, string targetExtractDir, out MacroPackageManifest manifest)
        {
            manifest = null;
            if (!Directory.Exists(targetExtractDir))
            {
                Directory.CreateDirectory(targetExtractDir);
            }

            string targetDirFullPath = Path.GetFullPath(targetExtractDir);
            if (!targetDirFullPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                targetDirFullPath += Path.DirectorySeparatorChar;
            }

            long totalUncompressedBytes = 0;
            var seenEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var zipStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Read))
            {
                if (archive.Entries.Count == 0)
                {
                    throw new InvalidOperationException("宏包为空 (条目数为 0)。");
                }

                if (archive.Entries.Count > MaxEntries)
                {
                    throw new InvalidOperationException(string.Format("宏包条目数 ({0}) 超过系统上限 ({1})。", archive.Entries.Count, MaxEntries));
                }

                // 第一阶段：包结构与安全性完整预检
                bool hasManifest = false;
                foreach (var entry in archive.Entries)
                {
                    string fullName = entry.FullName.Replace('\\', '/');

                    if (string.IsNullOrEmpty(fullName))
                    {
                        throw new InvalidOperationException("宏包内包含空名称条目。");
                    }

                    // 阻断绝对路径、盘符与 UNC 路径
                    if (fullName.StartsWith("/") || fullName.Contains(":") || fullName.StartsWith("//") || fullName.StartsWith("\\\\"))
                    {
                        throw new InvalidOperationException("检测到非法绝对路径或盘符条目，已安全阻断: " + fullName);
                    }

                    // 阻断路径穿越 (规范化路径必须严格落在目标临时目录内)
                    string destPath = Path.GetFullPath(Path.Combine(targetExtractDir, entry.FullName));
                    if (!destPath.StartsWith(targetDirFullPath, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException("检测到非法路径穿越条目，已安全阻断: " + fullName);
                    }

                    // 重复条目与大小写碰撞阻断
                    if (!seenEntries.Add(fullName.ToLowerInvariant()))
                    {
                        throw new InvalidOperationException("宏包内存在重复条目或名称碰撞: " + fullName);
                    }

                    // 严格限制合法结构：根目录 manifest.json 与 scripts/*.bas 源码文件
                    if (string.Equals(fullName, "manifest.json", StringComparison.OrdinalIgnoreCase))
                    {
                        hasManifest = true;
                        if (entry.Length > MaxManifestBytes)
                        {
                            throw new InvalidOperationException(string.Format("manifest.json 大小 ({0:F1} KB) 超过上限 ({1} KB)。", (double)entry.Length / 1024, MaxManifestBytes / 1024));
                        }
                    }
                    else if (fullName.StartsWith("scripts/", StringComparison.OrdinalIgnoreCase) && fullName.EndsWith(".bas", StringComparison.OrdinalIgnoreCase))
                    {
                        if (entry.Length > MaxSingleFileBytes)
                        {
                            throw new InvalidOperationException(string.Format("宏源码文件 [{0}] 解压大小 ({1:F1} MB) 超过单文件上限 ({2} MB)。", fullName, (double)entry.Length / (1024 * 1024), MaxSingleFileBytes / (1024 * 1024)));
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException("宏包包含不支持的目录结构或非法文件类型 (仅允许 manifest.json 与 scripts/*.bas): " + fullName);
                    }

                    // 压缩炸弹比例防御 (若未压缩大小 > 1KB 且解压比 > 20)
                    if (entry.CompressedLength > 0 && entry.Length > 1024)
                    {
                        double ratio = (double)entry.Length / entry.CompressedLength;
                        if (ratio > MaxCompressionRatio)
                        {
                            throw new InvalidOperationException(string.Format("条目 [{0}] 压缩比异常过高 ({1:F1} 倍)，疑似压缩炸弹，已安全阻断。", fullName, ratio));
                        }
                    }

                    totalUncompressedBytes += entry.Length;
                    if (totalUncompressedBytes > MaxTotalUncompressedBytes)
                    {
                        throw new InvalidOperationException(string.Format("宏包总解压大小 ({0:F1} MB) 超过系统上限 ({1} MB)。", (double)totalUncompressedBytes / (1024 * 1024), MaxTotalUncompressedBytes / (1024 * 1024)));
                    }
                }

                if (!hasManifest)
                {
                    throw new InvalidOperationException("宏包缺少核心描述文件 manifest.json，无法识别。");
                }

                // 第二阶段：安全解压至隔离临时目录
                foreach (var entry in archive.Entries)
                {
                    string destPath = Path.Combine(targetExtractDir, entry.FullName);
                    string destDir = Path.GetDirectoryName(destPath);
                    if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                    {
                        Directory.CreateDirectory(destDir);
                    }

                    using (var eStream = entry.Open())
                    using (var fStream = new FileStream(destPath, FileMode.CreateNew, FileAccess.Write))
                    {
                        eStream.CopyTo(fStream);
                    }
                }
            }

            // 第三阶段：校验 manifest.json 与包内文件哈希
            string manifestFile = Path.Combine(targetExtractDir, "manifest.json");
            string manifestContent = File.ReadAllText(manifestFile, Encoding.UTF8);

            try
            {
                manifest = SimpleJson.Deserialize<MacroPackageManifest>(manifestContent);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("解析 manifest.json 失败: " + ex.Message);
            }

            if (manifest == null)
            {
                throw new InvalidOperationException("manifest.json 内容为空或格式不合法。");
            }

            if (manifest.schemaVersion != "1.0")
            {
                throw new InvalidOperationException("不支持的宏包规范版本 (schemaVersion): " + manifest.schemaVersion + " (当前仅支持 1.0)");
            }

            if (manifest.entries == null || manifest.entries.Count == 0)
            {
                throw new InvalidOperationException("manifest.json 中未声明任何有效宏条目。");
            }

            // 逐个校验包内实际文件存在性与 SHA-256 哈希
            foreach (var entry in manifest.entries)
            {
                if (string.IsNullOrEmpty(entry.packageRelativePath))
                {
                    throw new InvalidOperationException("宏条目 packageRelativePath 不能为空。");
                }

                string entryFile = Path.Combine(targetExtractDir, entry.packageRelativePath);
                if (!File.Exists(entryFile))
                {
                    throw new InvalidOperationException("manifest 声明的文件在包内缺失: " + entry.packageRelativePath);
                }

                byte[] actualBytes = File.ReadAllBytes(entryFile);
                if (actualBytes.Length != entry.sourceByteLength)
                {
                    throw new InvalidOperationException(string.Format("文件 [{0}] 实际大小 ({1}) 与清单声明 ({2}) 不符，可能已被篡改。", entry.packageRelativePath, actualBytes.Length, entry.sourceByteLength));
                }

                string actualHash = ComputeSha256(actualBytes);
                if (!string.Equals(actualHash, entry.sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(string.Format("文件 [{0}] SHA-256 完整性校验失败 (实际: {1}, 清单: {2})，可能已被篡改或损坏。", entry.packageRelativePath, actualHash, entry.sha256));
                }
            }

            return totalUncompressedBytes;
        }
    }
}
