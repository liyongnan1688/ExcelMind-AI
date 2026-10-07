using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class RunRecordDto
    {
        public string id { get; set; }
        public string executedAt { get; set; }
        public string status { get; set; }          // "success" | "failed" | "blocked"
        public string phase { get; set; }           // "execution" | "precheck" | "compile" | "runtime" | "before_run"
        public string targetWorkbookName { get; set; }
        public string codeHash { get; set; }
        public string snapshotId { get; set; }
        public bool snapshotExists { get; set; }
        public string snapshotReason { get; set; }
        public int elapsedMs { get; set; }
        public string summary { get; set; }
        public string entryPoint { get; set; }
        public List<string> parameterTypes { get; set; }
        public Dictionary<string, string> parameterSummary { get; set; }
    }

    public class ScriptInfo
    {
        public string id { get; set; }
        public string name { get; set; }
        public string displayName { get; set; }
        public string fileName { get; set; }
        public string filePath { get; set; }
        public string createdAt { get; set; }
        public string updatedAt { get; set; }
        public string description { get; set; }
        public string category { get; set; }
        public string sourceType { get; set; }
        public string originalFileName { get; set; }
        public string encoding { get; set; }
        public string code { get; set; }
        public string originalCodeHash { get; set; }
        public string entryPoint { get; set; }
        public string rawBytesBase64 { get; set; }
        public string lastExecutionResult { get; set; }
        public string lastExecutedAt { get; set; }
        public bool isVerified { get; set; }
        public bool isFavorite { get; set; }
        public List<string> tags { get; set; }
        public List<RunRecordDto> runHistory { get; set; }
        public List<ScriptParameterDef> parameters { get; set; }
        public List<VbaEntryPointInfo> candidateEntryPoints { get; set; }
    }

    public class ScriptManager
    {
        private static string _customScriptsDir = null;
        private static readonly string _defaultScriptsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelMindAI",
            "Scripts"
        );

        public static string ScriptsDir
        {
            get { return _customScriptsDir ?? _defaultScriptsDir; }
        }

        public static void SetCustomScriptsDirForTesting(string customDir)
        {
            _customScriptsDir = customDir;
            if (!string.IsNullOrEmpty(_customScriptsDir) && !Directory.Exists(_customScriptsDir))
            {
                Directory.CreateDirectory(_customScriptsDir);
            }
        }

        public static void ResetCustomScriptsDirForTesting()
        {
            _customScriptsDir = null;
        }

        static ScriptManager()
        {
            if (!Directory.Exists(_defaultScriptsDir))
            {
                Directory.CreateDirectory(_defaultScriptsDir);
                try
                {
                    string oldDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "LeeExcel",
                        "Scripts"
                    );
                    if (Directory.Exists(oldDir))
                    {
                        foreach (var file in Directory.GetFiles(oldDir, "*.bas"))
                        {
                            string dest = Path.Combine(_defaultScriptsDir, Path.GetFileName(file));
                            if (!File.Exists(dest)) File.Copy(file, dest);
                        }
                    }
                }
                catch { }
            }
        }

        // 内部哈希计算辅助方法，收敛可见性为 internal，委托统一的 VbaRunner.ComputeSha256，避免被 Excel-DNA 自动导出为 Excel 工作表函数产生重复注册冲突
        internal static string ComputeSha256(string text)
        {
            return VbaRunner.ComputeSha256(text);
        }

        public static bool ResolveScriptFiles(string idOrFileName, out string basPath, out string metaPath, out string scriptId, out string error)
        {
            basPath = null;
            metaPath = null;
            scriptId = null;
            error = null;

            if (string.IsNullOrEmpty(idOrFileName))
            {
                error = "宏标识不能为空";
                return false;
            }

            // 1. 优先按元数据 "id" 字段精准匹配前端稳定 ID（避免因“文件名优先”而误操作文件名恰巧等于该 ID 的另一条宏）
            List<string> matchedMetaFiles = new List<string>();
            List<string> matchedIds = new List<string>();

            if (Directory.Exists(ScriptsDir))
            {
                foreach (var metaFile in Directory.GetFiles(ScriptsDir, "*.meta.json"))
                {
                    try
                    {
                        string metaContent = File.ReadAllText(metaFile, Encoding.UTF8);
                        var m = SimpleJson.ParseFlatObject(metaContent);
                        if (m != null && m.ContainsKey("id") && string.Equals(m["id"], idOrFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            matchedMetaFiles.Add(metaFile);
                            matchedIds.Add(m["id"]);
                        }
                    }
                    catch { }
                }
            }

            // 歧义核验 1：若存在多个元数据文件具有相同 ID，明确阻断歧义，严禁扫描后静默取第一个
            if (matchedMetaFiles.Count > 1)
            {
                error = "宏标识存在重复冲突: 发现多个元数据包含相同 ID【" + idOrFileName + "】";
                return false;
            }

            // 若精准命中唯一元数据 ID，定位成功
            if (matchedMetaFiles.Count == 1)
            {
                metaPath = matchedMetaFiles[0];
                string fileName = Path.GetFileName(metaPath);
                string bName = fileName.EndsWith(".meta.json", StringComparison.OrdinalIgnoreCase)
                    ? fileName.Substring(0, fileName.Length - ".meta.json".Length)
                    : (fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase)
                        ? Path.GetFileNameWithoutExtension(fileName)
                        : fileName);
                basPath = Path.Combine(ScriptsDir, bName + ".bas");
                scriptId = matchedIds[0];
                return true;
            }

            // 2. 若元数据 ID 未命中，再按磁盘物理文件名精准查找（兼容通过直接文件名或无元数据旧文件的操作）
            string baseName = idOrFileName;
            if (baseName.EndsWith(".meta.json", StringComparison.OrdinalIgnoreCase))
            {
                baseName = baseName.Substring(0, baseName.Length - ".meta.json".Length);
            }
            else
            {
                baseName = Path.GetFileNameWithoutExtension(baseName);
            }
            string directBas = Path.Combine(ScriptsDir, baseName + ".bas");
            string directMeta = Path.Combine(ScriptsDir, baseName + ".meta.json");

            if (File.Exists(directBas) || File.Exists(directMeta))
            {
                basPath = directBas;
                metaPath = directMeta;
                scriptId = baseName;
                if (File.Exists(directMeta))
                {
                    try
                    {
                        var m = SimpleJson.ParseFlatObject(File.ReadAllText(directMeta, Encoding.UTF8));
                        if (m != null && m.ContainsKey("id") && !string.IsNullOrEmpty(m["id"]))
                        {
                            scriptId = m["id"];
                        }
                    }
                    catch { }
                }
                return true;
            }

            // 3. 严格不按显示名称 (displayName) 模糊兜底，未找到即返回明确错误
            error = "未找到目标宏: " + idOrFileName;
            return false;
        }

        public static bool ResolveScriptFiles(string idOrFileName, out string basPath, out string metaPath, out string scriptId)
        {
            string dummyErr;
            return ResolveScriptFiles(idOrFileName, out basPath, out metaPath, out scriptId, out dummyErr);
        }

        public static ScriptInfo GetScriptById(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            string baseName = Path.GetFileNameWithoutExtension(id);
            return ListScripts().Find(s => 
                string.Equals(s.id, id, StringComparison.OrdinalIgnoreCase) || 
                string.Equals(s.fileName, id, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Path.GetFileNameWithoutExtension(s.fileName), baseName, StringComparison.OrdinalIgnoreCase));
        }

        public static List<ScriptInfo> ListScripts()
        {
            var list = new List<ScriptInfo>();
            if (!Directory.Exists(ScriptsDir)) return list;

            var files = Directory.GetFiles(ScriptsDir, "*.bas");
            Array.Sort(files, (a, b) => File.GetLastWriteTime(b).CompareTo(File.GetLastWriteTime(a)));

            foreach (var f in files)
            {
                try
                {
                    string baseName = Path.GetFileNameWithoutExtension(f);
                    string metaFile = Path.Combine(ScriptsDir, baseName + ".meta.json");
                    string content = File.ReadAllText(f, Encoding.UTF8);

                    if (File.Exists(metaFile))
                    {
                        try
                        {
                            string metaJson = File.ReadAllText(metaFile, Encoding.UTF8);
                            var meta = SimpleJson.ParseFlatObject(metaJson);

                            string id = meta.ContainsKey("id") ? meta["id"] : baseName;
                            string dispName = meta.ContainsKey("displayName") ? meta["displayName"] :
                                              (meta.ContainsKey("name") ? meta["name"] : baseName);
                            string desc = meta.ContainsKey("description") ? meta["description"] : "";
                            string cat = meta.ContainsKey("category") ? meta["category"] : "";
                            string srcType = meta.ContainsKey("sourceType") ? meta["sourceType"] : "paste";
                            string origName = meta.ContainsKey("originalFileName") ? meta["originalFileName"] : "";
                            string enc = meta.ContainsKey("encoding") ? meta["encoding"] : "UTF-8";
                            string hash = meta.ContainsKey("originalCodeHash") ? meta["originalCodeHash"] : ComputeSha256(content);
                            string entry = meta.ContainsKey("entryPoint") ? meta["entryPoint"] : "";
                            string rawBytes = meta.ContainsKey("rawBytesBase64") ? meta["rawBytesBase64"] : "";
                            string created = meta.ContainsKey("createdAt") ? meta["createdAt"] : File.GetCreationTime(f).ToString("yyyy-MM-dd HH:mm");
                            string updated = meta.ContainsKey("updatedAt") ? meta["updatedAt"] : File.GetLastWriteTime(f).ToString("yyyy-MM-dd HH:mm");
                            string lastRes = meta.ContainsKey("lastExecutionResult") ? meta["lastExecutionResult"] : "未运行";
                            string lastTime = meta.ContainsKey("lastExecutedAt") ? meta["lastExecutedAt"] : "";

                            // 增量解析 tags，缺失时赋空列表默认值，不重写旧文件
                            var tags = new List<string>();
                            if (meta.ContainsKey("tags") && !string.IsNullOrEmpty(meta["tags"]))
                            {
                                try
                                {
                                    tags = SimpleJson.ParseStringList(meta["tags"]);
                                }
                                catch { }
                            }

                            // 增量解析 runHistory，缺失时赋空列表默认值，不重写旧文件
                            var runHistory = new List<RunRecordDto>();
                            if (meta.ContainsKey("runHistory") && !string.IsNullOrEmpty(meta["runHistory"]))
                            {
                                try
                                {
                                    runHistory = SimpleJson.DeserializeList<RunRecordDto>(meta["runHistory"]) ?? new List<RunRecordDto>();
                                    // 动态核验各条记录快照物理文件是否依然可用
                                    foreach (var rec in runHistory)
                                    {
                                        if (!string.IsNullOrEmpty(rec.snapshotId))
                                        {
                                            rec.snapshotExists = SnapshotManager.SnapshotExists("", rec.snapshotId);
                                        }
                                        else
                                        {
                                            rec.snapshotExists = false;
                                        }
                                    }
                                }
                                catch { }
                            }

                            // 增量解析 isFavorite，缺失时默认为 false，不重写旧文件
                            bool isFavorite = meta.ContainsKey("isFavorite") && (meta["isFavorite"] == "true" || meta["isFavorite"] == "1" || meta["isFavorite"] == "True");

                            // 增量解析 parameters，缺失时为 null，不重写旧文件
                            var paramDefs = new List<ScriptParameterDef>();
                            if (meta.ContainsKey("parameters") && !string.IsNullOrEmpty(meta["parameters"]))
                            {
                                try
                                {
                                    paramDefs = SimpleJson.DeserializeList<ScriptParameterDef>(meta["parameters"]) ?? new List<ScriptParameterDef>();
                                }
                                catch { }
                            }

                            // 动态分析源码中的候选入口过程
                            var candidateEntryPoints = VbaSignatureParser.ParseSignatures(content);

                            list.Add(new ScriptInfo
                            {
                                id = id,
                                name = dispName,
                                displayName = dispName,
                                fileName = Path.GetFileName(f),
                                filePath = f,
                                createdAt = created,
                                updatedAt = updated,
                                description = desc,
                                category = cat,
                                sourceType = srcType,
                                originalFileName = origName,
                                encoding = enc,
                                code = content,
                                originalCodeHash = hash,
                                entryPoint = entry,
                                rawBytesBase64 = rawBytes,
                                lastExecutionResult = lastRes,
                                lastExecutedAt = lastTime,
                                isVerified = lastRes.Contains("成功"),
                                isFavorite = isFavorite,
                                tags = tags,
                                runHistory = runHistory,
                                parameters = paramDefs.Count > 0 ? paramDefs : null,
                                candidateEntryPoints = candidateEntryPoints
                            });
                        }
                        catch (Exception exMeta)
                        {
                            // 单项损坏容灾：不让整库加载失败，不自动覆盖原损坏文件，降级保护性呈现
                            list.Add(new ScriptInfo
                            {
                                id = baseName,
                                name = baseName,
                                displayName = baseName,
                                fileName = Path.GetFileName(f),
                                filePath = f,
                                createdAt = File.GetCreationTime(f).ToString("yyyy-MM-dd HH:mm"),
                                updatedAt = File.GetLastWriteTime(f).ToString("yyyy-MM-dd HH:mm"),
                                description = "⚠️ [元数据文件损坏，已保护性加载源码: " + exMeta.Message + "]",
                                category = "未分类",
                                sourceType = "corrupted_meta",
                                originalFileName = Path.GetFileName(f),
                                encoding = "UTF-8",
                                code = content,
                                originalCodeHash = ComputeSha256(content),
                                entryPoint = "",
                                rawBytesBase64 = "",
                                lastExecutionResult = "元数据损坏",
                                lastExecutedAt = "",
                                isVerified = false,
                                isFavorite = false,
                                tags = new List<string>(),
                                runHistory = new List<RunRecordDto>(),
                                parameters = null,
                                candidateEntryPoints = VbaSignatureParser.ParseSignatures(content)
                            });
                        }
                    }
                    else
                    {
                        // 兼容旧版纯 .bas 存储格式，正常读库绝不批量重写旧文件
                        string desc = ExtractHeaderField(content, "Description") ?? "自定义保存脚本";
                        string created = ExtractHeaderField(content, "Created At") ?? File.GetCreationTime(f).ToString("yyyy-MM-dd HH:mm");
                        string updated = File.GetLastWriteTime(f).ToString("yyyy-MM-dd HH:mm");
                        string name = Path.GetFileNameWithoutExtension(f);

                        list.Add(new ScriptInfo
                        {
                            id = name,
                            name = name,
                            displayName = name,
                            fileName = Path.GetFileName(f),
                            filePath = f,
                            createdAt = created,
                            updatedAt = updated,
                            description = desc,
                            category = "默认分类",
                            sourceType = "legacy",
                            originalFileName = Path.GetFileName(f),
                            encoding = "UTF-8",
                            code = content,
                            originalCodeHash = ComputeSha256(content),
                            entryPoint = "",
                            rawBytesBase64 = "",
                            lastExecutionResult = "未运行",
                            lastExecutedAt = "",
                            isVerified = false,
                            isFavorite = false,
                            tags = new List<string>(),
                            runHistory = new List<RunRecordDto>()
                        });
                    }
                }
                catch { }
            }
            return list;
        }

        public static bool SaveScript(
            string id,
            string displayName,
            string code,
            string description,
            string category,
            string sourceType,
            string originalFileName,
            string encoding,
            string entryPoint,
            string rawBytesBase64,
            bool overwrite,
            out string savedId,
            out string error)
        {
            savedId = null;
            error = null;

            try
            {
                if (string.IsNullOrEmpty(displayName))
                {
                    error = "宏显示名称不能为空";
                    return false;
                }

                displayName = displayName.Trim();

                if (code == null)
                {
                    code = "";
                }

                // 查重：检查是否已有相同显示名称的其他条目
                var existingList = ListScripts();
                ScriptInfo sameNameItem = null;
                foreach (var item in existingList)
                {
                    if (string.Equals(item.displayName, displayName, StringComparison.OrdinalIgnoreCase))
                    {
                        sameNameItem = item;
                        break;
                    }
                }

                if (sameNameItem != null)
                {
                    // 若存在同名条目且既非当前编辑ID亦未确认覆盖，安全拦截返回警告
                    if (!string.Equals(sameNameItem.id, id, StringComparison.OrdinalIgnoreCase) && !overwrite)
                    {
                        error = "DUPLICATE_NAME:已存在同名的宏【" + displayName + "】。若需覆盖请明确确认。";
                        return false;
                    }
                    if (overwrite && string.IsNullOrEmpty(id))
                    {
                        id = sameNameItem.id;
                    }
                }

                if (string.IsNullOrEmpty(id))
                {
                    // 生成安全唯一的 ID
                    string safePrefix = "macro_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    id = safePrefix;
                }
                else
                {
                    foreach (char c in Path.GetInvalidFileNameChars())
                    {
                        id = id.Replace(c, '_');
                    }
                }

                savedId = id;
                string basFileName = id.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) ? id : id + ".bas";
                string baseName = Path.GetFileNameWithoutExtension(basFileName);
                string basFullPath = Path.Combine(ScriptsDir, basFileName);
                string metaFullPath = Path.Combine(ScriptsDir, baseName + ".meta.json");

                // 1. 原文保真保存：写入纯净的原始源码（严禁强行添加注释头或修改正文）
                File.WriteAllText(basFullPath, code, Encoding.UTF8);

                // 2. 独立元数据持久化：保存显示名称、编码、哈希、入口等结构化信息
                string hash = ComputeSha256(code);
                string nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                string origCreated = nowStr;
                string prevLastRes = "未运行";
                string prevLastTime = "";
                if (sameNameItem != null)
                {
                    origCreated = sameNameItem.createdAt ?? nowStr;
                    prevLastRes = sameNameItem.lastExecutionResult ?? "未运行";
                    prevLastTime = sameNameItem.lastExecutedAt ?? "";
                }

                var meta = new Dictionary<string, string>
                {
                    { "id", id },
                    { "name", displayName },
                    { "displayName", displayName },
                    { "description", description ?? "" },
                    { "category", category ?? "" },
                    { "sourceType", sourceType ?? "paste" },
                    { "originalFileName", originalFileName ?? "" },
                    { "encoding", encoding ?? "UTF-8" },
                    { "originalCodeHash", hash },
                    { "entryPoint", entryPoint ?? "" },
                    { "rawBytesBase64", rawBytesBase64 ?? "" },
                    { "createdAt", origCreated },
                    { "updatedAt", nowStr },
                    { "lastExecutionResult", prevLastRes },
                    { "lastExecutedAt", prevLastTime }
                };

                string metaJson = SimpleJson.Serialize(meta);
                File.WriteAllText(metaFullPath, metaJson, Encoding.UTF8);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // 兼容旧接口重载
        public static bool SaveScript(string name, string code, string description, out string error)
        {
            string savedId;
            return SaveScript(
                id: null,
                displayName: name,
                code: code,
                description: description,
                category: "",
                sourceType: "paste",
                originalFileName: "",
                encoding: "UTF-8",
                entryPoint: "",
                rawBytesBase64: "",
                overwrite: true,
                savedId: out savedId,
                error: out error
            );
        }

        public static bool RenameScript(string idOrFileName, string newDisplayName, out string error)
        {
            error = null;
            try
            {
                if (string.IsNullOrEmpty(newDisplayName))
                {
                    error = "新显示名称不能为空";
                    return false;
                }

                newDisplayName = newDisplayName.Trim();

                if (string.IsNullOrEmpty(idOrFileName))
                {
                    error = "目标宏标识不能为空";
                    return false;
                }

                string basFullPath;
                string metaFullPath;
                string scriptId;
                if (!ResolveScriptFiles(idOrFileName, out basFullPath, out metaFullPath, out scriptId, out error))
                {
                    return false;
                }

                // 查重：新显示名称是否已被其他宏占用
                var existingList = ListScripts();
                foreach (var item in existingList)
                {
                    if (string.Equals(item.displayName, newDisplayName, StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(item.id, scriptId, StringComparison.OrdinalIgnoreCase))
                    {
                        error = "DUPLICATE_NAME:已存在同名的宏【" + newDisplayName + "】";
                        return false;
                    }
                }

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = scriptId;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                meta["displayName"] = newDisplayName;
                meta["name"] = newDisplayName;
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                File.WriteAllText(metaFullPath, SimpleJson.Serialize(meta), Encoding.UTF8);
                // 注意：绝不修改 .bas 文件源码内容，确保 originalCodeHash 100% 保持不变！
                try { LeeExcelRibbon.InvalidateRibbon(); } catch { }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool UpdateExecutionResult(string idOrFileName, string result, out string error)
        {
            error = null;
            try
            {
                string basFullPath;
                string metaFullPath;
                string scriptId;
                if (!ResolveScriptFiles(idOrFileName, out basFullPath, out metaFullPath, out scriptId, out error))
                {
                    return false;
                }

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = scriptId;
                    meta["name"] = scriptId;
                    meta["displayName"] = scriptId;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                meta["lastExecutionResult"] = result ?? "未运行";
                meta["lastExecutedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                SafeWriteMetaJson(metaFullPath, SimpleJson.Serialize(meta));
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void SafeWriteMetaJson(string metaFullPath, string jsonContent)
        {
            string tmpPath = metaFullPath + ".tmp";
            File.WriteAllText(tmpPath, jsonContent, Encoding.UTF8);
            if (File.Exists(metaFullPath))
            {
                File.Delete(metaFullPath);
            }
            File.Move(tmpPath, metaFullPath);
        }

        public static bool UpdateScriptTags(string idOrFileName, List<string> newTags, out string error)
        {
            error = null;
            try
            {
                if (string.IsNullOrEmpty(idOrFileName))
                {
                    error = "宏标识不能为空";
                    return false;
                }

                string basFullPath;
                string metaFullPath;
                string scriptId;
                if (!ResolveScriptFiles(idOrFileName, out basFullPath, out metaFullPath, out scriptId, out error))
                {
                    return false;
                }

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = scriptId;
                    meta["displayName"] = scriptId;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                // 规范化 tags 集合
                var cleanTags = new List<string>();
                if (newTags != null)
                {
                    foreach (var t in newTags)
                    {
                        if (!string.IsNullOrEmpty(t))
                        {
                            string trimmed = t.Trim();
                            if (!cleanTags.Contains(trimmed)) cleanTags.Add(trimmed);
                        }
                    }
                }

                meta["tags"] = SimpleJson.Serialize(cleanTags);
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                SafeWriteMetaJson(metaFullPath, SimpleJson.Serialize(meta));
                // 铁律：绝对不修改 .bas 文件源码内容，originalCodeHash 100% 保持不变！
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool UpdateScriptFavorite(string idOrFileName, bool isFavorite, out string error)
        {
            error = null;
            try
            {
                if (string.IsNullOrEmpty(idOrFileName))
                {
                    error = "宏标识为空";
                    return false;
                }

                string basFullPath;
                string metaFullPath;
                string scriptId;
                if (!ResolveScriptFiles(idOrFileName, out basFullPath, out metaFullPath, out scriptId, out error))
                {
                    return false;
                }

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = scriptId;
                    meta["displayName"] = scriptId;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                meta["isFavorite"] = isFavorite ? "true" : "false";
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                SafeWriteMetaJson(metaFullPath, SimpleJson.Serialize(meta));

                // 刷新 Ribbon 收藏菜单
                try
                {
                    LeeExcelRibbon.InvalidateRibbon();
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool AppendRunRecord(string idOrFileName, RunRecordDto record, out string error)
        {
            error = null;
            try
            {
                if (string.IsNullOrEmpty(idOrFileName) || record == null)
                {
                    error = "宏标识或运行记录为空";
                    return false;
                }

                string basFullPath;
                string metaFullPath;
                string scriptId;
                if (!ResolveScriptFiles(idOrFileName, out basFullPath, out metaFullPath, out scriptId, out error))
                {
                    return false;
                }

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = scriptId;
                    meta["displayName"] = scriptId;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                List<RunRecordDto> historyList = new List<RunRecordDto>();
                if (meta.ContainsKey("runHistory") && !string.IsNullOrEmpty(meta["runHistory"]))
                {
                    try
                    {
                        historyList = SimpleJson.DeserializeList<RunRecordDto>(meta["runHistory"]) ?? new List<RunRecordDto>();
                    }
                    catch { }
                }

                // 将本次执行记录追加到列表最前
                historyList.Insert(0, record);

                // 超过 10 条仅淘汰最旧记录，绝不删除对应快照文件！
                if (historyList.Count > 10)
                {
                    historyList = historyList.GetRange(0, 10);
                }

                meta["runHistory"] = SimpleJson.Serialize(historyList);
                meta["lastExecutionResult"] = record.status == "success" ? "执行成功" : (record.status == "blocked" ? "执行前阻断" : "执行失败: " + record.summary);
                meta["lastExecutedAt"] = record.executedAt ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                SafeWriteMetaJson(metaFullPath, SimpleJson.Serialize(meta));
                // 铁律：绝对不修改 .bas 文件源码内容，originalCodeHash 100% 保持不变！
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool DeleteScript(string fileNameOrId, out string error)
        {
            error = null;
            try
            {
                string basPath;
                string metaPath;
                string scriptId;
                if (!ResolveScriptFiles(fileNameOrId, out basPath, out metaPath, out scriptId, out error))
                {
                    return false;
                }

                bool deleted = false;
                if (!string.IsNullOrEmpty(basPath) && File.Exists(basPath))
                {
                    File.Delete(basPath);
                    deleted = true;
                }
                if (!string.IsNullOrEmpty(metaPath) && File.Exists(metaPath))
                {
                    File.Delete(metaPath);
                    deleted = true;
                }

                if (deleted)
                {
                    try { LeeExcelRibbon.InvalidateRibbon(); } catch { }
                    return true;
                }

                error = "脚本文件不存在: " + fileNameOrId;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static string ExtractHeaderField(string content, string fieldName)
        {
            var match = Regex.Match(content, @"'\s*" + Regex.Escape(fieldName) + @"\s*:\s*(.+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return null;
        }
    }
}
