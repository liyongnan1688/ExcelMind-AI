using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
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
    }

    public class ScriptManager
    {
        public static readonly string ScriptsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelMindAI",
            "Scripts"
        );

        static ScriptManager()
        {
            if (!Directory.Exists(ScriptsDir))
            {
                Directory.CreateDirectory(ScriptsDir);
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
                            string dest = Path.Combine(ScriptsDir, Path.GetFileName(file));
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
                            isVerified = lastRes.Contains("成功")
                        });
                    }
                    else
                    {
                        // 兼容旧版纯 .bas 存储格式
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
                            isVerified = false
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

                string baseName = Path.GetFileNameWithoutExtension(idOrFileName);
                string metaFullPath = Path.Combine(ScriptsDir, baseName + ".meta.json");
                string basFullPath = Path.Combine(ScriptsDir, baseName + ".bas");

                if (!File.Exists(basFullPath) && !File.Exists(metaFullPath))
                {
                    error = "未找到目标宏: " + idOrFileName;
                    return false;
                }

                // 查重：新显示名称是否已被其他宏占用
                var existingList = ListScripts();
                foreach (var item in existingList)
                {
                    if (string.Equals(item.displayName, newDisplayName, StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(item.id, baseName, StringComparison.OrdinalIgnoreCase))
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
                    meta["id"] = baseName;
                    meta["createdAt"] = File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                meta["displayName"] = newDisplayName;
                meta["name"] = newDisplayName;
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                File.WriteAllText(metaFullPath, SimpleJson.Serialize(meta), Encoding.UTF8);
                // 注意：绝不修改 .bas 文件源码内容，确保 originalCodeHash 100% 保持不变！
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
                string baseName = Path.GetFileNameWithoutExtension(idOrFileName);
                string metaFullPath = Path.Combine(ScriptsDir, baseName + ".meta.json");
                string basFullPath = Path.Combine(ScriptsDir, baseName + ".bas");

                Dictionary<string, string> meta;
                if (File.Exists(metaFullPath))
                {
                    meta = SimpleJson.ParseFlatObject(File.ReadAllText(metaFullPath, Encoding.UTF8));
                }
                else
                {
                    meta = new Dictionary<string, string>();
                    meta["id"] = baseName;
                    meta["name"] = baseName;
                    meta["displayName"] = baseName;
                    meta["createdAt"] = File.Exists(basFullPath) ? File.GetCreationTime(basFullPath).ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string code = File.Exists(basFullPath) ? File.ReadAllText(basFullPath, Encoding.UTF8) : "";
                    meta["originalCodeHash"] = ComputeSha256(code);
                }

                meta["lastExecutionResult"] = result ?? "未运行";
                meta["lastExecutedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                meta["updatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                File.WriteAllText(metaFullPath, SimpleJson.Serialize(meta), Encoding.UTF8);
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
                string baseName = Path.GetFileNameWithoutExtension(fileNameOrId);
                string basPath = Path.Combine(ScriptsDir, baseName + ".bas");
                string metaPath = Path.Combine(ScriptsDir, baseName + ".meta.json");

                bool deleted = false;
                if (File.Exists(basPath))
                {
                    File.Delete(basPath);
                    deleted = true;
                }
                if (File.Exists(metaPath))
                {
                    File.Delete(metaPath);
                    deleted = true;
                }

                if (deleted) return true;

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
