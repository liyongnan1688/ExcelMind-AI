using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class ScriptInfo
    {
        public string name { get; set; }
        public string fileName { get; set; }
        public string filePath { get; set; }
        public string createdAt { get; set; }
        public string description { get; set; }
        public string code { get; set; }
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
                    string content = File.ReadAllText(f, Encoding.UTF8);
                    string desc = ExtractHeaderField(content, "Description") ?? "自定义保存脚本";
                    string created = ExtractHeaderField(content, "Created At") ?? File.GetCreationTime(f).ToString("yyyy-MM-dd HH:mm");
                    string name = Path.GetFileNameWithoutExtension(f);

                    list.Add(new ScriptInfo
                    {
                        name = name,
                        fileName = Path.GetFileName(f),
                        filePath = f,
                        createdAt = created,
                        description = desc,
                        code = content
                    });
                }
                catch { }
            }
            return list;
        }

        public static bool SaveScript(string name, string code, string description, out string error)
        {
            error = null;
            try
            {
                if (string.IsNullOrEmpty(name))
                {
                    error = "脚本名称不能为空";
                    return false;
                }

                // 移除非法字符
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    name = name.Replace(c, '_');
                }

                string fileName = name.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) ? name : name + ".bas";
                string fullPath = Path.Combine(ScriptsDir, fileName);

                var sb = new StringBuilder();
                sb.AppendLine("' ==============================================================");
                sb.AppendLine("' Script Name: " + name);
                sb.AppendLine("' Created At : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                sb.AppendLine("' Description: " + (description ?? "用户保存的自动化宏"));
                sb.AppendLine("' ==============================================================");
                sb.AppendLine(code);

                File.WriteAllText(fullPath, sb.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static bool DeleteScript(string fileName, out string error)
        {
            error = null;
            try
            {
                string fullPath = Path.Combine(ScriptsDir, fileName);
                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                    return true;
                }
                error = "脚本文件不存在: " + fileName;
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
