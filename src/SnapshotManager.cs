using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LeeExcel
{
    public class SnapshotItem
    {
        public string id { get; set; }
        public string timestamp { get; set; }
        public string timeDisplay { get; set; }
        public string fileName { get; set; }
        public string originalPath { get; set; }
        public string promptSummary { get; set; }
        public string vbaPreview { get; set; }
    }

    public class SnapshotManager
    {
        private static readonly string BaseBackupDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "LeeExcel",
            "Backups"
        );

        private static string GetWorkbookHash(string fullPath)
        {
            using (var md5 = MD5.Create())
            {
                byte[] bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(fullPath.ToLowerInvariant()));
                var sb = new StringBuilder();
                foreach (byte b in bytes) sb.Append(b.ToString("x2"));
                return sb.ToString().Substring(0, 12);
            }
        }

        public static string GetBackupFolderForWorkbook(string fullPath)
        {
            string hash = GetWorkbookHash(fullPath);
            string folder = Path.Combine(BaseBackupDir, hash);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        public static List<SnapshotItem> LoadSnapshots(string fullPath)
        {
            var list = new List<SnapshotItem>();
            try
            {
                string folder = GetBackupFolderForWorkbook(fullPath);
                string metaFile = Path.Combine(folder, "snapshots.json");
                if (File.Exists(metaFile))
                {
                    string json = File.ReadAllText(metaFile, Encoding.UTF8);
                    list = SimpleJson.DeserializeList<SnapshotItem>(json) ?? new List<SnapshotItem>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadSnapshots error: " + ex.Message);
            }
            return list;
        }

        private static void SaveSnapshots(string fullPath, List<SnapshotItem> list)
        {
            try
            {
                string folder = GetBackupFolderForWorkbook(fullPath);
                string metaFile = Path.Combine(folder, "snapshots.json");
                string json = SimpleJson.Serialize(list);
                File.WriteAllText(metaFile, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SaveSnapshots error: " + ex.Message);
            }
        }

        public static SnapshotItem CreateSnapshot(dynamic workbook, string promptSummary, string vbaPreview)
        {
            string fullPath = workbook.FullName;
            string folder = GetBackupFolderForWorkbook(fullPath);

            string nowStr = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string id = "snap_" + nowStr + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
            string ext = Path.GetExtension(fullPath);
            if (string.IsNullOrEmpty(ext)) ext = ".xlsx";

            string snapshotFileName = id + ext;
            string snapshotFullPath = Path.Combine(folder, snapshotFileName);

            // 原生物理副本保存
            bool saveOk = false;
            try
            {
                workbook.SaveCopyAs(snapshotFullPath);
                saveOk = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SaveCopyAs warning (likely unsaved workbook): " + ex.Message);
            }

            if (!saveOk || !File.Exists(snapshotFullPath) || new FileInfo(snapshotFullPath).Length == 0)
            {
                System.Diagnostics.Debug.WriteLine("快照创建失败：物理备份文件不存在或为空 -> " + snapshotFullPath);
                return null;
            }

            var item = new SnapshotItem
            {
                id = id,
                timestamp = DateTime.Now.ToString("o"),
                timeDisplay = DateTime.Now.ToString("HH:mm:ss"),
                fileName = snapshotFileName,
                originalPath = fullPath,
                promptSummary = promptSummary ?? "自动执行前快照",
                vbaPreview = vbaPreview != null && vbaPreview.Length > 80 ? vbaPreview.Substring(0, 80) + "..." : vbaPreview
            };

            var existing = LoadSnapshots(fullPath);
            existing.Insert(0, item);
            if (existing.Count > 30)
            {
                // 最多保留30个快照，清理旧文件
                for (int i = 30; i < existing.Count; i++)
                {
                    try
                    {
                        string oldFile = Path.Combine(folder, existing[i].fileName);
                        if (File.Exists(oldFile)) File.Delete(oldFile);
                    }
                    catch { }
                }
                existing = existing.GetRange(0, 30);
            }

            SaveSnapshots(fullPath, existing);
            return item;
        }

        public static bool RestoreSnapshot(dynamic app, dynamic currentWorkbook, string snapshotId, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                string originalPath = currentWorkbook.FullName;
                string folder = GetBackupFolderForWorkbook(originalPath);
                var list = LoadSnapshots(originalPath);
                var target = list.Find(x => x.id == snapshotId);
                if (target == null)
                {
                    errorMessage = "未找到指定快照: " + snapshotId;
                    return false;
                }

                string snapshotFullPath = Path.Combine(folder, target.fileName);
                if (!File.Exists(snapshotFullPath) || new FileInfo(snapshotFullPath).Length == 0)
                {
                    errorMessage = "快照物理文件已不存在或损坏(0字节): " + snapshotFullPath;
                    return false;
                }

                // 0. 检查工作簿是否有未保存的手工修改，并在覆盖前尝试保存紧急救援快照
                bool isDirty = false;
                try { isDirty = !(bool)currentWorkbook.Saved; } catch { }

                bool rescueSuccess = false;
                string rescueFileName = "rescue_before_restore_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
                string rescueFullPath = Path.Combine(folder, rescueFileName);
                try
                {
                    currentWorkbook.SaveCopyAs(rescueFullPath);
                    if (File.Exists(rescueFullPath) && new FileInfo(rescueFullPath).Length > 0)
                    {
                        rescueSuccess = true;
                    }
                }
                catch (Exception rescueEx)
                {
                    System.Diagnostics.Debug.WriteLine("Rescue snapshot warning: " + rescueEx.Message);
                }

                // 若工作簿存在未保存的修改且救援快照未能生成，为防数据永久丢失，拒绝盲目覆盖回滚
                if (isDirty && !rescueSuccess)
                {
                    errorMessage = "恢复中止：工作簿存在未保存的手工修改，且紧急救援快照生成失败。为防止手工数据意外丢失，已安全终止回滚。";
                    return false;
                }

                // 1. 关闭当前工作簿（放弃修改）
                currentWorkbook.Close(false);

                if (Path.IsPathRooted(originalPath) && File.Exists(originalPath))
                {
                    // 2. 将快照文件复制覆盖原文件
                    try
                    {
                        File.Copy(snapshotFullPath, originalPath, true);
                    }
                    catch (Exception copyEx)
                    {
                        errorMessage = "文件覆盖失败（原文件可能被其他进程占用）: " + copyEx.Message + "。快照原件与恢复前紧急备份均完好无损。";
                        return false;
                    }

                    // 3. 重新打开工作簿
                    try
                    {
                        app.Workbooks.Open(originalPath);
                    }
                    catch (Exception openEx)
                    {
                        errorMessage = "恢复成功但重新打开文件失败: " + openEx.Message + "。物理文件已恢复为快照，您可手动从 Excel 打开：" + originalPath;
                        return false;
                    }
                }
                else
                {
                    // 未保存的工作簿，直接打开备份物理文件
                    app.Workbooks.Open(snapshotFullPath);
                }
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }
    }
}
