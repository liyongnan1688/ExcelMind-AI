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
        public string filePath
        {
            get
            {
                if (string.IsNullOrEmpty(originalPath) || string.IsNullOrEmpty(fileName)) return string.Empty;
                return Path.Combine(SnapshotManager.GetBackupFolderForWorkbook(originalPath), fileName);
            }
        }
        public string promptSummary { get; set; }
        public string vbaPreview { get; set; }
    }

    public class SnapshotManager
    {
        private static readonly string BaseBackupDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelMindAI",
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
                try
                {
                    string oldFolder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "LeeExcel",
                        "Backups",
                        hash
                    );
                    if (Directory.Exists(oldFolder))
                    {
                        Directory.CreateDirectory(folder);
                        foreach (var f in Directory.GetFiles(oldFolder))
                        {
                            string dest = Path.Combine(folder, Path.GetFileName(f));
                            if (!File.Exists(dest)) File.Copy(f, dest);
                        }
                    }
                }
                catch { }

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }
            return folder;
        }

        public static bool SnapshotExists(string fullPath, string snapshotId)
        {
            if (string.IsNullOrEmpty(snapshotId)) return false;
            try
            {
                if (!string.IsNullOrEmpty(fullPath))
                {
                    string folder = GetBackupFolderForWorkbook(fullPath);
                    if (Directory.Exists(folder))
                    {
                        var matches = Directory.GetFiles(folder, snapshotId + ".*");
                        if (matches.Length > 0 && new FileInfo(matches[0]).Length > 0) return true;
                    }
                }

                if (Directory.Exists(BaseBackupDir))
                {
                    var all = Directory.GetFiles(BaseBackupDir, snapshotId + ".*", SearchOption.AllDirectories);
                    return all.Length > 0 && new FileInfo(all[0]).Length > 0;
                }
            }
            catch { }
            return false;
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

        private static bool SaveSnapshots(string fullPath, List<SnapshotItem> list)
        {
            string tempFile = null;
            try
            {
                string folder = GetBackupFolderForWorkbook(fullPath);
                string metaFile = Path.Combine(folder, "snapshots.json");
                string json = SimpleJson.Serialize(list);
                tempFile = metaFile + ".tmp." + Guid.NewGuid().ToString("N");
                File.WriteAllText(tempFile, json, Encoding.UTF8);
                if (File.Exists(metaFile))
                {
                    File.Replace(tempFile, metaFile, null);
                }
                else
                {
                    File.Move(tempFile, metaFile);
                }
                tempFile = null;
                string verified = File.ReadAllText(metaFile, Encoding.UTF8);
                if (!string.Equals(verified, json, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("快照元数据回读校验失败");
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SaveSnapshots error: " + ex.Message);
                return false;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempFile))
                {
                    try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                }
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

            List<SnapshotItem> retained = existing.Count > 30 ? existing.GetRange(0, 30) : existing;
            if (!SaveSnapshots(fullPath, retained))
            {
                try { if (File.Exists(snapshotFullPath)) File.Delete(snapshotFullPath); } catch { }
                throw new InvalidOperationException("快照元数据保存失败，已中止执行以避免不可恢复修改。");
            }

            for (int i = retained.Count; i < existing.Count; i++)
            {
                try
                {
                    string oldFile = Path.Combine(folder, existing[i].fileName);
                    if (File.Exists(oldFile)) File.Delete(oldFile);
                }
                catch { }
            }
            return item;
        }

        private static void TryReopenOriginalWorkbook(dynamic app, string originalPath)
        {
            if (app == null || string.IsNullOrEmpty(originalPath)) return;
            try { app.Workbooks.Open(originalPath); } catch { }
        }

        public static bool RestoreSnapshot(dynamic app, dynamic currentWorkbook, string snapshotId, out string errorMessage)
        {
            errorMessage = null;
            string originalPath = null;
            bool workbookClosed = false;
            try
            {
                originalPath = currentWorkbook.FullName;
                string folder = GetBackupFolderForWorkbook(originalPath);
                var list = LoadSnapshots(originalPath);
                var target = list.Find(x => x.id == snapshotId);
                if (target == null)
                {
                    errorMessage = "未找到指定快照: " + snapshotId;
                    return false;
                }

                if (string.IsNullOrWhiteSpace(target.fileName) ||
                    !string.Equals(Path.GetFileName(target.fileName), target.fileName, StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = "快照元数据中的文件名非法，已拒绝恢复。";
                    return false;
                }

                string folderFullPath = Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                string snapshotFullPath = Path.GetFullPath(Path.Combine(folder, target.fileName));
                if (!snapshotFullPath.StartsWith(folderFullPath, StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = "快照路径越出备份目录，已拒绝恢复。";
                    return false;
                }
                if (!File.Exists(snapshotFullPath) || new FileInfo(snapshotFullPath).Length == 0)
                {
                    errorMessage = "快照物理文件已不存在或损坏(0字节): " + snapshotFullPath;
                    return false;
                }

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

                if (isDirty && !rescueSuccess)
                {
                    errorMessage = "恢复中止：工作簿存在未保存的手工修改，且紧急救援快照生成失败。为防止手工数据意外丢失，已安全终止回滚。";
                    return false;
                }

                currentWorkbook.Close(false);
                workbookClosed = true;

                if (Path.IsPathRooted(originalPath) && File.Exists(originalPath))
                {
                    try
                    {
                        File.Copy(snapshotFullPath, originalPath, true);
                    }
                    catch (Exception copyEx)
                    {
                        TryReopenOriginalWorkbook(app, originalPath);
                        errorMessage = "文件覆盖失败（原文件可能被其他进程占用）: " + copyEx.Message + "。快照原件与恢复前紧急备份均完好无损。";
                        return false;
                    }

                    try
                    {
                        app.Workbooks.Open(originalPath);
                        workbookClosed = false;
                    }
                    catch (Exception openEx)
                    {
                        errorMessage = "恢复成功但重新打开文件失败: " + openEx.Message + "。物理文件已恢复为快照，您可手动从 Excel 打开：" + originalPath;
                        return false;
                    }
                }
                else
                {
                    app.Workbooks.Open(snapshotFullPath);
                    workbookClosed = false;
                }
                return true;
            }
            catch (Exception ex)
            {
                if (workbookClosed) TryReopenOriginalWorkbook(app, originalPath);
                errorMessage = ex.Message;
                return false;
            }
        }
}
}