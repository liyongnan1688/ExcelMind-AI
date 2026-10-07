using System;
using System.IO;
using System.Runtime.InteropServices;

namespace LeeExcel.Tests
{
    public class VerifySaveCopyAsDirtyState
    {
        public static int Main(string[] args)
        {
            Console.WriteLine("=== [Verification Tool] Excel SaveCopyAs Behavior on Dirty (Unsaved) State ===");

            string testRunId = "run_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string testDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".artifacts", "tests", testRunId));
            Directory.CreateDirectory(testDir);

            string originalPath = Path.Combine(testDir, "original_test.xlsx");
            string snapshotPath = Path.Combine(testDir, "snapshot_copy.xlsx");

            object excelAppObj = null;
            dynamic app = null;
            dynamic workbooks = null;
            dynamic wb = null;
            dynamic readWb = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Console.WriteLine("[ERROR] Excel is not installed on this system.");
                    return 2;
                }

                excelAppObj = Activator.CreateInstance(excelType);
                app = excelAppObj;
                app.Visible = false;
                app.DisplayAlerts = false;

                workbooks = app.Workbooks;
                wb = workbooks.Add();

                // 步骤 1: 写入初始已保存内容
                wb.Sheets[1].Range("A1").Value = "SAVED_INITIAL_STATE";
                wb.SaveAs(originalPath);
                Console.WriteLine("[Step 1] Created and saved original workbook with A1='SAVED_INITIAL_STATE'");
                Console.WriteLine("         wb.Saved immediately after SaveAs: " + wb.Saved);

                // 步骤 2: 产生未保存的内存脏数据 (Dirty Modification)
                wb.Sheets[1].Range("A1").Value = "DIRTY_UNSAVED_MEMORY_EDIT";
                Console.WriteLine("[Step 2] Modified A1 to 'DIRTY_UNSAVED_MEMORY_EDIT' in memory without calling Save().");
                Console.WriteLine("         wb.Saved is now: " + wb.Saved);

                // 步骤 3: 模拟 SnapshotManager.CreateSnapshot 执行 SaveCopyAs
                Console.WriteLine("[Step 3] Calling wb.SaveCopyAs(snapshotPath)...");
                wb.SaveCopyAs(snapshotPath);
                Console.WriteLine("         SaveCopyAs completed. File exists: " + File.Exists(snapshotPath));
                Console.WriteLine("         wb.Saved after SaveCopyAs: " + wb.Saved);

                // 关闭原工作簿（不保存脏修改，还原磁盘文件为 SAVED_INITIAL_STATE）
                wb.Close(false);
                wb = null;

                // 步骤 4: 分别重新打开原磁盘文件和快照副本，比对 A1 内容
                readWb = app.Workbooks.Open(originalPath);
                string diskValue = Convert.ToString(readWb.Sheets[1].Range("A1").Value);
                readWb.Close(false);
                readWb = null;

                readWb = app.Workbooks.Open(snapshotPath);
                string snapshotValue = Convert.ToString(readWb.Sheets[1].Range("A1").Value);
                readWb.Close(false);
                readWb = null;

                Console.WriteLine("================ RESULT COMPARISON ================");
                Console.WriteLine("Disk original file A1 value:     [" + diskValue + "]");
                Console.WriteLine("Snapshot copy file A1 value:     [" + snapshotValue + "]");

                bool snapshotIncludesDirty = (snapshotValue == "DIRTY_UNSAVED_MEMORY_EDIT");
                bool diskRetainedSaved = (diskValue == "SAVED_INITIAL_STATE");

                Console.WriteLine("Does snapshot capture unsaved in-memory edits? -> " + snapshotIncludesDirty);
                Console.WriteLine("Was original disk file protected from premature overwrite? -> " + diskRetainedSaved);

                if (snapshotIncludesDirty && diskRetainedSaved)
                {
                    Console.WriteLine("[SUCCESS] Verified: SaveCopyAs serializes active IN-MEMORY state (including unsaved edits) into the backup copy without modifying original file.");
                    return 0;
                }
                else
                {
                    Console.WriteLine("[FAIL] Unexpected behavior observed.");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EXCEPTION] " + ex.ToString());
                return 1;
            }
            finally
            {
                // 确保严格释放所有 COM 对象并退出专属测试实例，绝不杀死用户的已有 Excel 进程
                try { if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); } } catch { }
                try { if (readWb != null) { readWb.Close(false); Marshal.ReleaseComObject(readWb); } } catch { }
                try { if (workbooks != null) Marshal.ReleaseComObject(workbooks); } catch { }
                try { if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); } } catch { }
                try { if (excelAppObj != null) Marshal.ReleaseComObject(excelAppObj); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
