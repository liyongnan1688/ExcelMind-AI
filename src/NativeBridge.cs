using System;
using System.Collections.Generic;

namespace LeeExcel
{
    public class WorkbookInfoDto
    {
        public string name { get; set; }
        public string fullName { get; set; }
        public bool isSaved { get; set; }
        public string activeSheetName { get; set; }
        public string usedRangeAddress { get; set; }
        public List<string> sheets { get; set; }
        public List<SnapshotItem> snapshots { get; set; }
    }

    public class ExecutionDataDto
    {
        public string summary { get; set; }
        public string error { get; set; }
        public double elapsedMs { get; set; }
        public SnapshotItem snapshot { get; set; }
        public string originalVbaCode { get; set; }
        public string executedVbaCode { get; set; }
        public List<string> transformSteps { get; set; }
        public WorkbookReadback readback { get; set; }
        public string targetWorkbookName { get; set; }
    }

    public class BridgeResponse
    {
        public bool ok { get; set; }
        public string action { get; set; }
        public string message { get; set; }
        public object data { get; set; }
        public string error { get; set; }
    }

    public class NativeBridge
    {
        public static string Dispatch(string jsonString, dynamic app)
        {
            try
            {
                var req = SimpleJson.ParseFlatObject(jsonString);
                string action = req.ContainsKey("action") ? req["action"] : "";

                switch (action)
                {
                    case "get_workbook_info":
                        return HandleGetWorkbookInfo(req, app);

                    case "execute_vba":
                        return HandleExecuteVba(req, app);

                    case "create_snapshot":
                        return HandleCreateSnapshot(req, app);

                    case "restore_snapshot":
                        return HandleRestoreSnapshot(req, app);

                    case "list_snapshots":
                        return HandleListSnapshots(req, app);

                    case "list_scripts":
                        return HandleListScripts();

                    case "save_script":
                        return HandleSaveScript(req);

                    case "delete_script":
                        return HandleDeleteScript(req);

                    default:
                        return SimpleJson.Serialize(new BridgeResponse
                        {
                            ok = false,
                            action = action,
                            error = "未知请求动作: " + action
                        });
                }
            }
            catch (Exception ex)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    error = "Bridge异常: " + ex.Message
                });
            }
        }

        public static dynamic FindTargetWorkbook(dynamic app, string targetFullName, string targetName)
        {
            if (app == null) return null;

            // 1. 优先按 FullName 匹配
            if (!string.IsNullOrEmpty(targetFullName))
            {
                try
                {
                    foreach (dynamic wb in app.Workbooks)
                    {
                        if (string.Equals((string)wb.FullName, targetFullName, StringComparison.OrdinalIgnoreCase))
                        {
                            return wb;
                        }
                    }
                }
                catch { }
            }

            // 2. 次选按 Name 匹配
            if (!string.IsNullOrEmpty(targetName))
            {
                try
                {
                    foreach (dynamic wb in app.Workbooks)
                    {
                        if (string.Equals((string)wb.Name, targetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return wb;
                        }
                    }
                }
                catch { }
            }

            // 3. 若均未传参且存在活动工作簿，作为保底
            if (string.IsNullOrEmpty(targetFullName) && string.IsNullOrEmpty(targetName))
            {
                try
                {
                    return app.ActiveWorkbook;
                }
                catch { }
            }

            return null;
        }

        private static string HandleGetWorkbookInfo(Dictionary<string, string> req, dynamic app)
        {
            try
            {
                string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
                string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

                dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

                if (wb == null)
                {
                    return SimpleJson.Serialize(new BridgeResponse
                    {
                        ok = true,
                        action = "get_workbook_info",
                        data = new WorkbookInfoDto
                        {
                            name = "未检测到活动工作簿",
                            fullName = "",
                            isSaved = false,
                            activeSheetName = "",
                            usedRangeAddress = "",
                            sheets = new List<string>(),
                            snapshots = new List<SnapshotItem>()
                        }
                    });
                }

                string name = (string)wb.Name;
                string fullName = (string)wb.FullName;
                bool isSaved = !string.IsNullOrEmpty((string)wb.Path);
                string activeSheetName = "";
                string usedRangeAddress = "";

                try
                {
                    dynamic sh = wb.ActiveSheet;
                    if (sh != null)
                    {
                        activeSheetName = (string)sh.Name;
                        dynamic ur = sh.UsedRange;
                        if (ur != null)
                        {
                            string rawAddr = (string)ur.Address;
                            usedRangeAddress = rawAddr != null ? rawAddr.Replace("$", "") : "";
                        }
                    }
                }
                catch { }

                var sheets = new List<string>();
                try
                {
                    foreach (dynamic sh in wb.Sheets)
                    {
                        sheets.Add((string)sh.Name);
                    }
                }
                catch { }

                var snapshots = isSaved ? SnapshotManager.LoadSnapshots(fullName) : new List<SnapshotItem>();

                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = true,
                    action = "get_workbook_info",
                    data = new WorkbookInfoDto
                    {
                        name = name,
                        fullName = fullName,
                        isSaved = isSaved,
                        activeSheetName = activeSheetName,
                        usedRangeAddress = usedRangeAddress,
                        sheets = sheets,
                        snapshots = snapshots
                    }
                });
            }
            catch (Exception ex)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "get_workbook_info",
                    error = ex.Message
                });
            }
        }

        private static string HandleExecuteVba(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);

            if (targetWb == null)
            {
                string targetDesc = !string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName;
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "未找到目标工作簿 (" + (targetDesc ?? "未指定") + ")，请确认该工作簿已在 Excel 中打开。"
                });
            }

            string code = req.ContainsKey("code") ? req["code"] : "";
            string prompt = req.ContainsKey("prompt") ? req["prompt"] : "自然语言操作";

            if (string.IsNullOrEmpty(code))
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "传入的 VBA 代码为空"
                });
            }

            // 1. 运行前自动执行整本物理副本快照 (针对目标工作簿 targetWb)
            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(targetWb, prompt, code);
            }
            catch (Exception snapEx)
            {
                System.Diagnostics.Debug.WriteLine("快照创建告警: " + snapEx.Message);
            }

            // 2. 在目标工作簿中执行动态 VBA 并读回实际状态
            var result = VbaRunner.RunVbaCode(app, targetWb, code);

            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = result.success,
                action = "execute_vba",
                message = result.summary,
                error = result.error,
                data = new ExecutionDataDto
                {
                    summary = result.summary,
                    error = result.error,
                    elapsedMs = result.elapsedMs,
                    snapshot = snap,
                    originalVbaCode = result.originalVbaCode,
                    executedVbaCode = result.executedVbaCode,
                    transformSteps = result.transformSteps,
                    readback = result.readback,
                    targetWorkbookName = (string)targetWb.Name
                }
            });
        }

        private static string HandleCreateSnapshot(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null || string.IsNullOrEmpty((string)wb.Path))
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "create_snapshot",
                    error = "目标工作簿未保存或未找到，无法创建物理备份"
                });
            }

            string prompt = req.ContainsKey("prompt") ? req["prompt"] : "手动创建快照";
            var snap = SnapshotManager.CreateSnapshot(wb, prompt, "");
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = true,
                action = "create_snapshot",
                data = snap
            });
        }

        private static string HandleRestoreSnapshot(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "restore_snapshot",
                    error = "未找到目标工作簿"
                });
            }

            string snapId = req.ContainsKey("snapshotId") ? req["snapshotId"] : "";
            string err;
            bool ok = SnapshotManager.RestoreSnapshot(app, wb, snapId, out err);
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = ok,
                action = "restore_snapshot",
                message = ok ? "已成功恢复到快照执行前的整本工作簿状态！" : err,
                error = err
            });
        }

        private static string HandleListSnapshots(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null || string.IsNullOrEmpty((string)wb.Path))
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = true,
                    action = "list_snapshots",
                    data = new List<SnapshotItem>()
                });
            }

            var list = SnapshotManager.LoadSnapshots((string)wb.FullName);
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = true,
                action = "list_snapshots",
                data = list
            });
        }

        private static string HandleListScripts()
        {
            var list = ScriptManager.ListScripts();
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = true,
                action = "list_scripts",
                data = list
            });
        }

        private static string HandleSaveScript(Dictionary<string, string> req)
        {
            string name = req.ContainsKey("name") ? req["name"] : "";
            string code = req.ContainsKey("code") ? req["code"] : "";
            string desc = req.ContainsKey("description") ? req["description"] : "";

            string err;
            bool ok = ScriptManager.SaveScript(name, code, desc, out err);
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = ok,
                action = "save_script",
                message = ok ? "脚本已成功保存到本地“我的脚本”库！" : err,
                error = err
            });
        }

        private static string HandleDeleteScript(Dictionary<string, string> req)
        {
            string fileName = req.ContainsKey("fileName") ? req["fileName"] : "";
            string err;
            bool ok = ScriptManager.DeleteScript(fileName, out err);
            return SimpleJson.Serialize(new BridgeResponse
            {
                ok = ok,
                action = "delete_script",
                message = ok ? "脚本已删除" : err,
                error = err
            });
        }
    }
}
