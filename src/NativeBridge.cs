using System;
using System.Collections.Generic;

namespace LeeExcel
{
    public class WorkbookInfoDto
    {
        public string name { get; set; }
        public string fullName { get; set; }
        public bool isSaved { get; set; }
        public List<string> sheets { get; set; }
        public List<SnapshotItem> snapshots { get; set; }
    }

    public class ExecutionDataDto
    {
        public string summary { get; set; }
        public string error { get; set; }
        public double elapsedMs { get; set; }
        public SnapshotItem snapshot { get; set; }
        public string vbaCode { get; set; }
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
                        return HandleGetWorkbookInfo(app);

                    case "execute_vba":
                        return HandleExecuteVba(req, app);

                    case "create_snapshot":
                        return HandleCreateSnapshot(req, app);

                    case "restore_snapshot":
                        return HandleRestoreSnapshot(req, app);

                    case "list_snapshots":
                        return HandleListSnapshots(app);

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

        private static string HandleGetWorkbookInfo(dynamic app)
        {
            try
            {
                dynamic wb = null;
                try
                {
                    wb = app.ActiveWorkbook;
                }
                catch { }

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
                            sheets = new List<string>(),
                            snapshots = new List<SnapshotItem>()
                        }
                    });
                }

                string name = (string)wb.Name;
                string fullName = (string)wb.FullName;
                bool isSaved = !string.IsNullOrEmpty((string)wb.Path);

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
            dynamic wb = app.ActiveWorkbook;
            if (wb == null)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "未检测到活动工作簿，无法执行"
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

            // 1. 运行前自动执行整本物理副本快照 (未保存工作簿也尽力导出临时快照)
            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(wb, prompt, code);
            }
            catch (Exception snapEx)
            {
                System.Diagnostics.Debug.WriteLine("快照创建告警: " + snapEx.Message);
            }

            // 2. 执行动态 VBA
            var result = VbaRunner.RunVbaCode(app, wb, code);

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
                    vbaCode = code
                }
            });
        }

        private static string HandleCreateSnapshot(Dictionary<string, string> req, dynamic app)
        {
            dynamic wb = app.ActiveWorkbook;
            if (wb == null || string.IsNullOrEmpty((string)wb.Path))
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "create_snapshot",
                    error = "工作簿未保存，无法创建物理备份"
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
            dynamic wb = app.ActiveWorkbook;
            if (wb == null)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    action = "restore_snapshot",
                    error = "未找到活动工作簿"
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

        private static string HandleListSnapshots(dynamic app)
        {
            dynamic wb = app.ActiveWorkbook;
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
