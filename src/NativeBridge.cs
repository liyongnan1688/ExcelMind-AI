using System;
using System.Collections.Generic;
using System.IO;

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
        public string rawModelResponse { get; set; }
        public string originalVbaCode { get; set; }
        public string executedVbaCode { get; set; }
        public string wrapperCode { get; set; }
        public string originalCodeHash { get; set; }
        public string executedCodeHash { get; set; }
        public bool isSourceIdentical { get; set; }
        public List<string> transformSteps { get; set; }
        public WorkbookReadback readback { get; set; }
        public string targetWorkbookName { get; set; }
        public string targetWorkbookFullName { get; set; }
        public string precheckStatus { get; set; }
        public string executionPhase { get; set; }
        public string riskNotice { get; set; }
        public string injectedModuleName { get; set; }
        public string failureStage { get; set; }
        public string rawErrorCode { get; set; }
        public string errorTriggerPoint { get; set; }
        public int? vbaErrNumber { get; set; }
        public string vbaErrDescription { get; set; }
        public string comHResult { get; set; }
        public string hostExecutionPhase { get; set; }
        public bool isPartiallyModified { get; set; }
        public bool hostStateRestored { get; set; }
        public string hostStateRestoreDetails { get; set; }
        public bool? origScreenUpdating { get; set; }
        public bool? origDisplayAlerts { get; set; }
        public bool? origEnableEvents { get; set; }
        public int? origCalculation { get; set; }
        public bool? restoredScreenUpdating { get; set; }
        public bool? restoredDisplayAlerts { get; set; }
        public bool? restoredEnableEvents { get; set; }
        public int? restoredCalculation { get; set; }
    }

    public class BridgeResponse
    {
        public bool ok { get; set; }
        public string action { get; set; }
        public string requestId { get; set; }
        public string message { get; set; }
        public object data { get; set; }
        public string error { get; set; }
    }

    public class NativeBridge
    {
        public static string Dispatch(string jsonString, dynamic app)
        {
            string requestId = null;
            try
            {
                var req = SimpleJson.ParseFlatObject(jsonString);
                string action = req.ContainsKey("action") ? req["action"] : "";
                if (req.ContainsKey("requestId")) requestId = req["requestId"];

                BridgeResponse resp;
                switch (action)
                {
                    case "get_workbook_info":
                        resp = HandleGetWorkbookInfo(req, app);
                        break;

                    case "get_selection_context":
                        resp = HandleGetSelectionContext(req, app);
                        break;

                    case "execute_vba":
                        resp = HandleExecuteVba(req, app);
                        break;

                    case "create_snapshot":
                        resp = HandleCreateSnapshot(req, app);
                        break;

                    case "restore_snapshot":
                        resp = HandleRestoreSnapshot(req, app);
                        break;

                    case "list_snapshots":
                        resp = HandleListSnapshots(req, app);
                        break;

                    case "list_scripts":
                        resp = HandleListScripts();
                        break;

                    case "save_script":
                        resp = HandleSaveScript(req);
                        break;

                    case "rename_script":
                        resp = HandleRenameScript(req);
                        break;

                    case "update_script_result":
                        resp = HandleUpdateScriptResult(req);
                        break;

                    case "delete_script":
                        resp = HandleDeleteScript(req);
                        break;

                    case "unlock_workbook":
                        resp = HandleUnlockWorkbook(req, app);
                        break;

                    default:
                        resp = new BridgeResponse
                        {
                            ok = false,
                            action = action,
                            error = "未知请求动作: " + action
                        };
                        break;
                }

                if (resp != null)
                {
                    resp.requestId = requestId;
                }
                return SimpleJson.Serialize(resp);
            }
            catch (Exception ex)
            {
                return SimpleJson.Serialize(new BridgeResponse
                {
                    ok = false,
                    requestId = requestId,
                    error = "Bridge异常: " + ex.Message
                });
            }
        }

        public static dynamic FindTargetWorkbook(dynamic app, string targetFullName, string targetName)
        {
            if (app == null) return null;

            if (string.Equals(targetName, "未检测到活动工作簿", StringComparison.OrdinalIgnoreCase))
            {
                targetName = "";
                targetFullName = "";
            }

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

            // 3. 保底：若目标未找到（例如被另存为改名、或原本为空），自动采用当前活动工作簿
            try
            {
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb != null)
                {
                    return activeWb;
                }
            }
            catch { }

            // 4. 若 ActiveWorkbook 为空但存在打开的工作簿，取首个工作簿
            try
            {
                if (app.Workbooks != null && app.Workbooks.Count > 0)
                {
                    return app.Workbooks[1];
                }
            }
            catch { }

            return null;
        }

        private static BridgeResponse HandleGetWorkbookInfo(Dictionary<string, string> req, dynamic app)
        {
            try
            {
                string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
                string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

                if (string.Equals(targetName, "未检测到活动工作簿", StringComparison.OrdinalIgnoreCase))
                {
                    targetName = "";
                    targetFullName = "";
                }

                dynamic wb = null;

                // 优先直接获取当前前台活动的 ActiveWorkbook（实时感知另存为重命名、新建、切换窗口）
                try
                {
                    if (app != null)
                    {
                        wb = app.ActiveWorkbook;
                    }
                }
                catch { }

                // 若 ActiveWorkbook 为空（例如失去焦点），按参数查找特定工作簿
                if (wb == null)
                {
                    wb = FindTargetWorkbook(app, targetFullName, targetName);
                }

                // 再次保底：如果依然为空，取首个打开的工作簿
                if (wb == null && app != null)
                {
                    try
                    {
                        if (app.Workbooks != null && app.Workbooks.Count > 0)
                        {
                            wb = app.Workbooks[1];
                        }
                    }
                    catch { }
                }

                if (wb == null)
                {
                    return new BridgeResponse
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
                    };
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

                return new BridgeResponse
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
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_workbook_info",
                    error = ex.Message
                };
            }
        }

        private static BridgeResponse HandleExecuteVba(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);

            if (targetWb == null)
            {
                string targetDesc = !string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName;
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "未找到目标工作簿 (" + (targetDesc ?? "未指定") + ")，请确认该工作簿已在 Excel 中打开。"
                };
            }

            string code = req.ContainsKey("code") ? req["code"] : "";
            string prompt = req.ContainsKey("prompt") ? req["prompt"] : "自然语言操作";

            if (string.IsNullOrEmpty(code))
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "传入的 VBA 代码为空"
                };
            }

            string rawModelResponse = req.ContainsKey("rawModelResponse") ? req["rawModelResponse"] : "";

            // 1. 运行前自动执行整本物理副本快照 (针对目标工作簿 targetWb)
            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(targetWb, prompt, code);
            }
            catch (Exception snapEx)
            {
                System.Diagnostics.Debug.WriteLine("快照创建失败: " + snapEx.Message);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "执行已安全中止：执行前目标工作簿快照备份失败 (" + snapEx.Message + ")。为保障数据可回滚安全，拒绝执行代码。"
                };
            }

            if (snap == null || string.IsNullOrEmpty(snap.id) || string.IsNullOrEmpty(snap.fileName))
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "执行已安全中止：目标工作簿物理快照备份未能成功生成有效副本。请确认工作簿已保存并具有有效磁盘路径。为保障数据可回滚安全，拒绝执行代码。"
                };
            }

            try
            {
                string wbPathForFolder = !string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName;
                string snapFolder = SnapshotManager.GetBackupFolderForWorkbook(wbPathForFolder);
                string snapFile = Path.Combine(snapFolder, snap.fileName);
                if (!File.Exists(snapFile) || new FileInfo(snapFile).Length == 0)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "execute_vba",
                        error = "执行已安全中止：快照物理文件未能有效写入磁盘（文件不存在或为0字节）。拒绝执行代码以防数据丢失。"
                    };
                }
            }
            catch (Exception checkEx)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = "执行已安全中止：快照文件物理校验异常 (" + checkEx.Message + ")。拒绝执行代码以防数据丢失。"
                };
            }

            // 2. 在目标工作簿中执行动态 VBA 并读回实际状态
            string entryPoint = req.ContainsKey("entryPoint") ? req["entryPoint"] : null;
            var result = VbaRunner.RunVbaCode(app, targetWb, code, rawModelResponse, entryPoint);

            string scriptId = req.ContainsKey("scriptId") ? req["scriptId"] : (req.ContainsKey("fileName") ? req["fileName"] : null);
            if (!string.IsNullOrEmpty(scriptId))
            {
                try
                {
                    string resStatus = result.success ? "执行成功" : ("执行失败: " + (result.summary ?? result.error));
                    string errIgnored;
                    ScriptManager.UpdateExecutionResult(scriptId, resStatus, out errIgnored);
                }
                catch { }
            }

            // 安全获取工作簿名称与路径：宏可能关闭了工作簿导致 COM 引用失效
            string resolvedTargetName = targetName;
            string resolvedTargetFullName = targetFullName;
            try
            {
                resolvedTargetName = (string)targetWb.Name;
                resolvedTargetFullName = (string)targetWb.FullName;
            }
            catch { }

            return new BridgeResponse
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
                    rawModelResponse = result.rawModelResponse,
                    originalVbaCode = result.originalVbaCode,
                    executedVbaCode = result.executedVbaCode,
                    wrapperCode = result.wrapperCode,
                    originalCodeHash = result.originalCodeHash,
                    executedCodeHash = result.executedCodeHash,
                    isSourceIdentical = result.isSourceIdentical,
                    transformSteps = result.transformSteps,
                    readback = result.readback,
                    targetWorkbookName = resolvedTargetName,
                    targetWorkbookFullName = resolvedTargetFullName,
                    precheckStatus = result.precheckStatus,
                    executionPhase = result.executionPhase,
                    riskNotice = result.riskNotice,
                    injectedModuleName = result.injectedModuleName,
                    failureStage = result.failureStage,
                    rawErrorCode = result.rawErrorCode,
                    errorTriggerPoint = result.errorTriggerPoint,
                    vbaErrNumber = result.vbaErrNumber,
                    vbaErrDescription = result.vbaErrDescription,
                    comHResult = result.comHResult,
                    hostExecutionPhase = result.hostExecutionPhase,
                    isPartiallyModified = result.isPartiallyModified,
                    hostStateRestored = result.hostStateRestored,
                    hostStateRestoreDetails = result.hostStateRestoreDetails,
                    origScreenUpdating = result.origScreenUpdating,
                    origDisplayAlerts = result.origDisplayAlerts,
                    origEnableEvents = result.origEnableEvents,
                    origCalculation = result.origCalculation,
                    restoredScreenUpdating = result.restoredScreenUpdating,
                    restoredDisplayAlerts = result.restoredDisplayAlerts,
                    restoredEnableEvents = result.restoredEnableEvents,
                    restoredCalculation = result.restoredCalculation
                }
            };
        }

        private static BridgeResponse HandleCreateSnapshot(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null || string.IsNullOrEmpty((string)wb.Path))
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "create_snapshot",
                    error = "目标工作簿未保存或未找到，无法创建物理备份"
                };
            }

            string prompt = req.ContainsKey("prompt") ? req["prompt"] : "手动创建快照";
            var snap = SnapshotManager.CreateSnapshot(wb, prompt, "");
            return new BridgeResponse
            {
                ok = true,
                action = "create_snapshot",
                data = snap
            };
        }

        private static BridgeResponse HandleRestoreSnapshot(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "restore_snapshot",
                    error = "未找到目标工作簿"
                };
            }

            string snapId = req.ContainsKey("snapshotId") ? req["snapshotId"] : "";
            string err;
            bool ok = SnapshotManager.RestoreSnapshot(app, wb, snapId, out err);
            if (ok)
            {
                try { VbaRunner.UnlockWorkbook((string)wb.Name); } catch { }
            }
            return new BridgeResponse
            {
                ok = ok,
                action = "restore_snapshot",
                message = ok ? "已成功恢复到快照执行前的整本工作簿状态，锁定已自动解除！" : err,
                error = err
            };
        }

        private static BridgeResponse HandleUnlockWorkbook(Dictionary<string, string> req, dynamic app)
        {
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            if (!string.IsNullOrEmpty(targetName))
            {
                VbaRunner.UnlockWorkbook(targetName);
            }
            return new BridgeResponse
            {
                ok = true,
                action = "unlock_workbook",
                message = "工作簿锁定已解除"
            };
        }

        private static BridgeResponse HandleListSnapshots(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic wb = FindTargetWorkbook(app, targetFullName, targetName);

            if (wb == null || string.IsNullOrEmpty((string)wb.Path))
            {
                return new BridgeResponse
                {
                    ok = true,
                    action = "list_snapshots",
                    data = new List<SnapshotItem>()
                };
            }

            var list = SnapshotManager.LoadSnapshots((string)wb.FullName);
            return new BridgeResponse
            {
                ok = true,
                action = "list_snapshots",
                data = list
            };
        }

        private static BridgeResponse HandleListScripts()
        {
            var list = ScriptManager.ListScripts();
            return new BridgeResponse
            {
                ok = true,
                action = "list_scripts",
                data = list
            };
        }

        private static BridgeResponse HandleSaveScript(Dictionary<string, string> req)
        {
            string id = req.ContainsKey("id") ? req["id"] : "";
            string displayName = req.ContainsKey("displayName") ? req["displayName"] : (req.ContainsKey("name") ? req["name"] : "");
            string code = req.ContainsKey("code") ? req["code"] : "";
            string desc = req.ContainsKey("description") ? req["description"] : "";
            string category = req.ContainsKey("category") ? req["category"] : "";
            string sourceType = req.ContainsKey("sourceType") ? req["sourceType"] : "paste";
            string originalFileName = req.ContainsKey("originalFileName") ? req["originalFileName"] : "";
            string encoding = req.ContainsKey("encoding") ? req["encoding"] : "UTF-8";
            string entryPoint = req.ContainsKey("entryPoint") ? req["entryPoint"] : "";
            string rawBytesBase64 = req.ContainsKey("rawBytesBase64") ? req["rawBytesBase64"] : "";
            bool overwrite = req.ContainsKey("overwrite") && (req["overwrite"] == "true" || req["overwrite"] == "1" || req["overwrite"] == "True");

            string savedId;
            string err;
            bool ok = ScriptManager.SaveScript(
                id,
                displayName,
                code,
                desc,
                category,
                sourceType,
                originalFileName,
                encoding,
                entryPoint,
                rawBytesBase64,
                overwrite,
                out savedId,
                out err
            );

            return new BridgeResponse
            {
                ok = ok,
                action = "save_script",
                message = ok ? ("宏【" + displayName + "】已成功保存！") : err,
                error = err,
                data = new Dictionary<string, string> { { "id", savedId ?? "" }, { "displayName", displayName } }
            };
        }

        private static BridgeResponse HandleRenameScript(Dictionary<string, string> req)
        {
            string id = req.ContainsKey("id") ? req["id"] : (req.ContainsKey("fileName") ? req["fileName"] : "");
            string newDisplayName = req.ContainsKey("newDisplayName") ? req["newDisplayName"] : (req.ContainsKey("displayName") ? req["displayName"] : "");

            string err;
            bool ok = ScriptManager.RenameScript(id, newDisplayName, out err);
            return new BridgeResponse
            {
                ok = ok,
                action = "rename_script",
                message = ok ? ("宏显示名称已成功更新为【" + newDisplayName + "】！") : err,
                error = err,
                data = new Dictionary<string, string> { { "id", id }, { "displayName", newDisplayName } }
            };
        }

        private static BridgeResponse HandleUpdateScriptResult(Dictionary<string, string> req)
        {
            string id = req.ContainsKey("id") ? req["id"] : (req.ContainsKey("fileName") ? req["fileName"] : "");
            string result = req.ContainsKey("lastExecutionResult") ? req["lastExecutionResult"] : "";

            string err;
            bool ok = ScriptManager.UpdateExecutionResult(id, result, out err);
            return new BridgeResponse
            {
                ok = ok,
                action = "update_script_result",
                message = ok ? "执行状态已同步更新" : err,
                error = err
            };
        }

        private static BridgeResponse HandleDeleteScript(Dictionary<string, string> req)
        {
            string fileName = req.ContainsKey("id") ? req["id"] : (req.ContainsKey("fileName") ? req["fileName"] : "");
            string err;
            bool ok = ScriptManager.DeleteScript(fileName, out err);
            return new BridgeResponse
            {
                ok = ok,
                action = "delete_script",
                message = ok ? "宏已成功删除" : err,
                error = err
            };
        }

        private static BridgeResponse HandleGetSelectionContext(Dictionary<string, string> req, dynamic app)
        {
            int sampleRows = SelectionContextService.DefaultMaxSampleRows;
            int sampleCols = SelectionContextService.DefaultMaxSampleCols;

            if (req.ContainsKey("sampleRows"))
            {
                int r;
                if (int.TryParse(req["sampleRows"], out r)) sampleRows = r;
            }
            if (req.ContainsKey("sampleCols"))
            {
                int c;
                if (int.TryParse(req["sampleCols"], out c)) sampleCols = c;
            }

            var result = SelectionContextService.GetSelectionContext(app, sampleRows, sampleCols);
            if (!result.ok)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_selection_context",
                    error = result.error,
                    data = new Dictionary<string, string> { { "errorType", result.errorType ?? "unknown" } }
                };
            }

            return new BridgeResponse
            {
                ok = true,
                action = "get_selection_context",
                data = result.data
            };
        }
    }
}
