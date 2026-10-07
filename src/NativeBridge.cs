using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

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
        public string metaSaveError { get; set; }
    }

    public class BridgeResponse
    {
        public bool ok { get; set; }
        public string action { get; set; }
        public string requestId { get; set; }
        public string message { get; set; }
        public string metaSaveError { get; set; }
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
                    case "execute_vba_code":
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

                    case "update_script_tags":
                        resp = HandleUpdateScriptTags(req);
                        break;

                    case "update_script_favorite":
                        resp = HandleUpdateScriptFavorite(req);
                        break;

                    case "get_target_sheets":
                        resp = HandleGetTargetSheets(req, app);
                        break;

                    case "inspect_macro_signature":
                        resp = HandleInspectMacroSignature(req);
                        break;

                    case "analyze_dedup":
                        resp = HandleAnalyzeDedup(req, app);
                        break;

                    case "apply_dedup":
                        resp = HandleApplyDedup(req, app);
                        break;

                    case "analyze_reconcile":
                        resp = HandleAnalyzeReconcile(req, app);
                        break;

                    case "apply_reconcile":
                        resp = HandleApplyReconcile(req, app);
                        break;

                    case "analyze_consolidation":
                        resp = HandleAnalyzeConsolidation(req, app);
                        break;

                    case "apply_consolidation":
                        resp = HandleApplyConsolidation(req, app);
                        break;

                    case "unlock_workbook":
                        resp = HandleUnlockWorkbook(req, app);
                        break;

                    case "validate_batch_job":
                        resp = HandleValidateBatchJob(req);
                        break;

                    case "start_batch_job":
                        resp = HandleStartBatchJob(req, app);
                        break;

                    case "cancel_batch_job":
                        resp = HandleCancelBatchJob(req);
                        break;

                    case "get_batch_job_status":
                        resp = HandleGetBatchJobStatus(req);
                        break;

                    case "open_output_folder":
                        resp = HandleOpenOutputFolder(req);
                        break;

                    case "browse_files":
                        resp = HandleBrowseFiles(req);
                        break;

                    case "browse_folder":
                        resp = HandleBrowseFolder(req);
                        break;

                    case "browse_save_file":
                        resp = HandleBrowseSaveFile(req);
                        break;

                    case "list_workflows":
                        resp = HandleListWorkflows();
                        break;

                    case "get_workflow":
                        resp = HandleGetWorkflow(req);
                        break;

                    case "save_workflow":
                        resp = HandleSaveWorkflow(jsonString);
                        break;

                    case "delete_workflow":
                        resp = HandleDeleteWorkflow(req);
                        break;

                    case "execute_workflow":
                        resp = HandleExecuteWorkflow(jsonString, app);
                        break;

                    case "cancel_workflow":
                        resp = HandleCancelWorkflow();
                        break;

                    case "list_managed_charts":
                        resp = HandleListManagedCharts(req, app);
                        break;

                    case "execute_quick_chart":
                        resp = HandleExecuteQuickChart(jsonString, app);
                        break;

                    case "preview_external_data":
                        resp = HandlePreviewExternalData(jsonString);
                        break;

                    case "import_external_data":
                        resp = HandleImportExternalData(jsonString, app);
                        break;

                    case "list_data_source_configs":
                        resp = HandleListDataSourceConfigs();
                        break;

                    case "save_data_source_config":
                        resp = HandleSaveDataSourceConfig(jsonString);
                        break;

                    case "export_macro_package":
                        resp = HandleExportMacroPackage(jsonString);
                        break;

                    case "preview_macro_package":
                        resp = HandlePreviewMacroPackage(jsonString);
                        break;

                    case "import_macro_package":
                        resp = HandleImportMacroPackage(jsonString);
                        break;

                    case "preview_diagnostics":
                        resp = HandlePreviewDiagnostics(jsonString, app);
                        break;

                    case "export_diagnostics":
                        resp = HandleExportDiagnostics(jsonString, app);
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

            // 若调用方明确指定了目标工作簿但未能匹配到已打开的工作簿，拒绝保底回退，防止误在活动工作簿执行
            if (!string.IsNullOrEmpty(targetFullName) || !string.IsNullOrEmpty(targetName))
            {
                return null;
            }

            // 3. 保底：仅当未指定特定工作簿时，自动采用当前活动工作簿
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

            string code = req.ContainsKey("code") ? req["code"] : "";
            string prompt = req.ContainsKey("prompt") ? req["prompt"] : "自然语言操作";

            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);

            if (targetWb == null)
            {
                string targetDesc = !string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName;
                string blockErr = "未找到目标工作簿 (" + (targetDesc ?? "未指定") + ")，请确认该工作簿已在 Excel 中打开。";
                string dummyMetaErr;
                RecordScriptExecutionHistory(req, code, targetName, null, "blocked", "before_run", 0, blockErr, null, out dummyMetaErr);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = blockErr
                };
            }

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
            string entryPoint = req.ContainsKey("entryPoint") ? req["entryPoint"] : null;

            // 解析参数列表 (支持 JSON 数组 [ {name, value} ] 或 JSON 对象字典 { key: val })
            List<VbaParameterInput> parsedParams = null;
            if (req.ContainsKey("parameters") && !string.IsNullOrEmpty(req["parameters"]))
            {
                string pRaw = req["parameters"].Trim();
                try
                {
                    if (pRaw.StartsWith("["))
                    {
                        parsedParams = SimpleJson.DeserializeList<VbaParameterInput>(pRaw);
                    }
                    else if (pRaw.StartsWith("{"))
                    {
                        var dict = SimpleJson.ParseFlatObject(pRaw);
                        if (dict != null)
                        {
                            parsedParams = new List<VbaParameterInput>();
                            foreach (var kvp in dict)
                            {
                                parsedParams.Add(new VbaParameterInput
                                {
                                    name = kvp.Key,
                                    value = kvp.Value
                                });
                            }
                        }
                    }
                }
                catch { }
            }

            // 0. 执行前参数与入口严格预检 (必须在快照创建之前执行！确保参数缺失/非法时绝对零快照、零执行)
            var precheck = VbaRunner.PrecheckParameters(app, targetWb, code, entryPoint, parsedParams);
            if (!precheck.isOk)
            {
                string blockErr = "执行已拒绝：" + precheck.error;
                string dummyMetaErr;
                RecordScriptExecutionHistory(req, code, targetName, null, "blocked", precheck.failureStage ?? "before_run", 0, blockErr, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out dummyMetaErr);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = blockErr
                };
            }

            // 1. 运行前自动执行整本物理副本快照 (针对目标工作簿 targetWb)
            SnapshotItem snap = null;
            try
            {
                snap = SnapshotManager.CreateSnapshot(targetWb, prompt, code);
            }
            catch (Exception snapEx)
            {
                System.Diagnostics.Debug.WriteLine("快照创建失败: " + snapEx.Message);
                string blockErr = "执行已安全中止：执行前目标工作簿快照备份失败 (" + snapEx.Message + ")。为保障数据可回滚安全，拒绝执行代码。";
                string dummyMetaErr;
                RecordScriptExecutionHistory(req, code, targetName, null, "blocked", "before_run", 0, blockErr, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out dummyMetaErr);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = blockErr
                };
            }

            if (snap == null || string.IsNullOrEmpty(snap.id) || string.IsNullOrEmpty(snap.fileName))
            {
                string blockErr = "执行已安全中止：目标工作簿物理快照备份未能成功生成有效副本。请确认工作簿已保存并具有有效磁盘路径。为保障数据可回滚安全，拒绝执行代码。";
                string dummyMetaErr;
                RecordScriptExecutionHistory(req, code, targetName, null, "blocked", "before_run", 0, blockErr, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out dummyMetaErr);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = blockErr
                };
            }

            try
            {
                string wbPathForFolder = !string.IsNullOrEmpty(snap.originalPath) ? snap.originalPath : (!string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName);
                string snapFolder = SnapshotManager.GetBackupFolderForWorkbook(wbPathForFolder);
                string snapFile = Path.Combine(snapFolder, snap.fileName);
                if (!File.Exists(snapFile) || new FileInfo(snapFile).Length == 0)
                {
                    string blockErr = "执行已安全中止：快照物理文件未能有效写入磁盘（文件不存在或为0字节）。拒绝执行代码以防数据丢失。";
                    string dummyMetaErr;
                    RecordScriptExecutionHistory(req, code, targetName, null, "blocked", "before_run", 0, blockErr, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out dummyMetaErr);
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "execute_vba",
                        error = blockErr
                    };
                }
            }
            catch (Exception checkEx)
            {
                string blockErr = "执行已安全中止：快照文件物理校验异常 (" + checkEx.Message + ")。拒绝执行代码以防数据丢失。";
                string dummyMetaErr;
                RecordScriptExecutionHistory(req, code, targetName, null, "blocked", "before_run", 0, blockErr, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out dummyMetaErr);
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_vba",
                    error = blockErr
                };
            }

            // 2. 在目标工作簿中执行动态 VBA 并读回实际状态
            var result = VbaRunner.RunVbaCode(app, targetWb, code, rawModelResponse, entryPoint, parsedParams);

            // 安全获取工作簿名称与路径：宏可能关闭了工作簿导致 COM 引用失效
            string resolvedTargetName = targetName;
            string resolvedTargetFullName = targetFullName;
            try
            {
                resolvedTargetName = (string)targetWb.Name;
                resolvedTargetFullName = (string)targetWb.FullName;
            }
            catch { }

            // 无论成功还是失败，均写入结构化运行记录与关联快照（区分已执行成功、已开始执行但失败、执行前被阻断）
            string execStatus = result.success ? "success" : "failed";
            string execPhase = result.success ? "execution" : (result.failureStage ?? "runtime");
            string execSummary = result.success 
                ? (result.summary ?? ("执行完成，区域 " + (result.readback != null ? result.readback.usedRangeAddress : "") + " 验证通过"))
                : (result.error ?? result.summary ?? "宏执行失败");

            string metaSaveWarning = null;
            RecordScriptExecutionHistory(req, code, resolvedTargetName, snap, execStatus, execPhase, (int)result.elapsedMs, execSummary, entryPoint, precheck.paramTypes, precheck.sanitizedSummary, out metaSaveWarning);

            string finalMsg = result.summary;
            if (!string.IsNullOrEmpty(metaSaveWarning))
            {
                finalMsg = (finalMsg != null ? finalMsg + "\n" : "") + "⚠️ 运行记录保存失败: " + metaSaveWarning;
            }

            return new BridgeResponse
            {
                ok = result.success,
                action = "execute_vba",
                message = finalMsg,
                metaSaveError = metaSaveWarning,
                error = result.error,
                data = new ExecutionDataDto
                {
                    summary = finalMsg,
                    error = result.error,
                    elapsedMs = result.elapsedMs,
                    metaSaveError = metaSaveWarning,
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

            string attachmentId = req.ContainsKey("attachmentId") ? req["attachmentId"] : null;
            string targetSheetName = req.ContainsKey("targetSheetName") ? req["targetSheetName"] : null;
            string targetAddress = req.ContainsKey("targetAddress") ? req["targetAddress"] : null;
            string targetWbFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : null;
            string targetWbName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : null;

            SelectionContextResult result;
            if (!string.IsNullOrEmpty(targetSheetName) && !string.IsNullOrEmpty(targetAddress))
            {
                // 刷新原区域：严格定向读取原工作簿、原工作表和原地址，不激活、不改选区
                result = SelectionContextService.GetSpecificRangeContext(
                    app, targetWbFullName, targetWbName, targetSheetName, targetAddress, sampleRows, sampleCols, attachmentId);
            }
            else
            {
                // 附加当前选区：读取当前活动选区
                result = SelectionContextService.GetSelectionContext(app, sampleRows, sampleCols, attachmentId);
            }

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

        private static BridgeResponse HandleUpdateScriptTags(Dictionary<string, string> req)
        {
            string id = req.ContainsKey("id") ? req["id"] : (req.ContainsKey("fileName") ? req["fileName"] : "");
            string tagsJson = req.ContainsKey("tags") ? req["tags"] : "[]";
            var tags = SimpleJson.ParseStringList(tagsJson);

            string err;
            bool ok = ScriptManager.UpdateScriptTags(id, tags, out err);
            return new BridgeResponse
            {
                ok = ok,
                action = "update_script_tags",
                message = ok ? "标签已成功更新" : err,
                error = err,
                data = new Dictionary<string, string> { { "id", id }, { "tags", tagsJson } }
            };
        }

        private static BridgeResponse HandleUpdateScriptFavorite(Dictionary<string, string> req)
        {
            string id = req.ContainsKey("id") ? req["id"] : (req.ContainsKey("fileName") ? req["fileName"] : "");
            string isFavStr = req.ContainsKey("isFavorite") ? req["isFavorite"] : "false";
            bool isFav = (isFavStr == "true" || isFavStr == "1" || isFavStr == "True");

            string err;
            bool ok = ScriptManager.UpdateScriptFavorite(id, isFav, out err);
            return new BridgeResponse
            {
                ok = ok,
                action = "update_script_favorite",
                message = ok ? (isFav ? "已成功添加到收藏宏" : "已取消收藏") : err,
                error = err,
                data = new Dictionary<string, string> { { "id", id }, { "isFavorite", isFav ? "true" : "false" } }
            };
        }

        private static bool RecordScriptExecutionHistory(
            Dictionary<string, string> req,
            string code,
            string targetWbName,
            SnapshotItem snap,
            string status,
            string phase,
            int elapsedMs,
            string summary,
            string entryPoint,
            out string metaSaveError)
        {
            return RecordScriptExecutionHistory(req, code, targetWbName, snap, status, phase, elapsedMs, summary, entryPoint, null, null, out metaSaveError);
        }

        private static bool RecordScriptExecutionHistory(
            Dictionary<string, string> req,
            string code,
            string targetWbName,
            SnapshotItem snap,
            string status,
            string phase,
            int elapsedMs,
            string summary,
            string entryPoint,
            List<string> parameterTypes,
            Dictionary<string, string> parameterSummary,
            out string metaSaveError)
        {
            metaSaveError = null;
            string scriptId = req.ContainsKey("scriptId") ? req["scriptId"] : (req.ContainsKey("fileName") ? req["fileName"] : null);
            if (string.IsNullOrEmpty(scriptId)) return true;

            try
            {
                var record = new RunRecordDto
                {
                    id = "run_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"),
                    executedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    status = status,
                    phase = phase ?? "execution",
                    targetWorkbookName = !string.IsNullOrEmpty(targetWbName) ? targetWbName : "未知工作簿",
                    codeHash = VbaRunner.ComputeSha256(code ?? ""),
                    snapshotId = snap != null ? (snap.id ?? "") : "",
                    snapshotExists = snap != null,
                    snapshotReason = (status == "blocked") ? ("执行前阻断，未生成快照 (" + (summary ?? "") + ")") : (snap == null ? "未生成物理快照" : ""),
                    elapsedMs = (status == "blocked") ? 0 : elapsedMs,
                    summary = summary ?? "",
                    entryPoint = entryPoint ?? "",
                    parameterTypes = parameterTypes,
                    parameterSummary = parameterSummary
                };

                string err;
                bool ok = ScriptManager.AppendRunRecord(scriptId, record, out err);
                if (!ok)
                {
                    metaSaveError = err;
                    System.Diagnostics.Debug.WriteLine("[ScriptManager] 运行记录保存失败: " + err);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                // 元数据写入失败绝对不得触发宏重新执行，也不得把已完成的宏误报为执行失败！
                metaSaveError = ex.Message;
                System.Diagnostics.Debug.WriteLine("[ScriptManager] 运行记录异常拦截: " + ex.Message);
                return false;
            }
        }

        private static BridgeResponse HandleGetTargetSheets(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);
            if (targetWb == null)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_target_sheets",
                    error = "未找到目标工作簿 (" + (!string.IsNullOrEmpty(targetFullName) ? targetFullName : targetName) + ")，请确认该工作簿已在 Excel 中打开。"
                };
            }

            var sheetList = new List<Dictionary<string, object>>();
            try
            {
                foreach (dynamic sh in targetWb.Worksheets)
                {
                    try
                    {
                        var sInfo = new Dictionary<string, object>();
                        sInfo["name"] = (string)sh.Name;
                        sInfo["visible"] = (int)sh.Visible == -1;
                        sInfo["index"] = (int)sh.Index;
                        sheetList.Add(sInfo);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_target_sheets",
                    error = "读取工作表列表失败: " + ex.Message
                };
            }

            return new BridgeResponse
            {
                ok = true,
                action = "get_target_sheets",
                data = sheetList
            };
        }

        private static BridgeResponse HandleInspectMacroSignature(Dictionary<string, string> req)
        {
            string code = req.ContainsKey("code") ? req["code"] : (req.ContainsKey("vbaCode") ? req["vbaCode"] : "");
            if (string.IsNullOrEmpty(code) && req.ContainsKey("scriptId") && !string.IsNullOrEmpty(req["scriptId"]))
            {
                var script = ScriptManager.GetScriptById(req["scriptId"]);
                if (script != null && !string.IsNullOrEmpty(script.code))
                {
                    code = script.code;
                }
            }
            if (string.IsNullOrEmpty(code))
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "inspect_macro_signature",
                    error = "VBA 源码为空"
                };
            }

            var procs = VbaSignatureParser.ParseSignatures(code);
            return new BridgeResponse
            {
                ok = true,
                action = "inspect_macro_signature",
                data = new Dictionary<string, object>
                {
                    { "entryPoints", procs }
                }
            };
        }

        private static BridgeResponse HandleAnalyzeDedup(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);

            if (targetWb == null)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "analyze_dedup",
                    error = "未找到指定的目标工作簿或工作簿已关闭。"
                };
            }

            string sheetName = req.ContainsKey("sheetName") ? req["sheetName"] : "";
            string rangeAddress = req.ContainsKey("rangeAddress") ? req["rangeAddress"] : "";
            bool hasHeader = req.ContainsKey("hasHeader") && (req["hasHeader"] == "true" || req["hasHeader"] == "1" || req["hasHeader"] == "True");

            List<int> keyCols = new List<int>();
            if (req.ContainsKey("keyColumnIndices") && !string.IsNullOrEmpty(req["keyColumnIndices"]))
            {
                string rawCols = req["keyColumnIndices"].Trim();
                try
                {
                    if (rawCols.StartsWith("["))
                    {
                        var listStr = SimpleJson.ParseStringList(rawCols);
                        if (listStr != null && listStr.Count > 0)
                        {
                            foreach (var s in listStr)
                            {
                                int idx;
                                if (int.TryParse(s, out idx)) keyCols.Add(idx);
                            }
                        }
                        else
                        {
                            // 兼容 [1, 2] 纯数字数组
                            string stripped = rawCols.Trim('[', ']');
                            var parts = stripped.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var p in parts)
                            {
                                int idx;
                                if (int.TryParse(p.Trim(), out idx)) keyCols.Add(idx);
                            }
                        }
                    }
                    else
                    {
                        var parts = rawCols.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var p in parts)
                        {
                            int idx;
                            if (int.TryParse(p.Trim(), out idx)) keyCols.Add(idx);
                        }
                    }
                }
                catch { }
            }

            var analysis = DataToolsService.AnalyzeDedup(targetWb, sheetName, rangeAddress, hasHeader, keyCols);
            return new BridgeResponse
            {
                ok = analysis.ok,
                action = "analyze_dedup",
                error = analysis.error,
                data = analysis
            };
        }

        private static BridgeResponse HandleApplyDedup(Dictionary<string, string> req, dynamic app)
        {
            string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
            string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
            dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);

            if (targetWb == null)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "apply_dedup",
                    error = "未找到指定的目标工作簿或工作簿已关闭。"
                };
            }

            string sheetName = req.ContainsKey("sheetName") ? req["sheetName"] : "";
            string rangeAddress = req.ContainsKey("rangeAddress") ? req["rangeAddress"] : "";
            bool hasHeader = req.ContainsKey("hasHeader") && (req["hasHeader"] == "true" || req["hasHeader"] == "1" || req["hasHeader"] == "True");
            string mode = req.ContainsKey("mode") ? req["mode"] : "highlight";

            List<int> keyCols = new List<int>();
            if (req.ContainsKey("keyColumnIndices") && !string.IsNullOrEmpty(req["keyColumnIndices"]))
            {
                string rawCols = req["keyColumnIndices"].Trim();
                try
                {
                    if (rawCols.StartsWith("["))
                    {
                        var listStr = SimpleJson.ParseStringList(rawCols);
                        if (listStr != null && listStr.Count > 0)
                        {
                            foreach (var s in listStr)
                            {
                                int idx;
                                if (int.TryParse(s, out idx)) keyCols.Add(idx);
                            }
                        }
                        else
                        {
                            string stripped = rawCols.Trim('[', ']');
                            var parts = stripped.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var p in parts)
                            {
                                int idx;
                                if (int.TryParse(p.Trim(), out idx)) keyCols.Add(idx);
                            }
                        }
                    }
                    else
                    {
                        var parts = rawCols.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var p in parts)
                        {
                            int idx;
                            if (int.TryParse(p.Trim(), out idx)) keyCols.Add(idx);
                        }
                    }
                }
                catch { }
            }

            string expectedFingerprint = req.ContainsKey("expectedFingerprint") ? req["expectedFingerprint"] : null;
            var execRes = DataToolsService.ExecuteDedup(app, targetWb, sheetName, rangeAddress, hasHeader, keyCols, mode, expectedFingerprint);
            return new BridgeResponse
            {
                ok = execRes.ok,
                action = "apply_dedup",
                message = execRes.message,
                error = execRes.error,
                data = execRes
            };
        }

        private static List<int> ParseIntList(string rawCols)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(rawCols)) return result;

            rawCols = rawCols.Trim();
            try
            {
                if (rawCols.StartsWith("["))
                {
                    var listStr = SimpleJson.ParseStringList(rawCols);
                    if (listStr != null && listStr.Count > 0)
                    {
                        foreach (var s in listStr)
                        {
                            int idx;
                            if (int.TryParse(s, out idx)) result.Add(idx);
                        }
                    }
                    else
                    {
                        string stripped = rawCols.Trim('[', ']');
                        var parts = stripped.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var p in parts)
                        {
                            int idx;
                            if (int.TryParse(p.Trim(), out idx)) result.Add(idx);
                        }
                    }
                }
                else
                {
                    var parts = rawCols.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var p in parts)
                    {
                        int idx;
                        if (int.TryParse(p.Trim(), out idx)) result.Add(idx);
                    }
                }
            }
            catch { }
            return result;
        }

        private static BridgeResponse HandleAnalyzeReconcile(Dictionary<string, string> req, dynamic app)
        {
            string leftWbName = req.ContainsKey("leftWorkbookName") ? req["leftWorkbookName"] : "";
            string leftSheetName = req.ContainsKey("leftSheetName") ? req["leftSheetName"] : "";
            string leftRangeAddress = req.ContainsKey("leftRangeAddress") ? req["leftRangeAddress"] : "";
            bool leftHasHeader = req.ContainsKey("leftHasHeader") && (req["leftHasHeader"] == "true" || req["leftHasHeader"] == "1" || req["leftHasHeader"] == "True");

            string rightWbName = req.ContainsKey("rightWorkbookName") ? req["rightWorkbookName"] : "";
            string rightSheetName = req.ContainsKey("rightSheetName") ? req["rightSheetName"] : "";
            string rightRangeAddress = req.ContainsKey("rightRangeAddress") ? req["rightRangeAddress"] : "";
            bool rightHasHeader = req.ContainsKey("rightHasHeader") && (req["rightHasHeader"] == "true" || req["rightHasHeader"] == "1" || req["rightHasHeader"] == "True");

            List<int> leftKeyCols = ParseIntList(req.ContainsKey("leftKeyCols") ? req["leftKeyCols"] : "");
            List<int> rightKeyCols = ParseIntList(req.ContainsKey("rightKeyCols") ? req["rightKeyCols"] : "");

            List<CompareColMappingDto> compareCols = new List<CompareColMappingDto>();
            if (req.ContainsKey("compareCols") && !string.IsNullOrWhiteSpace(req["compareCols"]))
            {
                try
                {
                    compareCols = SimpleJson.DeserializeList<CompareColMappingDto>(req["compareCols"]);
                }
                catch { }
            }

            var analysis = DataToolsService.AnalyzeReconcile(
                app,
                leftWbName, leftSheetName, leftRangeAddress, leftHasHeader, leftKeyCols,
                rightWbName, rightSheetName, rightRangeAddress, rightHasHeader, rightKeyCols,
                compareCols);

            return new BridgeResponse
            {
                ok = analysis.ok,
                action = "analyze_reconcile",
                error = analysis.error,
                data = analysis
            };
        }

        private static BridgeResponse HandleApplyReconcile(Dictionary<string, string> req, dynamic app)
        {
            string leftWbName = req.ContainsKey("leftWorkbookName") ? req["leftWorkbookName"] : "";
            string leftSheetName = req.ContainsKey("leftSheetName") ? req["leftSheetName"] : "";
            string leftRangeAddress = req.ContainsKey("leftRangeAddress") ? req["leftRangeAddress"] : "";
            bool leftHasHeader = req.ContainsKey("leftHasHeader") && (req["leftHasHeader"] == "true" || req["leftHasHeader"] == "1" || req["leftHasHeader"] == "True");

            string rightWbName = req.ContainsKey("rightWorkbookName") ? req["rightWorkbookName"] : "";
            string rightSheetName = req.ContainsKey("rightSheetName") ? req["rightSheetName"] : "";
            string rightRangeAddress = req.ContainsKey("rightRangeAddress") ? req["rightRangeAddress"] : "";
            bool rightHasHeader = req.ContainsKey("rightHasHeader") && (req["rightHasHeader"] == "true" || req["rightHasHeader"] == "1" || req["rightHasHeader"] == "True");

            List<int> leftKeyCols = ParseIntList(req.ContainsKey("leftKeyCols") ? req["leftKeyCols"] : "");
            List<int> rightKeyCols = ParseIntList(req.ContainsKey("rightKeyCols") ? req["rightKeyCols"] : "");

            List<CompareColMappingDto> compareCols = new List<CompareColMappingDto>();
            if (req.ContainsKey("compareCols") && !string.IsNullOrWhiteSpace(req["compareCols"]))
            {
                try
                {
                    compareCols = SimpleJson.DeserializeList<CompareColMappingDto>(req["compareCols"]);
                }
                catch { }
            }

            string outputWbName = req.ContainsKey("outputWorkbookName") ? req["outputWorkbookName"] : "";
            string expectedFingerprint = req.ContainsKey("expectedFingerprint") ? req["expectedFingerprint"] : null;

            var execRes = DataToolsService.ExecuteReconcile(
                app,
                leftWbName, leftSheetName, leftRangeAddress, leftHasHeader, leftKeyCols,
                rightWbName, rightSheetName, rightRangeAddress, rightHasHeader, rightKeyCols,
                compareCols, outputWbName, expectedFingerprint);

            return new BridgeResponse
            {
                ok = execRes.ok,
                action = "apply_reconcile",
                message = execRes.message,
                error = execRes.error,
                data = execRes
            };
        }

        private static BridgeResponse HandleAnalyzeConsolidation(Dictionary<string, string> req, dynamic app)
        {
            List<ConsolidationSourceDef> sourceDefs = new List<ConsolidationSourceDef>();
            if (req.ContainsKey("sources") && !string.IsNullOrWhiteSpace(req["sources"]))
            {
                try
                {
                    sourceDefs = SimpleJson.DeserializeList<ConsolidationSourceDef>(req["sources"]);
                }
                catch { }
            }

            List<string> userColumnOrder = new List<string>();
            if (req.ContainsKey("outputColumnOrder") && !string.IsNullOrWhiteSpace(req["outputColumnOrder"]))
            {
                userColumnOrder = ParseStringList(req["outputColumnOrder"]);
            }

            List<ConsolidationColMapping> explicitMappings = new List<ConsolidationColMapping>();
            if (req.ContainsKey("columnMappings") && !string.IsNullOrWhiteSpace(req["columnMappings"]))
            {
                try
                {
                    explicitMappings = SimpleJson.DeserializeList<ConsolidationColMapping>(req["columnMappings"]);
                }
                catch { }
            }

            var analysis = DataToolsService.AnalyzeConsolidation(app, sourceDefs, userColumnOrder, explicitMappings);

            return new BridgeResponse
            {
                ok = analysis.ok,
                action = "analyze_consolidation",
                error = analysis.error,
                data = analysis
            };
        }

        private static BridgeResponse HandleApplyConsolidation(Dictionary<string, string> req, dynamic app)
        {
            List<ConsolidationSourceDef> sourceDefs = new List<ConsolidationSourceDef>();
            if (req.ContainsKey("sources") && !string.IsNullOrWhiteSpace(req["sources"]))
            {
                try
                {
                    sourceDefs = SimpleJson.DeserializeList<ConsolidationSourceDef>(req["sources"]);
                }
                catch { }
            }

            string outputFilePath = req.ContainsKey("outputFilePath") ? req["outputFilePath"] : "";

            List<string> userColumnOrder = new List<string>();
            if (req.ContainsKey("outputColumnOrder") && !string.IsNullOrWhiteSpace(req["outputColumnOrder"]))
            {
                userColumnOrder = ParseStringList(req["outputColumnOrder"]);
            }

            List<ConsolidationColMapping> explicitMappings = new List<ConsolidationColMapping>();
            if (req.ContainsKey("columnMappings") && !string.IsNullOrWhiteSpace(req["columnMappings"]))
            {
                try
                {
                    explicitMappings = SimpleJson.DeserializeList<ConsolidationColMapping>(req["columnMappings"]);
                }
                catch { }
            }

            string expectedFingerprint = req.ContainsKey("expectedFingerprint") ? req["expectedFingerprint"] : null;
            bool includeMetadataCols = !req.ContainsKey("includeMetadataCols") || (req["includeMetadataCols"] == "true" || req["includeMetadataCols"] == "1" || req["includeMetadataCols"] == "True");

            var execRes = DataToolsService.ExecuteConsolidation(
                app,
                sourceDefs,
                outputFilePath,
                userColumnOrder,
                explicitMappings,
                expectedFingerprint,
                includeMetadataCols);

            return new BridgeResponse
            {
                ok = execRes.ok,
                action = "apply_consolidation",
                error = execRes.error,
                data = execRes
            };
        }

        private static List<string> ParseStringList(string raw)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(raw)) return result;
            string trimmed = raw.Trim().TrimEnd(',');
            try
            {
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    int i = 1;
                    int len = trimmed.Length - 1;
                    while (i < len)
                    {
                        while (i < len && (char.IsWhiteSpace(trimmed[i]) || trimmed[i] == ',')) i++;
                        if (i >= len) break;
                        if (trimmed[i] == '"' || trimmed[i] == '\'')
                        {
                            char quote = trimmed[i];
                            i++;
                            var sb = new StringBuilder();
                            while (i < len)
                            {
                                if (trimmed[i] == '\\' && i + 1 < len)
                                {
                                    char next = trimmed[i + 1];
                                    if (next == '\\') { sb.Append('\\'); i += 2; }
                                    else if (next == '"') { sb.Append('"'); i += 2; }
                                    else if (next == '\'') { sb.Append('\''); i += 2; }
                                    else if (next == '/') { sb.Append('/'); i += 2; }
                                    else if (next == 'u' && i + 5 < len)
                                    {
                                        string hex = trimmed.Substring(i + 2, 4);
                                        int code;
                                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out code))
                                        {
                                            sb.Append((char)code);
                                            i += 6;
                                        }
                                        else
                                        {
                                            sb.Append('\\');
                                            sb.Append(next);
                                            i += 2;
                                        }
                                    }
                                    else
                                    {
                                        // Windows 文件路径中，\Users, \batch_test, \test 等字面反斜杠必须完整保留，绝不能吞掉或误转为 ASCII 控制符
                                        sb.Append('\\');
                                        sb.Append(next);
                                        i += 2;
                                    }
                                }
                                else if (trimmed[i] == quote)
                                {
                                    i++;
                                    break;
                                }
                                else
                                {
                                    sb.Append(trimmed[i]);
                                    i++;
                                }
                            }
                            string val = sb.ToString().Trim();
                            if (!string.IsNullOrEmpty(val)) result.Add(val);
                        }
                        else
                        {
                            i++;
                        }
                    }
                    if (result.Count > 0) return result;
                }
            }
            catch { }

            try
            {
                string stripped = trimmed.Trim('[', ']', '\"', '\'');
                var lines = stripped.Split(new[] { '\r', '\n', ';', '|', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    string clean = line.Trim().Trim('\"', '\'', '[', ']');
                    clean = clean.Replace("\\\\", "\\").Replace("\\\"", "\"");
                    if (!string.IsNullOrEmpty(clean)) result.Add(clean);
                }
            }
            catch { }
            return result;
        }

        private static BridgeResponse HandleValidateBatchJob(Dictionary<string, string> req)
        {
            try
            {
                string rawFiles = req.ContainsKey("filePaths") ? req["filePaths"] : "";
                var filePaths = ParseStringList(rawFiles);
                string macroCode = req.ContainsKey("code") ? req["code"] : (req.ContainsKey("macroCode") ? req["macroCode"] : "");
                string entryPoint = req.ContainsKey("entryPoint") ? req["entryPoint"] : null;
                string parameters = req.ContainsKey("parameters") ? req["parameters"] : null;
                string outputDir = req.ContainsKey("outputDir") ? req["outputDir"] : "";
                bool stopOnError = !req.ContainsKey("stopOnError") || (req["stopOnError"] == "true" || req["stopOnError"] == "1" || req["stopOnError"] == "True");

                string valErr;
                var jobDef = BatchRunnerService.ValidateAndLockJob(filePaths, macroCode, entryPoint, parameters, outputDir, stopOnError, out valErr);
                if (jobDef == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "validate_batch_job",
                        error = valErr
                    };
                }

                return new BridgeResponse
                {
                    ok = true,
                    action = "validate_batch_job",
                    data = jobDef,
                    message = "批量任务队列已成功校验并固化。"
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "validate_batch_job",
                    error = "校验批量任务异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleStartBatchJob(Dictionary<string, string> req, dynamic app)
        {
            try
            {
                string jobId = req.ContainsKey("jobId") ? req["jobId"].Trim() : "";
                BatchJobDefinition jobDef = null;

                if (!string.IsNullOrEmpty(jobId))
                {
                    jobDef = BatchRunnerService.GetLockedJob(jobId);
                }

                string rawFiles = req.ContainsKey("filePaths") ? req["filePaths"] : "";
                var filePaths = ParseStringList(rawFiles);
                string macroCode = req.ContainsKey("code") ? req["code"] : (req.ContainsKey("macroCode") ? req["macroCode"] : "");
                string entryPoint = req.ContainsKey("entryPoint") ? req["entryPoint"] : null;
                string parameters = req.ContainsKey("parameters") ? req["parameters"] : null;
                string outputDir = req.ContainsKey("outputDir") ? req["outputDir"] : "";
                bool stopOnError = !req.ContainsKey("stopOnError") || (req["stopOnError"] == "true" || req["stopOnError"] == "1" || req["stopOnError"] == "True");

                if (jobDef != null)
                {
                    // 已有固化任务：若请求中重新传入了文件、代码、输出目录等字段，必须与固化定义严格一致，严禁静默改动
                    if (filePaths.Count > 0)
                    {
                        if (filePaths.Count != jobDef.files.Count)
                        {
                            return new BridgeResponse
                            {
                                ok = false,
                                action = "start_batch_job",
                                error = "启动请求传入的文件列表数量与已固化任务定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                            };
                        }
                        for (int i = 0; i < filePaths.Count; i++)
                        {
                            if (!string.Equals(filePaths[i], jobDef.files[i].originalFilePath, StringComparison.OrdinalIgnoreCase))
                            {
                                return new BridgeResponse
                                {
                                    ok = false,
                                    action = "start_batch_job",
                                    error = string.Format("启动请求第 {0} 个文件路径与已固化任务不一致。严禁静默修改，如需变更必须重新固化任务。", i + 1)
                                };
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(macroCode))
                    {
                        string inputMacroHash = BatchRunnerService.ComputeStringSha256(macroCode);
                        if (!string.Equals(inputMacroHash, jobDef.macroHash, StringComparison.OrdinalIgnoreCase))
                        {
                            return new BridgeResponse
                            {
                                ok = false,
                                action = "start_batch_job",
                                error = "启动请求传入的宏代码哈希与已固化任务定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                            };
                        }
                    }

                    if (!string.IsNullOrEmpty(entryPoint) && !string.Equals(entryPoint, jobDef.entryPoint ?? "", StringComparison.OrdinalIgnoreCase))
                    {
                        return new BridgeResponse
                        {
                            ok = false,
                            action = "start_batch_job",
                            error = "启动请求传入的入口过程与已固化任务定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                        };
                    }

                    if (!string.IsNullOrEmpty(parameters) && !string.Equals(parameters, jobDef.parametersJson ?? "", StringComparison.Ordinal))
                    {
                        return new BridgeResponse
                        {
                            ok = false,
                            action = "start_batch_job",
                            error = "启动请求传入的结构化参数与已固化任务定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                        };
                    }

                    if (req.ContainsKey("stopOnError"))
                    {
                        bool reqStop = (req["stopOnError"] == "true" || req["stopOnError"] == "1" || req["stopOnError"] == "True");
                        if (reqStop != jobDef.stopOnError)
                        {
                            return new BridgeResponse
                            {
                                ok = false,
                                action = "start_batch_job",
                                error = "启动请求的执行策略 (stopOnError) 与已固化定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                            };
                        }
                    }

                    if (!string.IsNullOrEmpty(outputDir))
                    {
                        string fullOut = Path.GetFullPath(outputDir);
                        if (!string.Equals(fullOut, jobDef.outputDir, StringComparison.OrdinalIgnoreCase))
                        {
                            return new BridgeResponse
                            {
                                ok = false,
                                action = "start_batch_job",
                                error = "启动请求传入的输出目录与已固化任务定义不一致。严禁静默修改，如需变更必须重新固化任务。"
                            };
                        }
                    }
                }
                else
                {
                    // 未找到固化定义：如果仅提供了 jobId 但后端未固化，阻断
                    if (!string.IsNullOrEmpty(jobId) && filePaths.Count == 0 && string.IsNullOrEmpty(macroCode))
                    {
                        return new BridgeResponse
                        {
                            ok = false,
                            action = "start_batch_job",
                            error = string.Format("未找到已固化的批量任务 ID: {0}。请先调用 validate_batch_job 固化队列并经用户确认后再启动。", jobId)
                        };
                    }

                    // 否则执行全新固化校验
                    string valErr;
                    jobDef = BatchRunnerService.ValidateAndLockJob(filePaths, macroCode, entryPoint, parameters, outputDir, stopOnError, out valErr);
                    if (jobDef == null)
                    {
                        return new BridgeResponse
                        {
                            ok = false,
                            action = "start_batch_job",
                            error = "批量任务固化校验失败: " + valErr
                        };
                    }
                    if (!string.IsNullOrEmpty(jobId))
                    {
                        jobDef.jobId = jobId;
                    }
                }

                bool isSync = req.ContainsKey("sync") && (req["sync"] == "true" || req["sync"] == "1" || req["sync"] == "True");
                if (isSync)
                {
                    var summary = BatchRunnerService.ExecuteBatchJob(jobDef, app);
                    return new BridgeResponse
                    {
                        ok = summary.status == "completed" || summary.status == "stopped_on_error" || summary.status == "cancelled",
                        action = "start_batch_job",
                        data = summary,
                        message = string.Format("批量任务已结束，状态: {0} (成功={1}, 失败={2}, 阻断={3}, 取消={4})",
                            summary.status, summary.successCount, summary.failedCount, summary.blockedCount, summary.cancelledCount)
                    };
                }
                else
                {
                    var initSummary = BatchRunnerService.StartBatchJobAsync(jobDef, app);
                    return new BridgeResponse
                    {
                        ok = true,
                        action = "start_batch_job",
                        data = initSummary,
                        message = "批量任务已在受控后台线程启动，正在串行执行中。"
                    };
                }
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "start_batch_job",
                    error = "启动批量任务异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleCancelBatchJob(Dictionary<string, string> req)
        {
            string jobId = req.ContainsKey("jobId") ? req["jobId"] : "";
            string msg;
            bool ok = BatchRunnerService.RequestCancel(jobId, out msg);
            return new BridgeResponse
            {
                ok = ok,
                action = "cancel_batch_job",
                message = msg,
                error = ok ? null : msg
            };
        }

        private static BridgeResponse HandleGetBatchJobStatus(Dictionary<string, string> req)
        {
            string jobId = req.ContainsKey("jobId") ? req["jobId"] : "";
            var summary = BatchRunnerService.GetJobSummary(jobId);
            if (summary == null)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_batch_job_status",
                    error = "未找到指定的批量任务 ID: " + jobId
                };
            }

            return new BridgeResponse
            {
                ok = true,
                action = "get_batch_job_status",
                data = summary
            };
        }

        private static BridgeResponse HandleOpenOutputFolder(Dictionary<string, string> req)
        {
            string jobId = req.ContainsKey("jobId") ? req["jobId"].Trim() : "";
            string outputDir = req.ContainsKey("outputDir") ? req["outputDir"].Trim() : null;

            if (string.IsNullOrEmpty(outputDir) && !string.IsNullOrEmpty(jobId))
            {
                var locked = BatchRunnerService.GetLockedJob(jobId);
                if (locked != null && !string.IsNullOrEmpty(locked.outputDir))
                {
                    outputDir = locked.outputDir;
                }
                else
                {
                    var summary = BatchRunnerService.GetJobSummary(jobId);
                    if (summary != null && summary.fileResults != null)
                    {
                        foreach (var f in summary.fileResults)
                        {
                            if (!string.IsNullOrEmpty(f.finalOutputPath))
                            {
                                outputDir = Path.GetDirectoryName(f.finalOutputPath);
                                break;
                            }
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(outputDir) || !Directory.Exists(outputDir))
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "open_output_folder",
                    error = "未找到本任务确认的有效输出目录: " + (outputDir ?? jobId)
                };
            }

            try
            {
                // 仅打开本次任务确认的输出目录，绝不接受外部任意可执行命令
                System.Diagnostics.Process.Start("explorer.exe", "\"" + outputDir + "\"");
                return new BridgeResponse
                {
                    ok = true,
                    action = "open_output_folder",
                    data = outputDir,
                    message = "已打开输出目录: " + outputDir
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "open_output_folder",
                    error = "打开输出目录失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleBrowseFiles(Dictionary<string, string> req)
        {
            try
            {
                var selectedFiles = new List<string>();
                Thread t = new Thread(() =>
                {
                    using (var ofd = new System.Windows.Forms.OpenFileDialog())
                    {
                        ofd.Multiselect = true;
                        ofd.Title = "请选择需要批量处理的 Excel 工作簿";
                        ofd.Filter = "Excel 工作簿 (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls";
                        if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            selectedFiles.AddRange(ofd.FileNames);
                        }
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                return new BridgeResponse
                {
                    ok = true,
                    action = "browse_files",
                    data = selectedFiles,
                    message = string.Format("已选择 {0} 个文件", selectedFiles.Count)
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "browse_files",
                    error = "选择文件失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleBrowseFolder(Dictionary<string, string> req)
        {
            try
            {
                string selectedPath = "";
                Thread t = new Thread(() =>
                {
                    using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        fbd.Description = "请选择批量处理输出目录（将生成带递增序号的独立输出文件）";
                        if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            selectedPath = fbd.SelectedPath;
                        }
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                if (string.IsNullOrEmpty(selectedPath))
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "browse_folder",
                        error = "未选择输出目录"
                    };
                }

                return new BridgeResponse
                {
                    ok = true,
                    action = "browse_folder",
                    data = selectedPath,
                    message = "已选择输出目录: " + selectedPath
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "browse_folder",
                    error = "选择输出目录失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleBrowseSaveFile(Dictionary<string, string> req)
        {
            try
            {
                string title = req != null && req.ContainsKey("title") ? req["title"] : "保存文件";
                string defaultName = req != null && req.ContainsKey("defaultName") ? req["defaultName"] : "macros.exmpack";
                string filter = req != null && req.ContainsKey("filter") ? req["filter"] : "ExcelMind 宏包 (*.exmpack)|*.exmpack|所有文件 (*.*)|*.*";
                string selectedPath = "";

                Thread t = new Thread(() =>
                {
                    using (var sfd = new System.Windows.Forms.SaveFileDialog())
                    {
                        sfd.Title = title;
                        sfd.FileName = defaultName;
                        sfd.Filter = filter;
                        sfd.DefaultExt = "exmpack";
                        sfd.AddExtension = true;
                        if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            selectedPath = sfd.FileName;
                        }
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                if (string.IsNullOrEmpty(selectedPath))
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "browse_save_file",
                        error = "已取消选择保存路径"
                    };
                }

                return new BridgeResponse
                {
                    ok = true,
                    action = "browse_save_file",
                    data = selectedPath,
                    message = "已选定保存路径: " + selectedPath
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "browse_save_file",
                    error = "选择保存路径失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleListWorkflows()
        {
            try
            {
                var list = WorkflowManager.ListWorkflows();
                return new BridgeResponse
                {
                    ok = true,
                    action = "list_workflows",
                    data = list
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "list_workflows",
                    error = ex.Message
                };
            }
        }

        private static BridgeResponse HandleGetWorkflow(Dictionary<string, string> req)
        {
            try
            {
                string id = req.ContainsKey("workflowId") ? req["workflowId"] : "";
                var def = WorkflowManager.GetWorkflow(id);
                if (def == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "get_workflow",
                        error = "找不到工作流定义: " + id
                    };
                }
                var runs = WorkflowManager.GetWorkflowRuns(id);
                return new BridgeResponse
                {
                    ok = true,
                    action = "get_workflow",
                    data = new Dictionary<string, object>
                    {
                        { "definition", def },
                        { "runs", runs }
                    }
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "get_workflow",
                    error = ex.Message
                };
            }
        }

        private static BridgeResponse HandleSaveWorkflow(string jsonString)
        {
            try
            {
                WorkflowDefinition def = null;
                int wfIdx = jsonString.IndexOf("\"workflow\":");
                if (wfIdx >= 0)
                {
                    string sub = jsonString.Substring(wfIdx + 11).Trim();
                    def = SimpleJson.Deserialize<WorkflowDefinition>(sub);
                }
                else
                {
                    def = SimpleJson.Deserialize<WorkflowDefinition>(jsonString);
                }

                if (def == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "save_workflow",
                        error = "工作流定义反序列化失败，请检查输入格式。"
                    };
                }

                var saved = WorkflowManager.SaveWorkflow(def);
                return new BridgeResponse
                {
                    ok = true,
                    action = "save_workflow",
                    data = saved,
                    message = "工作流定义已保存 (版本: v" + saved.definitionVersion + ")"
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "save_workflow",
                    error = "保存工作流失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleDeleteWorkflow(Dictionary<string, string> req)
        {
            try
            {
                string id = req.ContainsKey("workflowId") ? req["workflowId"] : "";
                bool deleted = WorkflowManager.DeleteWorkflow(id);
                return new BridgeResponse
                {
                    ok = deleted,
                    action = "delete_workflow",
                    message = deleted ? "工作流及执行记录已删除" : "未找到工作流或删除失败"
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "delete_workflow",
                    error = ex.Message
                };
            }
        }

        private static BridgeResponse HandleExecuteWorkflow(string jsonString, dynamic app)
        {
            try
            {
                var req = SimpleJson.ParseFlatObject(jsonString);
                string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
                string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";

                WorkflowDefinition def = null;
                if (req.ContainsKey("workflowId") && !string.IsNullOrEmpty(req["workflowId"]))
                {
                    def = WorkflowManager.GetWorkflow(req["workflowId"]);
                }

                if (def == null)
                {
                    int wfIdx = jsonString.IndexOf("\"workflow\":");
                    if (wfIdx >= 0)
                    {
                        string sub = jsonString.Substring(wfIdx + 11).Trim();
                        def = SimpleJson.Deserialize<WorkflowDefinition>(sub);
                    }
                }

                if (def == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "execute_workflow",
                        error = "未指定有效的工作流定义或 workflowId"
                    };
                }

                var runRecord = WorkflowManager.ExecuteWorkflow(def, app, targetFullName, targetName);
                bool isSuccess = runRecord.status == "completed";

                return new BridgeResponse
                {
                    ok = isSuccess,
                    action = "execute_workflow",
                    data = runRecord,
                    error = isSuccess ? null : runRecord.failureMessage,
                    message = isSuccess ? "流水线双步骤全部成功执行！" : runRecord.failureMessage
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_workflow",
                    error = "执行工作流异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleCancelWorkflow()
        {
            WorkflowManager.CancelWorkflow();
            return new BridgeResponse
            {
                ok = true,
                action = "cancel_workflow",
                message = "工作流取消指令已登记，将在当前步骤边界安全生效。"
            };
        }

        private static BridgeResponse HandleListManagedCharts(Dictionary<string, string> req, dynamic app)
        {
            try
            {
                string targetFullName = req.ContainsKey("targetWorkbookFullName") ? req["targetWorkbookFullName"] : "";
                string targetName = req.ContainsKey("targetWorkbookName") ? req["targetWorkbookName"] : "";
                string sheetName = req.ContainsKey("sheetName") ? req["sheetName"] : "";

                dynamic targetWb = FindTargetWorkbook(app, targetFullName, targetName);
                if (targetWb == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "list_managed_charts",
                        error = "未找到目标工作簿，系统拒绝隐式回退当前活动工作簿！"
                    };
                }

                var list = ChartService.ListManagedCharts(targetWb, sheetName);
                return new BridgeResponse
                {
                    ok = true,
                    action = "list_managed_charts",
                    data = list
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "list_managed_charts",
                    error = "枚举图表失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleExecuteQuickChart(string jsonString, dynamic app)
        {
            try
            {
                QuickChartParams p = null;
                int paramsIdx = jsonString.IndexOf("\"params\":");
                if (paramsIdx >= 0)
                {
                    string sub = jsonString.Substring(paramsIdx + 9).Trim();
                    p = SimpleJson.Deserialize<QuickChartParams>(sub);
                }
                else
                {
                    p = SimpleJson.Deserialize<QuickChartParams>(jsonString);
                }

                if (p == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "execute_quick_chart",
                        error = "解析图表参数失败，请检查输入格式。"
                    };
                }

                var result = ChartService.ExecuteQuickChart(app, p);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "execute_quick_chart",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? result.summaryText : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "execute_quick_chart",
                    error = "执行快捷图表生成异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandlePreviewExternalData(string jsonString)
        {
            try
            {
                ExternalDataPreviewParams p = null;
                int paramsIdx = jsonString.IndexOf("\"params\":");
                if (paramsIdx >= 0)
                {
                    string sub = jsonString.Substring(paramsIdx + 9).Trim();
                    p = SimpleJson.Deserialize<ExternalDataPreviewParams>(sub);
                }
                else
                {
                    p = SimpleJson.Deserialize<ExternalDataPreviewParams>(jsonString);
                }

                if (p == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "preview_external_data",
                        error = "解析外部数据预览参数失败。"
                    };
                }

                var result = ExternalDataService.Preview(p);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "preview_external_data",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? result.sanitizedSummary : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "preview_external_data",
                    error = "外部数据预览异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleImportExternalData(string jsonString, dynamic app)
        {
            try
            {
                ExternalDataImportParams p = null;
                int paramsIdx = jsonString.IndexOf("\"params\":");
                if (paramsIdx >= 0)
                {
                    string sub = jsonString.Substring(paramsIdx + 9).Trim();
                    p = SimpleJson.Deserialize<ExternalDataImportParams>(sub);
                }
                else
                {
                    p = SimpleJson.Deserialize<ExternalDataImportParams>(jsonString);
                }

                if (p == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "import_external_data",
                        error = "解析外部数据导入参数失败。"
                    };
                }

                var result = ExternalDataService.Import(p, app);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "import_external_data",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? string.Format("已成功导入 {0} 行数据至工作表【{1}】", result.importedRowCount, result.sheetName) : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "import_external_data",
                    error = "外部数据导入异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleListDataSourceConfigs()
        {
            try
            {
                var list = ExternalDataService.ListDataSourceConfigs();
                return new BridgeResponse
                {
                    ok = true,
                    action = "list_data_source_configs",
                    data = list
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "list_data_source_configs",
                    error = "获取数据源配置列表失败: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleSaveDataSourceConfig(string jsonString)
        {
            try
            {
                DataSourceConfigDto dto = null;
                string secret = null;

                int paramsIdx = jsonString.IndexOf("\"params\":");
                string targetJson = paramsIdx >= 0 ? jsonString.Substring(paramsIdx + 9).Trim() : jsonString;

                dto = SimpleJson.Deserialize<DataSourceConfigDto>(targetJson);
                var flat = SimpleJson.ParseFlatObject(targetJson);
                if (flat.ContainsKey("secret")) secret = flat["secret"];
                else if (flat.ContainsKey("credential")) secret = flat["credential"];
                else if (flat.ContainsKey("token")) secret = flat["token"];

                if (dto == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "save_data_source_config",
                        error = "解析数据源配置失败。"
                    };
                }

                ExternalDataService.SaveDataSourceConfig(dto, secret);
                return new BridgeResponse
                {
                    ok = true,
                    action = "save_data_source_config",
                    data = dto,
                    message = "数据源配置已保存（凭据已通过 Windows DPAPI 加密存储）"
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "save_data_source_config",
                    error = "保存数据源配置异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleExportMacroPackage(string jsonString)
        {
            try
            {
                MacroPackageExportParams p = null;
                var flat = SimpleJson.ParseFlatObject(jsonString);
                string targetJson = jsonString;
                if (flat.ContainsKey("params") && !string.IsNullOrEmpty(flat["params"]))
                {
                    targetJson = flat["params"].Trim();
                }
                else
                {
                    int paramsIdx = jsonString.IndexOf("\"params\":");
                    if (paramsIdx >= 0)
                    {
                        targetJson = jsonString.Substring(paramsIdx + 9).Trim();
                    }
                }

                p = SimpleJson.Deserialize<MacroPackageExportParams>(targetJson);

                if (p == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "export_macro_package",
                        error = "解析宏包导出参数失败。"
                    };
                }

                var result = MacroPackageManager.ExportPackage(p);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "export_macro_package",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? string.Format("宏包导出成功: {0} 个宏已打包至 {1}", result.exportedCount, Path.GetFileName(result.packageFilePath)) : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "export_macro_package",
                    error = "导出宏包异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandlePreviewMacroPackage(string jsonString)
        {
            try
            {
                string packageFilePath = "";
                var flat = SimpleJson.ParseFlatObject(jsonString);
                if (flat.ContainsKey("packageFilePath")) packageFilePath = flat["packageFilePath"];
                else if (flat.ContainsKey("filePath")) packageFilePath = flat["filePath"];
                else if (flat.ContainsKey("path")) packageFilePath = flat["path"];

                int paramsIdx = jsonString.IndexOf("\"params\":");
                if (paramsIdx >= 0)
                {
                    string sub = jsonString.Substring(paramsIdx + 9).Trim();
                    var pFlat = SimpleJson.ParseFlatObject(sub);
                    if (pFlat.ContainsKey("packageFilePath")) packageFilePath = pFlat["packageFilePath"];
                    else if (pFlat.ContainsKey("filePath")) packageFilePath = pFlat["filePath"];
                    else if (pFlat.ContainsKey("path")) packageFilePath = pFlat["path"];
                }

                var result = MacroPackageManager.PreviewPackage(packageFilePath);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "preview_macro_package",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? string.Format("宏包解析成功，共包含 {0} 个宏条目", result.macroCount) : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "preview_macro_package",
                    error = "预览宏包异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleImportMacroPackage(string jsonString)
        {
            try
            {
                MacroPackageImportParams p = null;
                var flat = SimpleJson.ParseFlatObject(jsonString);
                string targetJson = jsonString;
                if (flat.ContainsKey("params") && !string.IsNullOrEmpty(flat["params"]))
                {
                    targetJson = flat["params"].Trim();
                }
                else
                {
                    int paramsIdx = jsonString.IndexOf("\"params\":");
                    if (paramsIdx >= 0)
                    {
                        targetJson = jsonString.Substring(paramsIdx + 9).Trim();
                    }
                }

                p = SimpleJson.Deserialize<MacroPackageImportParams>(targetJson);

                if (p == null)
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "import_macro_package",
                        error = "解析宏包导入参数失败。"
                    };
                }

                var result = MacroPackageManager.ImportPackage(p);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "import_macro_package",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? string.Format("宏包导入成功: 成功导入 {0} 个宏到宏库（零自动执行）", result.importedCount) : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "import_macro_package",
                    error = "导入宏包异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandlePreviewDiagnostics(string jsonString, dynamic app)
        {
            try
            {
                string customLogs = null;
                if (!string.IsNullOrEmpty(jsonString))
                {
                    int paramsIdx = jsonString.IndexOf("\"params\":");
                    if (paramsIdx >= 0)
                    {
                        string sub = jsonString.Substring(paramsIdx + 9).Trim();
                        var pFlat = SimpleJson.ParseFlatObject(sub);
                        if (pFlat.ContainsKey("customLogContent")) customLogs = pFlat["customLogContent"];
                        else if (pFlat.ContainsKey("logs")) customLogs = pFlat["logs"];
                    }
                }

                var result = DiagnosticsService.PreviewDiagnostics(app, customLogs);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "preview_diagnostics",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? "系统环境与诊断预览生成成功" : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "preview_diagnostics",
                    error = "生成诊断预览异常: " + ex.Message
                };
            }
        }

        private static BridgeResponse HandleExportDiagnostics(string jsonString, dynamic app)
        {
            try
            {
                string targetZipPath = null;
                string customLogs = null;

                if (!string.IsNullOrEmpty(jsonString))
                {
                    var rootObj = SimpleJson.ParseFlatObject(jsonString);
                    if (rootObj.ContainsKey("targetZipPath")) targetZipPath = rootObj["targetZipPath"];
                    else if (rootObj.ContainsKey("filePath")) targetZipPath = rootObj["filePath"];
                    else if (rootObj.ContainsKey("outputPath")) targetZipPath = rootObj["outputPath"];

                    if (rootObj.ContainsKey("customLogContent")) customLogs = rootObj["customLogContent"];
                    else if (rootObj.ContainsKey("logs")) customLogs = rootObj["logs"];

                    int paramsIdx = jsonString.IndexOf("\"params\":");
                    if (paramsIdx >= 0)
                    {
                        string sub = jsonString.Substring(paramsIdx + 9).Trim();
                        var pFlat = SimpleJson.ParseFlatObject(sub);
                        if (pFlat.ContainsKey("targetZipPath")) targetZipPath = pFlat["targetZipPath"];
                        else if (pFlat.ContainsKey("filePath")) targetZipPath = pFlat["filePath"];
                        else if (pFlat.ContainsKey("outputPath")) targetZipPath = pFlat["outputPath"];

                        if (pFlat.ContainsKey("customLogContent")) customLogs = pFlat["customLogContent"];
                        else if (pFlat.ContainsKey("logs")) customLogs = pFlat["logs"];
                    }
                }

                if (targetZipPath == "__CANCEL__" || targetZipPath == "CANCEL")
                {
                    return new BridgeResponse
                    {
                        ok = false,
                        action = "export_diagnostics",
                        error = "已取消选择保存路径"
                    };
                }

                // 若未传入路径，调起原生另存为对话框
                if (string.IsNullOrEmpty(targetZipPath))
                {
                    Thread t = new Thread(() =>
                    {
                        using (var sfd = new System.Windows.Forms.SaveFileDialog())
                        {
                            sfd.Title = "选择脱敏诊断包保存路径";
                            sfd.FileName = string.Format("ExcelMindAI_Diagnostics_{0}.zip", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                            sfd.Filter = "ZIP 压缩包 (*.zip)|*.zip|所有文件 (*.*)|*.*";
                            sfd.DefaultExt = "zip";
                            sfd.AddExtension = true;
                            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                targetZipPath = sfd.FileName;
                            }
                        }
                    });
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();
                    t.Join();

                    if (string.IsNullOrEmpty(targetZipPath))
                    {
                        return new BridgeResponse
                        {
                            ok = false,
                            action = "export_diagnostics",
                            error = "已取消选择保存路径"
                        };
                    }
                }

                var result = DiagnosticsService.ExportDiagnostics(app, targetZipPath, customLogs);
                return new BridgeResponse
                {
                    ok = result.ok,
                    action = "export_diagnostics",
                    data = result,
                    error = result.ok ? null : result.error,
                    message = result.ok ? string.Format("脱敏诊断包已成功导出至: {0}", result.zipFilePath) : result.error
                };
            }
            catch (Exception ex)
            {
                return new BridgeResponse
                {
                    ok = false,
                    action = "export_diagnostics",
                    error = "导出诊断包异常: " + ex.Message
                };
            }
        }
    }
}

