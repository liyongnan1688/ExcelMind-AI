using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace LeeExcel
{
    /// <summary>
    /// 去重步骤专用参数 DTO
    /// </summary>
    public class DedupStepParams
    {
        public string targetSheet { get; set; }
        public string sheetName { get { return targetSheet; } set { targetSheet = value; } }
        public string rangeAddress { get; set; }
        public bool hasHeader { get; set; }
        public List<int> keyColumns { get; set; }
        public string mode { get; set; } // "highlight" | "export_unique"
        public string outputSheetName { get; set; }
        public string expectedFingerprint { get; set; }

        public DedupStepParams()
        {
            keyColumns = new List<int>();
            mode = "export_unique";
            hasHeader = true;
        }
    }

    /// <summary>
    /// 对账步骤专用参数 DTO
    /// </summary>
    public class ReconcileStepParams
    {
        public string leftSheet { get; set; }
        public string leftSheetName { get { return leftSheet; } set { leftSheet = value; } }
        public string leftAddress { get; set; }
        public string leftRangeAddress { get { return leftAddress; } set { leftAddress = value; } }
        public bool leftHasHeader { get; set; }
        public List<int> leftKeyCols { get; set; }

        public string rightSheet { get; set; }
        public string rightSheetName { get { return rightSheet; } set { rightSheet = value; } }
        public string rightAddress { get; set; }
        public string rightRangeAddress { get { return rightAddress; } set { rightAddress = value; } }
        public bool rightHasHeader { get; set; }
        public List<int> rightKeyCols { get; set; }

        public List<CompareColMappingDto> compareCols { get; set; }
        public string outputSheetName { get; set; }
        public string expectedFingerprint { get; set; }

        public ReconcileStepParams()
        {
            leftKeyCols = new List<int>();
            rightKeyCols = new List<int>();
            compareCols = new List<CompareColMappingDto>();
            leftHasHeader = true;
            rightHasHeader = true;
        }
    }

    /// <summary>
    /// 已保存宏步骤专用参数 DTO
    /// </summary>
    public class SavedMacroStepParams
    {
        public string scriptId { get; set; }
        public string entryPoint { get; set; }
        public string macroSha256 { get; set; } // 定义保存时固化的宏正文哈希
        public string targetSheet { get; set; }
        public string targetRange { get; set; }
        public List<VbaParameterInput> parameters { get; set; }
        public string declaredOutputSheet { get; set; } // 用户声明的输出工作表 (供下一步验证与消费)
        public string declaredOutputRange { get; set; }

        public SavedMacroStepParams()
        {
            parameters = new List<VbaParameterInput>();
        }
    }

    /// <summary>
    /// 图表步骤专用参数 DTO (TASK-R5b-01)
    /// </summary>
    public class ChartStepParams
    {
        public string sourceSheet { get; set; }
        public string sourceRange { get; set; }
        public bool hasHeaders { get; set; }
        public int dataStartRow { get; set; }
        public int dataEndRow { get; set; }
        public int categoryColIndex { get; set; }
        public string categoryColName { get; set; }
        public List<int> seriesColIndices { get; set; }
        public List<string> seriesNames { get; set; }
        public string chartType { get; set; }
        public string title { get; set; }
        public string targetSheet { get; set; }
        public string placementMode { get; set; }
        public string targetCell { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public string action { get; set; }
        public string targetChartId { get; set; }
        public string errorHandling { get; set; }

        // 别名兼容属性（保障流水线执行中对不同风格命名的 100% 互操作性）
        public string rangeAddress { get { return sourceRange; } set { if (!string.IsNullOrEmpty(value)) sourceRange = value; } }
        public bool hasHeader { get { return hasHeaders; } set { hasHeaders = value; } }
        public int categoryColumn { get { return categoryColIndex; } set { if (value > 0) categoryColIndex = value; } }
        public List<int> valueSeriesColumns { get { return seriesColIndices; } set { if (value != null && value.Count > 0) seriesColIndices = value; } }
        public string chartTitle { get { return title; } set { if (!string.IsNullOrEmpty(value)) title = value; } }
        public string placementSheet { get { return targetSheet; } set { if (!string.IsNullOrEmpty(value)) targetSheet = value; } }
        public string placementCell { get { return targetCell; } set { if (!string.IsNullOrEmpty(value)) targetCell = value; } }
        public string mode { get { return action; } set { if (!string.IsNullOrEmpty(value)) action = value; } }
        public string errorHandlingRule { get { return errorHandling; } set { if (!string.IsNullOrEmpty(value)) errorHandling = value; } }

        public ChartStepParams()
        {
            hasHeaders = true;
            dataStartRow = 2;
            dataEndRow = 0;
            categoryColIndex = 1;
            seriesColIndices = new List<int>();
            seriesNames = new List<string>();
            chartType = "column";
            placementMode = "cell";
            targetCell = "F2";
            width = 480;
            height = 300;
            action = "create_new";
            errorHandling = "reject_on_invalid";
        }
    }

    /// <summary>
    /// 步骤执行输出结构化引用契约
    /// </summary>
    public class StepOutputReference
    {
        public string targetWorkbookName { get; set; }
        public string sheetName { get; set; }
        public string rangeAddress { get; set; }
        public bool hasHeader { get; set; }
        public int headerRowIndex { get; set; }
        public int rowCount { get; set; }
        public int columnCount { get; set; }
        public string dataFingerprint { get; set; }
    }

    /// <summary>
    /// 工作流步骤定义 DTO（严格限定为双步骤中的单步）
    /// </summary>
    public class WorkflowStepDefinition
    {
        public int stepIndex { get; set; } // 1 或 2
        public string stepName { get; set; }
        public string toolType { get; set; } // "dedup" | "reconcile" | "saved_macro" | "chart"
        public string inputSource { get; set; } // "initial_selection" | "prev_step_output"

        public DedupStepParams dedupParams { get; set; }
        public ReconcileStepParams reconcileParams { get; set; }
        public SavedMacroStepParams macroParams { get; set; }
        public ChartStepParams chartParams { get; set; }

        public WorkflowStepDefinition()
        {
            inputSource = "initial_selection";
        }
    }

    /// <summary>
    /// 工作流定义模型（支持本地 JSON 持久化与版本递增）
    /// </summary>
    public class WorkflowDefinition
    {
        public string workflowId { get; set; }
        public int definitionVersion { get; set; } // 1, 2, ...
        public string name { get; set; }
        public string description { get; set; }
        public string targetWorkbookName { get; set; }
        public string targetWorkbookPath { get; set; }
        public bool isPreset { get; set; }
        public List<WorkflowStepDefinition> steps { get; set; } // 必须且仅包含 2 个步骤
        public string createdAt { get; set; }
        public string updatedAt { get; set; }

        public WorkflowDefinition()
        {
            definitionVersion = 1;
            steps = new List<WorkflowStepDefinition>();
        }
    }

    /// <summary>
    /// 单步执行结果
    /// </summary>
    public class StepRunResult
    {
        public int stepIndex { get; set; }
        public string stepName { get; set; }
        public string toolType { get; set; }
        public string status { get; set; } // "success" | "failed" | "skipped" | "blocked"
        public string failureStage { get; set; }
        public string summaryText { get; set; }
        public string error { get; set; }
        public StepOutputReference outputRef { get; set; }
        public long elapsedMs { get; set; }
    }

    /// <summary>
    /// 工作流执行记录 DTO（与定义分开保存）
    /// </summary>
    public class WorkflowRunRecord
    {
        public string runId { get; set; }
        public string workflowId { get; set; }
        public int definitionVersion { get; set; }
        public string executedAt { get; set; }
        public string targetWorkbookName { get; set; }
        public string targetWorkbookPath { get; set; }
        public string snapshotId { get; set; }
        public bool snapshotExists { get; set; }
        public string status { get; set; } // "completed" | "stopped_on_step1" | "stopped_on_step2" | "blocked" | "cancelled"
        public string phase { get; set; } // "precheck" | "snapshot" | "step1" | "step2" | "completed"
        public int completedSteps { get; set; } // 0, 1, 2
        public int failedStepIndex { get; set; } // 0, 1, 2
        public string failureMessage { get; set; }
        public string recoveryNotice { get; set; }
        public List<StepRunResult> stepResults { get; set; }
        public long totalElapsedMs { get; set; }

        public WorkflowRunRecord()
        {
            stepResults = new List<StepRunResult>();
        }
    }

    /// <summary>
    /// 双步骤任务流水线管理与调度引擎 (TASK-R5a-01)
    /// </summary>
    public static class WorkflowManager
    {
        public static readonly string WorkflowsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelMindAI",
            "Workflows"
        );

        public static readonly string RunsDir = Path.Combine(WorkflowsDir, "Runs");

        private static readonly object _fileLock = new object();
        private static readonly object _runLock = new object();
        private static volatile bool _isRunning = false;
        private static volatile bool _cancelRequested = false;

        static WorkflowManager()
        {
            EnsureDirectories();
            InitializePresetsIfEmpty();
        }

        public static void EnsureDirectories()
        {
            try
            {
                if (!Directory.Exists(WorkflowsDir)) Directory.CreateDirectory(WorkflowsDir);
                if (!Directory.Exists(RunsDir)) Directory.CreateDirectory(RunsDir);
            }
            catch { }
        }

        private static void InitializePresetsIfEmpty()
        {
            try
            {
                if (!Directory.Exists(WorkflowsDir)) return;
                var files = Directory.GetFiles(WorkflowsDir, "wf_*.json");
                if (files.Length > 0) return;

                // 预设 1: 去重导出唯一表 -> 两表主键对账
                var p1 = new WorkflowDefinition
                {
                    workflowId = "wf_preset_dedup_reconcile",
                    definitionVersion = 1,
                    name = "去重导出唯一表 → 两表对账",
                    description = "第一步从指定数据表中按主键提取唯一行导出新表；第二步自动将该唯一新表与基准表执行两表主键差异对账。",
                    isPreset = true,
                    createdAt = DateTime.Now.ToString("o"),
                    updatedAt = DateTime.Now.ToString("o"),
                    steps = new List<WorkflowStepDefinition>
                    {
                        new WorkflowStepDefinition
                        {
                            stepIndex = 1,
                            stepName = "按主键提取唯一记录",
                            toolType = "dedup",
                            inputSource = "initial_selection",
                            dedupParams = new DedupStepParams
                            {
                                mode = "export_unique",
                                hasHeader = true,
                                keyColumns = new List<int> { 1 }
                            }
                        },
                        new WorkflowStepDefinition
                        {
                            stepIndex = 2,
                            stepName = "唯一表与基准表对账",
                            toolType = "reconcile",
                            inputSource = "prev_step_output", // 明确消费前序输出
                            reconcileParams = new ReconcileStepParams
                            {
                                leftHasHeader = true,
                                leftKeyCols = new List<int> { 1 },
                                rightHasHeader = true,
                                rightKeyCols = new List<int> { 1 }
                            }
                        }
                    }
                };
                SaveWorkflowInternal(p1, false);

                // 预设 2: 已保存宏连续调度流水线
                var p2 = new WorkflowDefinition
                {
                    workflowId = "wf_preset_macro_chain",
                    definitionVersion = 1,
                    name = "双宏顺序批处理流水线",
                    description = "在同一任务目标工作簿中连续执行两个已保存的经过验证的宏，执行前全本备份，第一步失败安全阻断。",
                    isPreset = true,
                    createdAt = DateTime.Now.ToString("o"),
                    updatedAt = DateTime.Now.ToString("o"),
                    steps = new List<WorkflowStepDefinition>
                    {
                        new WorkflowStepDefinition
                        {
                            stepIndex = 1,
                            stepName = "前置处理宏",
                            toolType = "saved_macro",
                            inputSource = "initial_selection",
                            macroParams = new SavedMacroStepParams()
                        },
                        new WorkflowStepDefinition
                        {
                            stepIndex = 2,
                            stepName = "后续加工宏",
                            toolType = "saved_macro",
                            inputSource = "initial_selection",
                            macroParams = new SavedMacroStepParams()
                        }
                    }
                };
                SaveWorkflowInternal(p2, false);
            }
            catch { }
        }

        public static List<WorkflowDefinition> ListWorkflows()
        {
            EnsureDirectories();
            var list = new List<WorkflowDefinition>();
            lock (_fileLock)
            {
                try
                {
                    var files = Directory.GetFiles(WorkflowsDir, "wf_*.json");
                    foreach (var f in files)
                    {
                        try
                        {
                            string json = File.ReadAllText(f, Encoding.UTF8);
                            var def = SimpleJson.Deserialize<WorkflowDefinition>(json);
                            if (def != null && !string.IsNullOrEmpty(def.workflowId))
                            {
                                list.Add(def);
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }
            return list;
        }

        public static WorkflowDefinition GetWorkflow(string workflowId)
        {
            if (string.IsNullOrEmpty(workflowId)) return null;
            EnsureDirectories();
            lock (_fileLock)
            {
                try
                {
                    string path = Path.Combine(WorkflowsDir, workflowId + ".json");
                    if (!File.Exists(path)) return null;
                    string json = File.ReadAllText(path, Encoding.UTF8);
                    return SimpleJson.Deserialize<WorkflowDefinition>(json);
                }
                catch { return null; }
            }
        }

        public static List<WorkflowRunRecord> GetWorkflowRuns(string workflowId)
        {
            if (string.IsNullOrEmpty(workflowId)) return new List<WorkflowRunRecord>();
            EnsureDirectories();
            lock (_fileLock)
            {
                try
                {
                    string path = Path.Combine(RunsDir, workflowId + "_runs.json");
                    if (!File.Exists(path)) return new List<WorkflowRunRecord>();
                    string json = File.ReadAllText(path, Encoding.UTF8);
                    return SimpleJson.DeserializeList<WorkflowRunRecord>(json) ?? new List<WorkflowRunRecord>();
                }
                catch { return new List<WorkflowRunRecord>(); }
            }
        }

        public static WorkflowDefinition SaveWorkflow(WorkflowDefinition def)
        {
            return SaveWorkflowInternal(def, true);
        }

        private static WorkflowDefinition SaveWorkflowInternal(WorkflowDefinition def, bool checkVersionIncrement)
        {
            if (def == null) throw new ArgumentNullException("def");
            EnsureDirectories();

            lock (_fileLock)
            {
                if (string.IsNullOrEmpty(def.workflowId))
                {
                    def.workflowId = "wf_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    def.definitionVersion = 1;
                    def.createdAt = DateTime.Now.ToString("o");
                }
                else if (checkVersionIncrement)
                {
                    string existingPath = Path.Combine(WorkflowsDir, def.workflowId + ".json");
                    if (File.Exists(existingPath))
                    {
                        try
                        {
                            string oldJson = File.ReadAllText(existingPath, Encoding.UTF8);
                            var oldDef = SimpleJson.Deserialize<WorkflowDefinition>(oldJson);
                            if (oldDef != null)
                            {
                                def.definitionVersion = oldDef.definitionVersion + 1;
                                def.createdAt = oldDef.createdAt ?? DateTime.Now.ToString("o");
                            }
                        }
                        catch { }
                    }
                }

                def.updatedAt = DateTime.Now.ToString("o");

                // 对已保存宏固化 MacroSha256
                if (def.steps != null)
                {
                    foreach (var s in def.steps)
                    {
                        if (s.toolType == "saved_macro" && s.macroParams != null && !string.IsNullOrEmpty(s.macroParams.scriptId))
                        {
                            try
                            {
                                var script = ScriptManager.GetScriptById(s.macroParams.scriptId);
                                if (script != null && !string.IsNullOrEmpty(script.code))
                                {
                                    s.macroParams.macroSha256 = VbaRunner.ComputeSha256(script.code);
                                    if (string.IsNullOrEmpty(s.macroParams.entryPoint))
                                    {
                                        s.macroParams.entryPoint = script.entryPoint;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }

                string path = Path.Combine(WorkflowsDir, def.workflowId + ".json");
                string json = SimpleJson.Serialize(def);
                File.WriteAllText(path, json, Encoding.UTF8);
                return def;
            }
        }

        public static bool DeleteWorkflow(string workflowId)
        {
            if (string.IsNullOrEmpty(workflowId)) return false;
            EnsureDirectories();
            lock (_fileLock)
            {
                try
                {
                    string defPath = Path.Combine(WorkflowsDir, workflowId + ".json");
                    if (File.Exists(defPath)) File.Delete(defPath);

                    string runPath = Path.Combine(RunsDir, workflowId + "_runs.json");
                    if (File.Exists(runPath)) File.Delete(runPath);
                    return true;
                }
                catch { return false; }
            }
        }

        public static void CancelWorkflow()
        {
            _cancelRequested = true;
        }

        public static bool IsRunning
        {
            get { return _isRunning; }
        }

        /// <summary>
        /// 执行指定工作流（两步骤串联、强制前置快照、失败停止后续、现场保留与明确恢复范围）
        /// </summary>
        public static WorkflowRunRecord ExecuteWorkflow(
            WorkflowDefinition def,
            dynamic app,
            string targetFullName,
            string targetName)
        {
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            var record = new WorkflowRunRecord
            {
                runId = "wfrun_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4),
                workflowId = def != null ? def.workflowId : "unknown",
                definitionVersion = def != null ? def.definitionVersion : 1,
                executedAt = DateTime.Now.ToString("o"),
                targetWorkbookName = targetName,
                targetWorkbookPath = targetFullName,
                status = "blocked",
                phase = "precheck",
                completedSteps = 0,
                failedStepIndex = 0
            };

            lock (_runLock)
            {
                if (_isRunning)
                {
                    record.failureMessage = "已有工作流正在执行中，请勿重复启动。";
                    return record;
                }
                _isRunning = true;
                _cancelRequested = false;
            }

            try
            {
                // 1. 结构与参数前置校验
                if (def == null || def.steps == null || def.steps.Count != 2)
                {
                    record.failureMessage = "工作流定义必须且仅允许包含严格 2 个步骤。";
                    return record;
                }

                var step1 = def.steps[0];
                var step2 = def.steps[1];

                // 2. 切片支持组合校验 (严禁宣称任意组合可执行，明确不支持 consolidation 外部工作簿)
                bool isSupportedCombination = false;
                if (step1.toolType == "dedup" && step2.toolType == "reconcile") isSupportedCombination = true;
                else if (step1.toolType == "saved_macro" && step2.toolType == "saved_macro") isSupportedCombination = true;
                else if (step1.toolType == "dedup" && step2.toolType == "saved_macro") isSupportedCombination = true;
                else if (step1.toolType == "dedup" && step2.toolType == "chart") isSupportedCombination = true;
                else if (step1.toolType == "reconcile" && step2.toolType == "chart") isSupportedCombination = true;
                else if (step1.toolType == "saved_macro" && step2.toolType == "chart") isSupportedCombination = true;

                if (!isSupportedCombination)
                {
                    record.phase = "precheck";
                    record.failureMessage = string.Format(
                        "当前第一切片不支持步骤组合【{0} → {1}】。目前已验证并支持：去重输出→对账、已保存宏→已保存宏、去重输出→已保存宏、去重输出→图表、对账输出→图表、已保存宏→图表。多文件汇总涉及跨工作簿外部生命周期与独立回滚范围，因创建外部新工作簿生命周期不匹配，本切片明确暂不支持。",
                        step1.toolType, step2.toolType);
                    return record;
                }

                // 3. 目标工作簿身份可靠解析与锁定（严禁 fallback 到 ActiveWorkbook）
                if (app == null)
                {
                    record.failureMessage = "Excel 宿主实例为空。";
                    return record;
                }

                dynamic targetWb = null;
                try
                {
                    if (!string.IsNullOrEmpty(targetFullName))
                    {
                        foreach (dynamic wb in app.Workbooks)
                        {
                            if (string.Equals((string)wb.FullName, targetFullName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWb = wb;
                                break;
                            }
                        }
                    }
                    if (targetWb == null && !string.IsNullOrEmpty(targetName))
                    {
                        foreach (dynamic wb in app.Workbooks)
                        {
                            if (string.Equals((string)wb.Name, targetName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWb = wb;
                                break;
                            }
                        }
                    }
                }
                catch (Exception exFind)
                {
                    record.failureMessage = "查找目标工作簿异常: " + exFind.Message;
                    return record;
                }

                if (targetWb == null)
                {
                    record.failureMessage = string.Format("未找到指定的目标工作簿 (FullName: {0}, Name: {1})。系统严格拒绝隐式回退当前活动工作簿，防止串改！",
                        targetFullName ?? "空", targetName ?? "空");
                    return record;
                }

                try { record.targetWorkbookName = (string)targetWb.Name; } catch { }
                try { record.targetWorkbookPath = (string)targetWb.FullName; } catch { }

                // 4. 已保存宏哈希漂移校验
                if (step1.toolType == "saved_macro")
                {
                    string hashErr = ValidateMacroIntegrity(step1);
                    if (!string.IsNullOrEmpty(hashErr))
                    {
                        record.failureMessage = "第 1 步宏代码校验失败: " + hashErr;
                        return record;
                    }
                }
                if (step2.toolType == "saved_macro")
                {
                    string hashErr = ValidateMacroIntegrity(step2);
                    if (!string.IsNullOrEmpty(hashErr))
                    {
                        record.failureMessage = "第 2 步宏代码校验失败: " + hashErr;
                        return record;
                    }
                }

                // 5. 创建任务级前置快照（保护的工作簿必须与两步实际写入的工作簿绝对一致）
                record.phase = "snapshot";
                string snapPrompt = string.Format("工作流前置快照: 【{0}】(v{1}) 两步流水线执行前整本物理备份", def.name, def.definitionVersion);
                SnapshotItem snapshotItem = null;
                try
                {
                    snapshotItem = SnapshotManager.CreateSnapshot(targetWb, snapPrompt, "");
                }
                catch (Exception exSnap)
                {
                    record.failureMessage = "任务级前置快照创建异常 (" + exSnap.Message + ")。承诺快照失败零步骤执行，已安全阻断。";
                    return record;
                }

                if (snapshotItem == null || string.IsNullOrEmpty(snapshotItem.id) || !SnapshotManager.SnapshotExists(record.targetWorkbookPath, snapshotItem.id))
                {
                    record.failureMessage = "任务级前置快照物理文件写入失败。承诺快照失败零步骤执行，已安全阻断。";
                    return record;
                }

                record.snapshotId = snapshotItem.id;
                record.snapshotExists = true;

                // 6. 执行 Step 1
                record.phase = "step1";
                if (_cancelRequested)
                {
                    record.status = "cancelled";
                    record.failureMessage = "任务在执行第 1 步前被用户取消。";
                    return record;
                }

                var r1 = ExecuteStep(step1, app, targetWb, null);
                record.stepResults.Add(r1);

                if (r1.status != "success")
                {
                    record.status = "stopped_on_step1";
                    record.failedStepIndex = 1;
                    record.failureMessage = "第 1 步执行失败: " + (r1.error ?? r1.summaryText);
                    record.recoveryNotice = string.Format("失败后停止后续步骤；第 2 步已跳过。用户可按已验证范围恢复任务目标工作簿（快照 ID: {0}）。", record.snapshotId);
                    
                    // 标记 Step 2 为 skipped
                    record.stepResults.Add(new StepRunResult
                    {
                        stepIndex = 2,
                        stepName = step2.stepName,
                        toolType = step2.toolType,
                        status = "skipped",
                        summaryText = "因第 1 步失败，第 2 步已安全跳过，零执行。"
                    });
                    return record;
                }

                record.completedSteps = 1;

                // 7. 步骤边界取消检测
                if (_cancelRequested)
                {
                    record.status = "cancelled";
                    record.failureMessage = "第 1 步已执行完成；任务在步骤边界被用户取消，第 2 步已跳过。";
                    record.recoveryNotice = string.Format("第 1 步产生的结果已如实保留。若需撤销，用户可按已验证范围恢复任务目标工作簿（快照 ID: {0}）。", record.snapshotId);
                    record.stepResults.Add(new StepRunResult
                    {
                        stepIndex = 2,
                        stepName = step2.stepName,
                        toolType = step2.toolType,
                        status = "skipped",
                        summaryText = "任务在步骤边界被用户取消，第 2 步已跳过。"
                    });
                    return record;
                }

                // 8. 执行 Step 2（消费 Step 1 确权输出）
                record.phase = "step2";
                var r2 = ExecuteStep(step2, app, targetWb, r1.outputRef);
                record.stepResults.Add(r2);

                if (r2.status != "success")
                {
                    record.status = "stopped_on_step2";
                    record.failedStepIndex = 2;
                    record.failureMessage = "第 2 步执行失败: " + (r2.error ?? r2.summaryText);
                    record.recoveryNotice = string.Format("第 2 步失败后停止后续步骤；第 1 步已产生结果（表: {0}），用户可按已验证范围恢复任务目标工作簿（快照 ID: {1}）。",
                        r1.outputRef != null ? r1.outputRef.sheetName : "原表", record.snapshotId);
                    return record;
                }

                // 9. 双步均成功
                record.status = "completed";
                record.phase = "completed";
                record.completedSteps = 2;
                record.recoveryNotice = string.Format("双步骤流水线已全部执行完成。若需撤销全部修改，用户可按已验证范围恢复任务目标工作簿（快照 ID: {0}）。", record.snapshotId);
                return record;
            }
            finally
            {
                swTotal.Stop();
                record.totalElapsedMs = swTotal.ElapsedMilliseconds;
                SaveRunRecord(record);
                lock (_runLock)
                {
                    _isRunning = false;
                    _cancelRequested = false;
                }
            }
        }

        private static string ValidateMacroIntegrity(WorkflowStepDefinition step)
        {
            if (step.macroParams == null || string.IsNullOrEmpty(step.macroParams.scriptId))
            {
                return "未指定已保存宏 ID (scriptId)。";
            }
            var script = ScriptManager.GetScriptById(step.macroParams.scriptId);
            if (script == null)
            {
                return "找不到 ID 为 " + step.macroParams.scriptId + " 的宏脚本。";
            }
            string currentHash = VbaRunner.ComputeSha256(script.code ?? "");
            if (!string.IsNullOrEmpty(step.macroParams.macroSha256))
            {
                if (!string.Equals(currentHash, step.macroParams.macroSha256, StringComparison.OrdinalIgnoreCase))
                {
                    return string.Format("宏代码已发生变更（定义固化哈希: {0} vs 当前实际哈希: {1}）。系统拒绝隐式使用最新版本，请重新在工作流中确认并保存新版本。",
                        step.macroParams.macroSha256, currentHash);
                }
            }
            return null;
        }

        private static StepRunResult ExecuteStep(
            WorkflowStepDefinition step,
            dynamic app,
            dynamic targetWb,
            StepOutputReference prevOutput)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var res = new StepRunResult
            {
                stepIndex = step.stepIndex,
                stepName = step.stepName,
                toolType = step.toolType,
                status = "failed"
            };

            try
            {
                if (step.toolType == "dedup")
                {
                    var p = step.dedupParams ?? new DedupStepParams();
                    string targetSheet = p.targetSheet;
                    string targetAddress = p.rangeAddress;

                    // 若第 2 步为去重且消费前一步输出
                    if (step.inputSource == "prev_step_output" && prevOutput != null)
                    {
                        targetSheet = prevOutput.sheetName;
                        targetAddress = prevOutput.rangeAddress;
                    }

                    if (string.IsNullOrEmpty(targetSheet))
                    {
                        res.status = "failed";
                        res.failureStage = "missing_sheet";
                        res.error = "未指定待去重工作表名称，严禁隐式回退当前活动工作表。";
                        res.summaryText = res.error;
                        return res;
                    }

                    var dedupExec = DataToolsService.ExecuteDedup(
                        app,
                        targetWb,
                        targetSheet,
                        targetAddress,
                        p.hasHeader,
                        p.keyColumns,
                        p.mode ?? "export_unique",
                        p.expectedFingerprint
                    );

                    if (!dedupExec.ok)
                    {
                        res.status = "failed";
                        res.failureStage = "dedup_failed";
                        res.error = dedupExec.error;
                        res.summaryText = "去重执行失败: " + dedupExec.error;
                        return res;
                    }

                    string outSheet = (p.mode == "export_unique") ? dedupExec.resultSheetName : targetSheet;
                    string outRange = "";
                    try
                    {
                        dynamic wsOut = targetWb.Worksheets[outSheet];
                        outRange = ((string)wsOut.UsedRange.Address).Replace("$", "");
                    }
                    catch { }
                    if (string.IsNullOrEmpty(outRange))
                    {
                        outRange = (p.mode == "export_unique") ? ("A1:C" + (p.hasHeader ? (dedupExec.uniqueCount + 1) : dedupExec.uniqueCount)) : dedupExec.highlightedRange;
                    }

                    res.status = "success";
                    res.summaryText = dedupExec.message;
                    res.outputRef = new StepOutputReference
                    {
                        targetWorkbookName = (string)targetWb.Name,
                        sheetName = outSheet,
                        rangeAddress = outRange,
                        hasHeader = p.hasHeader,
                        headerRowIndex = 1,
                        rowCount = (p.mode == "export_unique") ? (dedupExec.uniqueCount + (p.hasHeader ? 1 : 0)) : dedupExec.dataRowCount,
                        columnCount = p.keyColumns != null ? p.keyColumns.Count : 1
                    };
                    return res;
                }
                else if (step.toolType == "reconcile")
                {
                    var p = step.reconcileParams ?? new ReconcileStepParams();
                    string leftSheet = p.leftSheet;
                    string leftAddress = p.leftAddress;
                    bool leftHasHeader = p.leftHasHeader;

                    // 若消费前序输出
                    if (step.inputSource == "prev_step_output" && prevOutput != null)
                    {
                        leftSheet = prevOutput.sheetName;
                        leftAddress = prevOutput.rangeAddress;
                        leftHasHeader = prevOutput.hasHeader;
                    }

                    if (string.IsNullOrEmpty(leftAddress) && !string.IsNullOrEmpty(leftSheet))
                    {
                        try
                        {
                            dynamic wsLeft = targetWb.Worksheets[leftSheet];
                            leftAddress = ((string)wsLeft.UsedRange.Address).Replace("$", "");
                        }
                        catch { }
                    }

                    // 必须核验消费工作表确实存在于目标工作簿
                    bool leftSheetFound = false;
                    try
                    {
                        foreach (dynamic ws in targetWb.Worksheets)
                        {
                            if (string.Equals((string)ws.Name, leftSheet, StringComparison.OrdinalIgnoreCase))
                            {
                                leftSheetFound = true;
                                break;
                            }
                        }
                    }
                    catch { }

                    if (!leftSheetFound)
                    {
                        res.status = "blocked";
                        res.failureStage = "output_verification_failed";
                        res.error = string.Format("无法核验前序输出工作表【{0}】，目标工作簿中未找到该表，第二步已安全阻断。", leftSheet ?? "空");
                        res.summaryText = res.error;
                        return res;
                    }

                    string targetWbName = (string)targetWb.Name;
                    var recExec = DataToolsService.ExecuteReconcile(
                        app,
                        targetWbName,
                        leftSheet,
                        leftAddress,
                        leftHasHeader,
                        p.leftKeyCols,
                        targetWbName,
                        p.rightSheet,
                        p.rightAddress,
                        p.rightHasHeader,
                        p.rightKeyCols,
                        p.compareCols,
                        targetWbName,
                        p.expectedFingerprint
                    );

                    if (!recExec.ok)
                    {
                        res.status = "failed";
                        res.failureStage = "reconcile_failed";
                        res.error = recExec.error;
                        res.summaryText = "对账执行失败: " + recExec.error;
                        return res;
                    }

                    res.status = "success";
                    res.summaryText = recExec.message;
                    res.outputRef = new StepOutputReference
                    {
                        targetWorkbookName = (string)targetWb.Name,
                        sheetName = recExec.resultSheetName,
                        rangeAddress = "A1:" + recExec.exportedRowCount,
                        hasHeader = true,
                        headerRowIndex = 1,
                        rowCount = recExec.exportedRowCount,
                        columnCount = (p.compareCols != null ? p.compareCols.Count : 1) + 2
                    };
                    return res;
                }
                else if (step.toolType == "saved_macro")
                {
                    var p = step.macroParams ?? new SavedMacroStepParams();
                    var script = ScriptManager.GetScriptById(p.scriptId);
                    if (script == null)
                    {
                        res.status = "failed";
                        res.failureStage = "script_not_found";
                        res.error = "找不到宏脚本: " + p.scriptId;
                        res.summaryText = res.error;
                        return res;
                    }

                    // 准备参数
                    var runParams = p.parameters != null ? new List<VbaParameterInput>(p.parameters) : new List<VbaParameterInput>();

                    // 若声明需要前序输出并包含输出目标表
                    if (step.inputSource == "prev_step_output" && prevOutput != null && !string.IsNullOrEmpty(prevOutput.sheetName))
                    {
                        // 若宏参数包含 Sheet 或 Range，尝试自动补齐或核验
                        foreach (var rp in runParams)
                        {
                            if (rp.name.IndexOf("sheet", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                rp.value = prevOutput.sheetName;
                            }
                        }
                    }

                    string ep = !string.IsNullOrEmpty(p.entryPoint) ? p.entryPoint : script.entryPoint;
                    var execResult = VbaRunner.RunVbaCode(
                        app,
                        targetWb,
                        script.code,
                        "工作流调用宏: " + (script.displayName ?? script.name),
                        ep,
                        runParams
                    );

                    if (execResult == null || !execResult.success)
                    {
                        res.status = "failed";
                        res.failureStage = execResult != null ? execResult.failureStage : "execution_failed";
                        res.error = execResult != null ? (execResult.vbaErrDescription ?? execResult.summary) : "宏执行失败";
                        res.summaryText = res.error;
                        return res;
                    }

                    // 校验用户声明的输出
                    if (!string.IsNullOrEmpty(p.declaredOutputSheet))
                    {
                        bool sheetFound = false;
                        try
                        {
                            foreach (dynamic ws in targetWb.Worksheets)
                            {
                                if (string.Equals((string)ws.Name, p.declaredOutputSheet, StringComparison.OrdinalIgnoreCase))
                                {
                                    sheetFound = true;
                                    break;
                                }
                            }
                        }
                        catch { }

                        if (!sheetFound)
                        {
                            res.status = "failed";
                            res.failureStage = "declared_output_missing";
                            res.error = string.Format("宏执行成功但未能在目标工作簿中检测到声明的输出工作表【{0}】，步骤核验失败。", p.declaredOutputSheet);
                            res.summaryText = res.error;
                            return res;
                        }
                    }

                    res.status = "success";
                    res.summaryText = string.Format("宏【{0}】执行完成。耗时 {1}ms", script.displayName ?? script.name, execResult.elapsedMs);
                    res.outputRef = new StepOutputReference
                    {
                        targetWorkbookName = (string)targetWb.Name,
                        sheetName = !string.IsNullOrEmpty(p.declaredOutputSheet) ? p.declaredOutputSheet : ((string)targetWb.ActiveSheet.Name),
                        rangeAddress = !string.IsNullOrEmpty(p.declaredOutputRange) ? p.declaredOutputRange : "A1",
                        hasHeader = true,
                        headerRowIndex = 1,
                        rowCount = 1,
                        columnCount = 1
                    };
                    return res;
                }
                else if (step.toolType == "chart")
                {
                    var p = step.chartParams ?? new ChartStepParams();
                    string srcSheet = p.sourceSheet;
                    string srcRange = p.sourceRange;
                    bool hasHeaders = p.hasHeaders;

                    if (step.inputSource == "prev_step_output")
                    {
                        if (prevOutput == null || string.IsNullOrEmpty(prevOutput.sheetName) || string.IsNullOrEmpty(prevOutput.rangeAddress))
                        {
                            res.status = "failed";
                            res.failureStage = "missing_input_source";
                            res.error = "步骤配置为消费前序输出，但前置步骤未产生有效的结构化输出。";
                            res.summaryText = res.error;
                            return res;
                        }
                        srcSheet = prevOutput.sheetName;
                        srcRange = prevOutput.rangeAddress;
                        hasHeaders = prevOutput.hasHeader;
                    }

                    if (string.IsNullOrEmpty(srcSheet))
                    {
                        res.status = "failed";
                        res.failureStage = "missing_sheet";
                        res.error = "未指定图表数据源工作表 (sourceSheet)。系统严格拒绝隐式回退当前活动工作表！";
                        res.summaryText = res.error;
                        return res;
                    }

                    if (string.IsNullOrEmpty(srcRange))
                    {
                        res.status = "failed";
                        res.failureStage = "missing_range";
                        res.error = "未指定图表数据源区域 (sourceRange)。系统严格拒绝隐式猜测！";
                        res.summaryText = res.error;
                        return res;
                    }

                    // 必须指定有效的类别列和至少一个数值列
                    if (p.categoryColIndex < 1)
                    {
                        res.status = "failed";
                        res.failureStage = "invalid_category_col";
                        res.error = "图表必须显式指定有效的类别 (X轴) 列索引 (categoryColIndex)。";
                        res.summaryText = res.error;
                        return res;
                    }

                    if (p.seriesColIndices == null || p.seriesColIndices.Count == 0)
                    {
                        res.status = "failed";
                        res.failureStage = "invalid_series_cols";
                        res.error = "图表必须显式指定至少一个数值系列列索引 (seriesColIndices)。系统拒绝全列盲目绘制！";
                        res.summaryText = res.error;
                        return res;
                    }

                    var qcP = new QuickChartParams
                    {
                        targetWorkbookName = (string)targetWb.Name,
                        targetWorkbookFullName = (string)targetWb.FullName,
                        sourceSheet = srcSheet,
                        sourceRange = srcRange,
                        hasHeaders = hasHeaders,
                        dataStartRow = p.dataStartRow,
                        dataEndRow = p.dataEndRow,
                        categoryColIndex = p.categoryColIndex,
                        categoryColName = p.categoryColName,
                        seriesColIndices = p.seriesColIndices,
                        seriesNames = p.seriesNames,
                        chartType = p.chartType,
                        title = p.title,
                        targetSheet = !string.IsNullOrEmpty(p.targetSheet) ? p.targetSheet : srcSheet,
                        placementMode = p.placementMode,
                        targetCell = p.targetCell,
                        left = p.left,
                        top = p.top,
                        width = p.width,
                        height = p.height,
                        action = p.action,
                        targetChartId = p.targetChartId,
                        errorHandling = p.errorHandling
                    };

                    var qcRes = ChartService.ExecuteQuickChart(app, qcP, skipSnapshot: true, existingSnapshotId: null);

                    if (!qcRes.ok)
                    {
                        res.status = "failed";
                        res.failureStage = qcRes.failureStage;
                        res.error = qcRes.error;
                        res.summaryText = qcRes.error;
                        return res;
                    }

                    res.status = "success";
                    res.summaryText = qcRes.summaryText;
                    res.outputRef = new StepOutputReference
                    {
                        targetWorkbookName = (string)targetWb.Name,
                        sheetName = qcP.targetSheet,
                        rangeAddress = !string.IsNullOrEmpty(qcP.targetCell) ? qcP.targetCell : ("Chart:" + qcRes.chartName),
                        hasHeader = false,
                        headerRowIndex = 0,
                        rowCount = 1,
                        columnCount = 1
                    };
                    return res;
                }
                else
                {
                    res.status = "failed";
                    res.failureStage = "unsupported_tool";
                    res.error = "不支持的步骤工具类型: " + step.toolType;
                    res.summaryText = res.error;
                    return res;
                }
            }
            catch (Exception ex)
            {
                res.status = "failed";
                res.failureStage = "unexpected_exception";
                res.error = ex.Message;
                res.summaryText = "步骤执行发生异常: " + ex.Message;
                return res;
            }
            finally
            {
                sw.Stop();
                res.elapsedMs = sw.ElapsedMilliseconds;
            }
        }

        private static void SaveRunRecord(WorkflowRunRecord record)
        {
            if (record == null || string.IsNullOrEmpty(record.workflowId)) return;
            EnsureDirectories();
            lock (_fileLock)
            {
                try
                {
                    string path = Path.Combine(RunsDir, record.workflowId + "_runs.json");
                    var list = new List<WorkflowRunRecord>();
                    if (File.Exists(path))
                    {
                        string json = File.ReadAllText(path, Encoding.UTF8);
                        list = SimpleJson.DeserializeList<WorkflowRunRecord>(json) ?? new List<WorkflowRunRecord>();
                    }

                    // 限制最大保留 20 条执行记录
                    list.Insert(0, record);
                    if (list.Count > 20)
                    {
                        list.RemoveRange(20, list.Count - 20);
                    }

                    File.WriteAllText(path, SimpleJson.Serialize(list), Encoding.UTF8);
                }
                catch { }
            }
        }
    }
}
