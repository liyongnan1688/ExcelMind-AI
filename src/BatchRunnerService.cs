using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace LeeExcel
{
    /// <summary>
    /// 单个文件的来源身份信息与固化快照指纹
    /// </summary>
    public class BatchFileItem
    {
        public string originalFilePath { get; set; }
        public string originalFileName { get; set; }
        public string originalFileHash { get; set; }
        public long originalFileSize { get; set; }
        public string originalLastModified { get; set; }
    }

    /// <summary>
    /// 批量宏任务队列固化定义
    /// </summary>
    public class BatchJobDefinition
    {
        public string jobId { get; set; }
        public string macroCode { get; set; }
        public string macroHash { get; set; }
        public string entryPoint { get; set; }
        public string parametersJson { get; set; }
        public string outputDir { get; set; }
        public bool stopOnError { get; set; }
        public List<BatchFileItem> files { get; set; }
        public string riskNotice { get; set; }

        public BatchJobDefinition()
        {
            files = new List<BatchFileItem>();
            stopOnError = true; // 默认遇错停止后续任务
            riskNotice = BatchRunnerService.FIXED_REFERENCE_RISK_NOTICE;
        }
    }

    /// <summary>
    /// 单个文件的执行结果记录（六态 + failureStage + 读回）
    /// </summary>
    public class BatchFileTaskResult
    {
        public int fileIndex { get; set; }
        public string originalFilePath { get; set; }
        public string originalFileName { get; set; }
        public string workingCopyPath { get; set; }
        public string snapshotPath { get; set; }
        public string snapshotId { get; set; }
        public string finalOutputPath { get; set; }
        
        /// <summary>
        /// 六态之一：pending, running, success, failed, blocked, cancelled
        /// </summary>
        public string status { get; set; }
        
        /// <summary>
        /// 失败阶段：file_validation, macro_hash_mismatch, copy_creation, file_protected_or_corrupted,
        /// target_binding, precheck, snapshot, execution, readback, save_output, meta_record
        /// </summary>
        public string failureStage { get; set; }
        
        public string error { get; set; }
        public double elapsedMs { get; set; }
        public WorkbookReadback readback { get; set; }
        public string executionSummary { get; set; }
        public bool isPartiallyModified { get; set; }
        public bool temporaryCleaned { get; set; }
    }

    /// <summary>
    /// 批量任务总体汇总摘要
    /// </summary>
    public class BatchJobSummary
    {
        public string jobId { get; set; }
        public string status { get; set; } // pending, running, completed, stopped_on_error, cancelled
        public int totalFiles { get; set; }
        public int successCount { get; set; }
        public int failedCount { get; set; }
        public int blockedCount { get; set; }
        public int cancelledCount { get; set; }
        public int pendingCount { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string cancelRequestedAt { get; set; }
        public string riskNotice { get; set; }
        public List<BatchFileTaskResult> fileResults { get; set; }

        public BatchJobSummary()
        {
            fileResults = new List<BatchFileTaskResult>();
            riskNotice = BatchRunnerService.FIXED_REFERENCE_RISK_NOTICE;
        }
    }

    /// <summary>
    /// 批量宏任务队列调度服务
    /// 遵循最小切片原则：单任务、单实例、严格串行调度，原文件只读，隔离工作副本执行，同名冲突递增重命名，遇错即停
    /// </summary>
    public class BatchRunnerService
    {
        public const string FIXED_REFERENCE_RISK_NOTICE =
            "【重要安全提示与执行边界说明】\n" +
            "1. 本任务调度与保存机制严格保证：原文件仅作为复制来源绝对只读；执行过程完全在隔离工作副本中进行，不保存或修改原文件及用户已打开的外部工作簿；\n" +
            "2. 隔离工作副本、快照及独立 Excel 实例不能作为任意 VBA 代码的安全沙箱；\n" +
            "3. 若宏代码中包含硬编码的固定绝对路径（如 Workbooks.Open 外部文件）、外部文件操作（如 Kill、FileSystemObject 写入）或其他系统级副作用，此类操作仍会直接作用于系统外部环境；\n" +
            "4. 系统不会改写宏代码来消除固定引用，请在执行前确认宏代码逻辑不包含未授权破坏性操作。";

        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx", ".xlsm", ".xlsb", ".xls"
        };

        private static readonly object _stateLock = new object();
        private static readonly Dictionary<string, BatchJobDefinition> _lockedJobs = new Dictionary<string, BatchJobDefinition>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, BatchJobSummary> _jobHistory = new Dictionary<string, BatchJobSummary>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> _cancelRequests = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static string _currentRunningJobId = null;

        /// <summary>
        /// 获取已固化的批量任务定义
        /// </summary>
        public static BatchJobDefinition GetLockedJob(string jobId)
        {
            if (string.IsNullOrEmpty(jobId)) return null;
            lock (_stateLock)
            {
                if (_lockedJobs.ContainsKey(jobId))
                {
                    return _lockedJobs[jobId];
                }
                return null;
            }
        }

        /// <summary>
        /// 计算文件的 SHA256 哈希
        /// </summary>
        public static string ComputeFileSha256(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return string.Empty;
            try
            {
                using (var sha = SHA256.Create())
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    byte[] bytes = sha.ComputeHash(fs);
                    var sb = new StringBuilder();
                    foreach (byte b in bytes) sb.Append(b.ToString("x2"));
                    return sb.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 计算字符串的 SHA256 哈希
        /// </summary>
        public static string ComputeStringSha256(string text)
        {
            if (text == null) text = string.Empty;
            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
                var sb = new StringBuilder();
                foreach (byte b in bytes) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        /// <summary>
        /// 验证并固化批量任务队列定义（执行前固定：文件清单、来源身份、宏哈希、输出规则）
        /// </summary>
        public static BatchJobDefinition ValidateAndLockJob(
            List<string> rawFilePaths,
            string macroCode,
            string entryPoint,
            string parametersJson,
            string outputDir,
            bool stopOnError,
            out string validationError)
        {
            validationError = null;

            if (rawFilePaths == null || rawFilePaths.Count == 0)
            {
                validationError = "文件清单为空，请至少选择一个有效的 Excel 文件。";
                return null;
            }

            if (string.IsNullOrWhiteSpace(macroCode))
            {
                validationError = "宏代码内容为空，无法生成批量执行队列。";
                return null;
            }

            if (string.IsNullOrWhiteSpace(outputDir))
            {
                validationError = "输出目录未指定，必须显式指定独立的输出目录。";
                return null;
            }

            try
            {
                outputDir = Path.GetFullPath(outputDir);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                validationError = "输出目录无效或无权创建: " + ex.Message;
                return null;
            }

            string macroHash = ComputeStringSha256(macroCode);
            string jobId = "batch_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);

            var jobDef = new BatchJobDefinition
            {
                jobId = jobId,
                macroCode = macroCode,
                macroHash = macroHash,
                entryPoint = entryPoint,
                parametersJson = parametersJson,
                outputDir = outputDir,
                stopOnError = stopOnError,
                riskNotice = FIXED_REFERENCE_RISK_NOTICE
            };

            for (int i = 0; i < rawFilePaths.Count; i++)
            {
                string p = rawFilePaths[i];
                if (string.IsNullOrWhiteSpace(p))
                {
                    validationError = string.Format("第 {0} 个文件路径为空。", i + 1);
                    return null;
                }

                string fullPath = p;
                try
                {
                    fullPath = Path.GetFullPath(p);
                }
                catch (Exception exP)
                {
                    validationError = string.Format("第 {0} 个文件路径格式无效: {1} ({2})", i + 1, p, exP.Message);
                    return null;
                }

                if (!File.Exists(fullPath))
                {
                    validationError = string.Format("文件不存在: {0}", fullPath);
                    return null;
                }

                string ext = Path.GetExtension(fullPath);
                if (!SupportedExtensions.Contains(ext))
                {
                    validationError = string.Format("不支持的文件格式 '{0}' ({1})。当前仅支持 .xlsx, .xlsm, .xlsb, .xls。", ext, Path.GetFileName(fullPath));
                    return null;
                }

                // 输出路径与原文件路径重合检查（输出目录不能是原文件直接覆盖）
                string targetInOutputDir = Path.Combine(outputDir, Path.GetFileName(fullPath));
                if (string.Equals(fullPath, targetInOutputDir, StringComparison.OrdinalIgnoreCase))
                {
                    // 输出目录与原文件同目录时，必须通过同名递增机制保护原文件不被覆盖，此处给出显式保障标记
                }

                var fi = new FileInfo(fullPath);
                string fileHash = ComputeFileSha256(fullPath);

                jobDef.files.Add(new BatchFileItem
                {
                    originalFilePath = fullPath,
                    originalFileName = fi.Name,
                    originalFileHash = fileHash,
                    originalFileSize = fi.Length,
                    originalLastModified = fi.LastWriteTime.ToString("o")
                });
            }

            lock (_stateLock)
            {
                _lockedJobs[jobId] = jobDef;
            }

            return jobDef;
        }

        /// <summary>
        /// 请求取消当前运行中的批量任务（在任务边界安全生效，绝不强杀宿主）
        /// </summary>
        public static bool RequestCancel(string jobId, out string message)
        {
            lock (_stateLock)
            {
                if (string.IsNullOrEmpty(jobId))
                {
                    message = "未指定任务 ID";
                    return false;
                }

                _cancelRequests.Add(jobId);
                if (_jobHistory.ContainsKey(jobId))
                {
                    _jobHistory[jobId].cancelRequestedAt = DateTime.Now.ToString("o");
                }

                message = "已收到取消请求。当前正在处理的文件将在任务边界安全结束，后续待处理文件将置为已取消并停止调度。";
                return true;
            }
        }

        /// <summary>
        /// 检查指定任务是否已被请求取消
        /// </summary>
        public static bool IsCancelRequested(string jobId)
        {
            if (string.IsNullOrEmpty(jobId)) return false;
            lock (_stateLock)
            {
                return _cancelRequests.Contains(jobId);
            }
        }

        /// <summary>
        /// 生成 BatchJobSummary 的线程安全只读快照
        /// </summary>
        private static BatchJobSummary CloneSummarySnapshot(BatchJobSummary orig)
        {
            if (orig == null) return null;
            var clone = new BatchJobSummary
            {
                jobId = orig.jobId,
                status = orig.status,
                totalFiles = orig.totalFiles,
                successCount = orig.successCount,
                failedCount = orig.failedCount,
                blockedCount = orig.blockedCount,
                cancelledCount = orig.cancelledCount,
                pendingCount = orig.pendingCount,
                startTime = orig.startTime,
                endTime = orig.endTime,
                cancelRequestedAt = orig.cancelRequestedAt,
                riskNotice = orig.riskNotice
            };
            if (orig.fileResults != null)
            {
                foreach (var f in orig.fileResults)
                {
                    clone.fileResults.Add(new BatchFileTaskResult
                    {
                        fileIndex = f.fileIndex,
                        originalFilePath = f.originalFilePath,
                        originalFileName = f.originalFileName,
                        workingCopyPath = f.workingCopyPath,
                        snapshotPath = f.snapshotPath,
                        snapshotId = f.snapshotId,
                        finalOutputPath = f.finalOutputPath,
                        status = f.status,
                        failureStage = f.failureStage,
                        error = f.error,
                        elapsedMs = f.elapsedMs,
                        executionSummary = f.executionSummary,
                        isPartiallyModified = f.isPartiallyModified,
                        temporaryCleaned = f.temporaryCleaned,
                        readback = f.readback
                    });
                }
            }
            return clone;
        }

        /// <summary>
        /// 获取批量任务最新运行摘要（只读线程安全快照，绝不调用 Excel COM）
        /// </summary>
        public static BatchJobSummary GetJobSummary(string jobId)
        {
            lock (_stateLock)
            {
                if (string.IsNullOrEmpty(jobId))
                {
                    if (!string.IsNullOrEmpty(_currentRunningJobId) && _jobHistory.ContainsKey(_currentRunningJobId))
                    {
                        return CloneSummarySnapshot(_jobHistory[_currentRunningJobId]);
                    }
                    return null;
                }
                if (_jobHistory.ContainsKey(jobId))
                {
                    return CloneSummarySnapshot(_jobHistory[jobId]);
                }
                return null;
            }
        }

        /// <summary>
        /// 在受控独立 STA 工作线程上异步执行批量任务，立即向调用方返回初始状态
        /// </summary>
        public static BatchJobSummary StartBatchJobAsync(BatchJobDefinition jobDef, dynamic app = null)
        {
            if (jobDef == null || jobDef.files == null || jobDef.files.Count == 0)
            {
                throw new ArgumentException("批量任务定义为空或无文件。");
            }

            lock (_stateLock)
            {
                // 1. 检查单实例并发限制
                if (!string.IsNullOrEmpty(_currentRunningJobId) && !string.Equals(_currentRunningJobId, jobDef.jobId, StringComparison.OrdinalIgnoreCase))
                {
                    if (_jobHistory.ContainsKey(_currentRunningJobId) && _jobHistory[_currentRunningJobId].status == "running")
                    {
                        throw new InvalidOperationException(string.Format("已有批量任务正在执行中 (jobId={0})，单实例模式下严禁并发启动新任务。", _currentRunningJobId));
                    }
                }

                // 2. 检查同一 jobId 防重复执行
                if (_jobHistory.ContainsKey(jobDef.jobId))
                {
                    var existing = _jobHistory[jobDef.jobId];
                    if (existing.status == "running")
                    {
                        throw new InvalidOperationException(string.Format("任务 {0} 正在执行中，严禁重复启动。", jobDef.jobId));
                    }
                    if (existing.status == "completed" || existing.status == "stopped_on_error" || existing.status == "cancelled")
                    {
                        throw new InvalidOperationException(string.Format("任务 {0} 已处于终端状态 ({1})，严禁再次启动重复执行已处理文件。如需再次运行必须重新固化新任务并确认。", jobDef.jobId, existing.status));
                    }
                }

                // 预置初始运行状态摘要
                var initSummary = new BatchJobSummary
                {
                    jobId = jobDef.jobId,
                    status = "running",
                    totalFiles = jobDef.files.Count,
                    pendingCount = jobDef.files.Count,
                    startTime = DateTime.Now.ToString("o"),
                    riskNotice = FIXED_REFERENCE_RISK_NOTICE
                };
                for (int i = 0; i < jobDef.files.Count; i++)
                {
                    initSummary.fileResults.Add(new BatchFileTaskResult
                    {
                        fileIndex = i,
                        originalFilePath = jobDef.files[i].originalFilePath,
                        originalFileName = jobDef.files[i].originalFileName,
                        status = "pending",
                        failureStage = null
                    });
                }

                _currentRunningJobId = jobDef.jobId;
                _jobHistory[jobDef.jobId] = initSummary;
            }

            // 启动受控 STA 工作线程
            Thread worker = new Thread(() =>
            {
                try
                {
                    ExecuteBatchJob(jobDef, app);
                }
                catch (Exception)
                {
                    lock (_stateLock)
                    {
                        if (_jobHistory.ContainsKey(jobDef.jobId))
                        {
                            var s = _jobHistory[jobDef.jobId];
                            s.status = "stopped_on_error";
                            s.endTime = DateTime.Now.ToString("o");
                        }
                        _currentRunningJobId = null;
                    }
                }
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.IsBackground = true;
            worker.Start();

            return GetJobSummary(jobDef.jobId);
        }

        /// <summary>
        /// 解析最终唯一的输出文件路径，若存在同名文件则自动递增 _1, _2，绝不覆盖已有文件。
        /// 限定为 internal，避免被 Excel-DNA 自动导出为 Excel 工作表函数产生重复注册冲突。
        /// </summary>
        internal static string ResolveUniqueOutputPath(string outputDir, string originalFileName)
        {
            string baseName = Path.GetFileNameWithoutExtension(originalFileName);
            string ext = Path.GetExtension(originalFileName);

            string candidate = Path.Combine(outputDir, originalFileName);
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            int index = 1;
            while (true)
            {
                string newName = string.Format("{0}_{1}{2}", baseName, index, ext);
                candidate = Path.Combine(outputDir, newName);
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
                index++;
            }
        }

        /// <summary>
        /// 核心批量宏执行流水线（串行、单实例、独立对象管理、保护用户已有工作簿）
        /// </summary>
        public static BatchJobSummary ExecuteBatchJob(BatchJobDefinition jobDef, dynamic app)
        {
            if (jobDef == null || jobDef.files == null || jobDef.files.Count == 0)
            {
                throw new ArgumentException("批量任务定义为空或无文件。");
            }

            lock (_stateLock)
            {
                // 1. 检查单实例并发限制
                if (!string.IsNullOrEmpty(_currentRunningJobId) && !string.Equals(_currentRunningJobId, jobDef.jobId, StringComparison.OrdinalIgnoreCase))
                {
                    if (_jobHistory.ContainsKey(_currentRunningJobId) && _jobHistory[_currentRunningJobId].status == "running")
                    {
                        throw new InvalidOperationException(string.Format("已有批量任务正在执行中 (jobId={0})，单实例模式下严禁并发启动新任务。", _currentRunningJobId));
                    }
                }

                // 2. 检查同一 jobId 防重复执行
                if (_jobHistory.ContainsKey(jobDef.jobId))
                {
                    var existing = _jobHistory[jobDef.jobId];
                    if (existing.status == "completed" || existing.status == "stopped_on_error" || existing.status == "cancelled")
                    {
                        throw new InvalidOperationException(string.Format("任务 {0} 已处于终端状态 ({1})，严禁再次启动重复执行已处理文件。如需再次运行必须重新固化新任务并确认。", jobDef.jobId, existing.status));
                    }
                }
            }

            bool createdOwnExcel = false;
            dynamic runningApp = app;
            if (runningApp == null)
            {
                try
                {
                    Type excelType = Type.GetTypeFromProgID("Excel.Application");
                    runningApp = Activator.CreateInstance(excelType);
                    runningApp.Visible = false;
                    runningApp.DisplayAlerts = false;
                    runningApp.AskToUpdateLinks = false;
                    createdOwnExcel = true;
                }
                catch { }
            }

            var summary = new BatchJobSummary
            {
                jobId = jobDef.jobId,
                status = "running",
                totalFiles = jobDef.files.Count,
                startTime = DateTime.Now.ToString("o"),
                riskNotice = FIXED_REFERENCE_RISK_NOTICE
            };

            // 初始化所有文件的结果对象（初始状态为 pending）
            for (int i = 0; i < jobDef.files.Count; i++)
            {
                var fileItem = jobDef.files[i];
                summary.fileResults.Add(new BatchFileTaskResult
                {
                    fileIndex = i,
                    originalFilePath = fileItem.originalFilePath,
                    originalFileName = fileItem.originalFileName,
                    status = "pending",
                    failureStage = null
                });
            }

            lock (_stateLock)
            {
                _currentRunningJobId = jobDef.jobId;
                _jobHistory[jobDef.jobId] = summary;
            }

            // 专用隔离工作副本临时目录（按任务 ID 隔离）
            string baseWorkDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelMindAI",
                "BatchWorkCopies",
                jobDef.jobId
            );

            try
            {
                if (!Directory.Exists(baseWorkDir))
                {
                    Directory.CreateDirectory(baseWorkDir);
                }
            }
            catch (Exception exDir)
            {
                summary.status = "stopped_on_error";
                summary.endTime = DateTime.Now.ToString("o");
                for (int i = 0; i < summary.fileResults.Count; i++)
                {
                    summary.fileResults[i].status = "blocked";
                    summary.fileResults[i].failureStage = "copy_creation";
                    summary.fileResults[i].error = "无法创建隔离工作副本临时目录: " + exDir.Message;
                }
                summary.blockedCount = summary.fileResults.Count;
                return summary;
            }

            // 解析结构化参数（复用 R2c 契约）
            List<VbaParameterInput> parsedParams = null;
            if (!string.IsNullOrEmpty(jobDef.parametersJson))
            {
                string pRaw = jobDef.parametersJson.Trim();
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

            // 保存并临时调整宿主交互静默状态（防止密码、链接更新等交互弹窗导致队列无限挂死）
            bool origDisplayAlerts = true;
            bool origAskToUpdateLinks = true;
            try
            {
                if (app != null)
                {
                    origDisplayAlerts = (bool)app.DisplayAlerts;
                    origAskToUpdateLinks = (bool)app.AskToUpdateLinks;
                    app.DisplayAlerts = false;
                    app.AskToUpdateLinks = false;
                }
            }
            catch { }

            bool stopTriggered = false;

            try
            {
                // 逐文件严格串行流水线
                for (int i = 0; i < jobDef.files.Count; i++)
                {
                    var fileItem = jobDef.files[i];
                    var taskResult = summary.fileResults[i];

                    // 1. 任务边界取消检查
                    if (IsCancelRequested(jobDef.jobId))
                    {
                        taskResult.status = "cancelled";
                        taskResult.failureStage = null;
                        taskResult.error = "任务已根据用户请求安全取消。";
                        summary.cancelledCount++;
                        continue;
                    }

                    if (stopTriggered)
                    {
                        // 前置文件失败触发了遇错即停（Stop on Error），后续未启动项保持 pending
                        continue;
                    }

                    taskResult.status = "running";
                    var sw = System.Diagnostics.Stopwatch.StartNew();

                    // 执行逐文件安全流水线
                    ExecuteSingleFileTask(jobDef, fileItem, taskResult, baseWorkDir, runningApp, parsedParams);

                    sw.Stop();
                    taskResult.elapsedMs = sw.ElapsedMilliseconds;

                    // 状态分类统计
                    if (taskResult.status == "success")
                    {
                        summary.successCount++;
                    }
                    else if (taskResult.status == "failed")
                    {
                        summary.failedCount++;
                        if (jobDef.stopOnError)
                        {
                            stopTriggered = true;
                        }
                    }
                    else if (taskResult.status == "blocked")
                    {
                        summary.blockedCount++;
                        if (jobDef.stopOnError)
                        {
                            stopTriggered = true;
                        }
                    }
                    else if (taskResult.status == "cancelled")
                    {
                        summary.cancelledCount++;
                    }
                }
            }
            finally
            {
                // 恢复宿主原始状态
                try
                {
                    if (runningApp != null)
                    {
                        runningApp.DisplayAlerts = origDisplayAlerts;
                        runningApp.AskToUpdateLinks = origAskToUpdateLinks;
                    }
                }
                catch { }

                // 最终统计 pending 数量
                int pendingLeft = 0;
                foreach (var r in summary.fileResults)
                {
                    if (r.status == "pending") pendingLeft++;
                }
                summary.pendingCount = pendingLeft;

                // 汇总状态判定
                if (IsCancelRequested(jobDef.jobId))
                {
                    summary.status = "cancelled";
                }
                else if (stopTriggered)
                {
                    summary.status = "stopped_on_error";
                }
                else
                {
                    summary.status = "completed";
                }

                summary.endTime = DateTime.Now.ToString("o");

                // 若是本任务自己创建的受控后台 Excel 实例，安全关闭并释放 COM 句柄
                if (createdOwnExcel && runningApp != null)
                {
                    try { runningApp.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(runningApp); } catch { }
                    runningApp = null;
                }

                lock (_stateLock)
                {
                    _currentRunningJobId = null;
                }
            }

            return summary;
        }

        /// <summary>
        /// 单个文件的完整隔离执行流水线（11步严格隔离操作）
        /// </summary>
        private static void ExecuteSingleFileTask(
            BatchJobDefinition jobDef,
            BatchFileItem fileItem,
            BatchFileTaskResult taskResult,
            string baseWorkDir,
            dynamic app,
            List<VbaParameterInput> parsedParams)
        {
            // 步骤 1：每项开始前再次严格校验（防版本漂移与静默篡改）
            if (!File.Exists(fileItem.originalFilePath))
            {
                taskResult.status = "blocked";
                taskResult.failureStage = "file_validation";
                taskResult.error = "原文件已不存在: " + fileItem.originalFilePath;
                return;
            }

            string currentFileHash = ComputeFileSha256(fileItem.originalFilePath);
            if (!string.Equals(currentFileHash, fileItem.originalFileHash, StringComparison.OrdinalIgnoreCase))
            {
                taskResult.status = "blocked";
                taskResult.failureStage = "file_hash_mismatch";
                taskResult.error = "检测到原文件自固化队列后已被外部修改（SHA-256 哈希发生漂移）。为防止静默使用新版本导致数据不一致，已严格阻断执行。";
                return;
            }

            string currentMacroHash = ComputeStringSha256(jobDef.macroCode);
            if (!string.Equals(currentMacroHash, jobDef.macroHash, StringComparison.OrdinalIgnoreCase))
            {
                taskResult.status = "blocked";
                taskResult.failureStage = "macro_hash_mismatch";
                taskResult.error = "检测到宏正文哈希与队列锁定时不一致。为防止静默使用非确认代码，已严格阻断执行。";
                return;
            }

            // 步骤 2：创建【隔离工作副本】（原文件保持只读，绝不直接写入或作为宏的目标）
            string workingCopyName = string.Format("work_{0}_{1}", taskResult.fileIndex, fileItem.originalFileName);
            string workingCopyPath = Path.Combine(baseWorkDir, workingCopyName);
            taskResult.workingCopyPath = workingCopyPath;

            try
            {
                File.Copy(fileItem.originalFilePath, workingCopyPath, true);
                // 确保工作副本去除了只读属性
                File.SetAttributes(workingCopyPath, FileAttributes.Normal);
            }
            catch (Exception exCopy)
            {
                taskResult.status = "failed";
                taskResult.failureStage = "copy_creation";
                taskResult.error = "创建隔离工作副本失败: " + exCopy.Message;
                return;
            }

            // 步骤 3：在 Excel 实例中以显式全路径打开【隔离工作副本】，严密限定归属
            dynamic targetWb = null;
            try
            {
                if (app == null)
                {
                    taskResult.status = "blocked";
                    taskResult.failureStage = "target_binding";
                    taskResult.error = "Excel Application 宿主对象无效，无法打开工作簿。";
                    return;
                }

                // 打开隔离副本（显式禁止链接更新）
                targetWb = app.Workbooks.Open(workingCopyPath, 0, false);
            }
            catch (Exception exOpen)
            {
                taskResult.status = "blocked";
                taskResult.failureStage = "file_protected_or_corrupted";
                taskResult.error = "打开工作簿失败（文件可能受密码保护、格式损坏或被排他锁定）: " + exOpen.Message;
                return;
            }

            if (targetWb == null)
            {
                taskResult.status = "blocked";
                taskResult.failureStage = "target_binding";
                taskResult.error = "无法获取工作簿对象实例，禁止使用 ActiveWorkbook 兜底。";
                return;
            }

            try
            {
                // 步骤 4：参数与入口签名严格预检（复用 R2c 契约与执行器）
                var precheck = VbaRunner.PrecheckParameters(app, targetWb, jobDef.macroCode, jobDef.entryPoint, parsedParams);
                if (!precheck.isOk)
                {
                    taskResult.status = "blocked";
                    taskResult.failureStage = precheck.failureStage ?? "precheck";
                    taskResult.error = "参数或签名预检失败: " + precheck.error;
                    return;
                }

                // 步骤 5：执行前创建整本物理快照（快照失败承诺零业务写入）
                SnapshotItem snap = null;
                try
                {
                    snap = SnapshotManager.CreateSnapshot(targetWb, "批量任务执行前快照: " + fileItem.originalFileName, jobDef.macroCode);
                }
                catch (Exception exSnap)
                {
                    taskResult.status = "blocked";
                    taskResult.failureStage = "snapshot";
                    taskResult.error = "执行前创建快照失败 (" + exSnap.Message + ")。为保障数据可回滚安全，绝对零写入阻断。";
                    return;
                }

                if (snap == null || string.IsNullOrEmpty(snap.id) || string.IsNullOrEmpty(snap.fileName))
                {
                    taskResult.status = "blocked";
                    taskResult.failureStage = "snapshot";
                    taskResult.error = "快照对象生成无效。为保障数据可回滚安全，绝对零写入阻断。";
                    return;
                }

                string snapFolder = SnapshotManager.GetBackupFolderForWorkbook(snap.originalPath ?? workingCopyPath);
                string snapFullPath = Path.Combine(snapFolder, snap.fileName);
                if (!File.Exists(snapFullPath) || new FileInfo(snapFullPath).Length == 0)
                {
                    taskResult.status = "blocked";
                    taskResult.failureStage = "snapshot";
                    taskResult.error = "快照物理文件验证失败（文件不存在或为0字节）。绝对零写入阻断。";
                    return;
                }

                taskResult.snapshotId = snap.id;
                taskResult.snapshotPath = snapFullPath;

                // 步骤 6：在隔离工作副本上执行宏
                var execResult = VbaRunner.RunVbaCode(app, targetWb, jobDef.macroCode, "batch_runner_execution", jobDef.entryPoint, parsedParams);
                taskResult.readback = execResult.readback;
                taskResult.isPartiallyModified = execResult.isPartiallyModified;

                if (!execResult.success)
                {
                    taskResult.status = "failed";
                    taskResult.failureStage = execResult.failureStage ?? "execution";
                    taskResult.error = execResult.error ?? "宏执行失败";
                    taskResult.executionSummary = execResult.summary;
                    // 保存隔离工作副本以留存失败现场证据
                    try { targetWb.Save(); } catch { }
                    return;
                }

                // 步骤 7：客观读回与状态核验
                // 没有特定断言时，只能声明执行及保存完成，不能宣称任意业务需求已满足
                if (execResult.readback == null || !execResult.readback.targetVerified)
                {
                    taskResult.status = "failed";
                    taskResult.failureStage = "readback";
                    taskResult.error = "宏执行后未能成功读回目标工作簿状态。";
                    return;
                }

                taskResult.executionSummary = string.Format("宏执行成功，读回区域 {0} (行={1}, 列={2}, 公式={3})",
                    execResult.readback.usedRangeAddress ?? "未知",
                    execResult.readback.rowCount,
                    execResult.readback.columnCount,
                    execResult.readback.hasFormulas ? "是" : "否");

                // 步骤 8：另存为最终输出（同名文件自动递增 _1，绝不覆盖已有文件）
                string finalOutputPath = ResolveUniqueOutputPath(jobDef.outputDir, fileItem.originalFileName);
                taskResult.finalOutputPath = finalOutputPath;

                try
                {
                    // 先保存工作副本
                    targetWb.Save();
                    // 复制到最终输出路径
                    File.Copy(workingCopyPath, finalOutputPath, false);
                }
                catch (Exception exSave)
                {
                    taskResult.status = "failed";
                    taskResult.failureStage = "save_output";
                    taskResult.error = "保存输出文件失败: " + exSave.Message;
                    return;
                }

                // 物理核验输出文件存在且字节数大于0
                if (!File.Exists(finalOutputPath) || new FileInfo(finalOutputPath).Length == 0)
                {
                    taskResult.status = "failed";
                    taskResult.failureStage = "save_output";
                    taskResult.error = "输出文件物理核验失败（文件不存在或大小为0）。";
                    return;
                }

                // 步骤 9：只有执行、读回和输出保存均成功，才标记该文件 success
                taskResult.status = "success";
                taskResult.failureStage = null;
            }
            finally
            {
                // 步骤 10：严格资源释放（仅关闭本任务创建的隔离工作副本，绝不关闭用户其他工作簿）
                if (targetWb != null)
                {
                    try
                    {
                        targetWb.Close(false);
                    }
                    catch { }
                    finally
                    {
                        try { Marshal.ReleaseComObject(targetWb); } catch { }
                        targetWb = null;
                    }
                }

                // 步骤 11：临时文件清理规则（成功时清理隔离工作副本；失败或阻断时保留作为排查恢复材料）
                if (taskResult.status == "success")
                {
                    try
                    {
                        if (File.Exists(workingCopyPath))
                        {
                            File.Delete(workingCopyPath);
                            taskResult.temporaryCleaned = true;
                        }
                    }
                    catch { }
                }
                else
                {
                    // 失败保留副本供分析
                    taskResult.temporaryCleaned = false;
                }
            }
        }
    }
}
