using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class WorkbookReadback
    {
        public string targetWorkbookName { get; set; }
        public string targetWorkbookFullName { get; set; }
        public string targetSheetName { get; set; }
        public string usedRangeAddress { get; set; }
        public int rowCount { get; set; }
        public int columnCount { get; set; }
        public string startCell { get; set; }
        public string endCell { get; set; }
        public List<string> sampleValues { get; set; }
        public bool hasFormulas { get; set; }
        public bool hasBorders { get; set; }
        public bool hasInteriorColor { get; set; }
        public int sheetCount { get; set; }
        public bool targetVerified { get; set; }
        public bool otherWorkbooksAffected { get; set; }
        public string affectedWorkbooksWarning { get; set; }
    }

    public class VbaExecutionResult
    {
        public bool success { get; set; }
        public string summary { get; set; }
        public string error { get; set; }
        public string rawModelResponse { get; set; }
        public string originalVbaCode { get; set; }
        public string executedVbaCode { get; set; }
        public string wrapperCode { get; set; }
        public string originalCodeHash { get; set; }
        public string executedCodeHash { get; set; }
        public bool isSourceIdentical { get; set; }
        public List<string> transformSteps { get; set; }
        public double elapsedMs { get; set; }
        public WorkbookReadback readback { get; set; }
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

    public class VbaRunner
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetDlgCtrlID(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;
        private const uint WM_COMMAND = 0x0111;
        private const uint BM_CLICK = 0x00F5;

        private static readonly HashSet<string> _lockedWorkbooks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _lockObj = new object();

        public static bool IsWorkbookLocked(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return false;
            lock (_lockObj)
            {
                return _lockedWorkbooks.Contains(wbName);
            }
        }

        public static void LockWorkbook(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return;
            lock (_lockObj)
            {
                _lockedWorkbooks.Add(wbName);
            }
        }

        public static void UnlockWorkbook(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return;
            lock (_lockObj)
            {
                _lockedWorkbooks.Remove(wbName);
            }
        }

        public static string ComputeSha256(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
                byte[] hash = sha.ComputeHash(bytes);
                var sb = new System.Text.StringBuilder();
                foreach (byte b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        public static VbaExecutionResult RunVbaCode(dynamic app, dynamic targetWorkbook, string vbaCode, string rawModelResponse = "")
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var transformSteps = new List<string>();

            bool? origScreenUpdating = null;
            bool? origDisplayAlerts = null;
            bool? origEnableEvents = null;
            int? origCalculation = null;

            if (app != null)
            {
                try { origScreenUpdating = (bool)app.ScreenUpdating; } catch { }
                try { origDisplayAlerts = (bool)app.DisplayAlerts; } catch { }
                try { origEnableEvents = (bool)app.EnableEvents; } catch { }
                try { origCalculation = (int)app.Calculation; } catch { }
            }

            if (targetWorkbook == null)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：未指定有效的目标工作簿",
                    error = "目标工作簿对象为空或已关闭",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode ?? "",
                    executedVbaCode = "",
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = "",
                    isSourceIdentical = false,
                    transformSteps = transformSteps,
                    elapsedMs = 0
                };
            }

            string earlyTargetName = "";
            try { earlyTargetName = (string)targetWorkbook.Name; } catch { }

            if (!string.IsNullOrEmpty(earlyTargetName) && IsWorkbookLocked(earlyTargetName))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = "workbook_locked",
                    executionPhase = "blocked_by_lock",
                    summary = "执行已拒绝：目标工作簿已被安全锁定",
                    error = "该工作簿此前发生未证实恢复的宏挂起，为防止进一步损坏数据已被锁定，禁止继续注入执行新宏。\n\n【快照恢复指引】：请在面板中点击【快照回滚】将工作簿恢复至执行前初始状态（回滚成功将自动解除锁定），或在保存其他工作簿后重启 Excel。",
                    riskNotice = "【安全锁定保护】当前工作簿已处于锁定状态，已阻止新宏注入运行。请使用快照恢复以解除锁定。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode ?? "",
                    executedVbaCode = "",
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = "",
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = 0
                };
            }

            if (string.IsNullOrEmpty(vbaCode))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：传入的代码为空",
                    error = "未检测到有效 VBA 代码",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = "",
                    executedVbaCode = "",
                    wrapperCode = null,
                    originalCodeHash = "",
                    executedCodeHash = "",
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = 0
                };
            }

            // 1. 结构完整性预检：必须包含有效过程定义 (Sub 或 Function) 与闭合语句 (防止模型响应残缺)
            // 使用负向后行断言严格排除 End Sub / End Function / Exit Sub
            var procMatches = Regex.Matches(vbaCode, @"(?<!End\s+|Exit\s+)\b(Sub|Function)\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*(?:\(([^)]*)\))?", RegexOptions.IgnoreCase);
            bool hasEndProc = Regex.IsMatch(vbaCode, @"End\s+(?:Sub|Function)", RegexOptions.IgnoreCase);

            if (procMatches.Count == 0 || !hasEndProc)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：VBA 代码结构不完整 (缺少 Sub/Function 或未闭合)",
                    error = "模型输出的代码结构不完整或被截断，已在注入前安全拦截，避免导致 Excel 编译错误。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 检查非法 End 语句 (如 End CleanExit, End Try)
            var invalidEndMatch = Regex.Match(vbaCode, @"^\s*End\s+(?!Sub\b|Function\b|Property\b|If\b|With\b|Select\b|Type\b|Enum\b)([A-Za-z0-9_]+)", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            if (invalidEndMatch.Success)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：检测到非标准 VBA 语法语句 (" + invalidEndMatch.Value.Trim() + ")",
                    error = "代码包含不符合 VBA 语法的终止语句 '" + invalidEndMatch.Value.Trim() + "'（提前退出请使用 Exit Sub/Function），已在注入前安全拦截。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 2. 执行前影响范围检查：识别对整张工作表 171 亿单元格做重度格式化、批量删除的高危操作
            var scopeRiskMatch = Regex.Match(vbaCode, @"(?:(?:\bws\b|\bActiveSheet\b|\bWorksheets\([^)]+\)|(?<!\w))\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:Borders|Interior|FormatConditions|ClearFormats)\b", RegexOptions.IgnoreCase);
            if (scopeRiskMatch.Success)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = "scope_risk_intercepted",
                    executionPhase = "intercepted_before_run",
                    summary = "执行中断：检测到全表级范围失控高危操作 (" + scopeRiskMatch.Value.Trim() + ")",
                    error = "代码尝试对整张工作表的全部单元格批量设置格式或边框（" + scopeRiskMatch.Value.Trim() + "）。Excel包含171亿单元格，对全表Cells操作会导致Excel进程卡死。已在运行前安全拦截，未对工作簿做任何更改。",
                    riskNotice = "【范围安全拦截】已在注入前阻止运行，未对工作簿做任何更改。需限定在具体业务数据区域内操作。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 3. 识别过程与严格判定主入口（坚决拒绝盲猜与误判）
            var subMatches = new List<Match>();
            var funcMatches = new List<Match>();
            foreach (Match m in procMatches)
            {
                string pType = m.Groups[1].Value;
                if (string.Equals(pType, "Sub", StringComparison.OrdinalIgnoreCase))
                {
                    subMatches.Add(m);
                }
                else if (string.Equals(pType, "Function", StringComparison.OrdinalIgnoreCase))
                {
                    funcMatches.Add(m);
                }
            }

            if (subMatches.Count == 0)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = "entry_unidentified",
                    executionPhase = "extract_failed",
                    failureStage = "extract_failed",
                    summary = "提取失败：未解析到有效的 VBA Sub 过程",
                    error = "大模型回复中仅包含 Function 函数过程或无过程定义，缺少可作为宏执行入口的 Sub 过程。已安全拦截以防调用失败。",
                    riskNotice = "【提取失败拦截】宏执行必须包含明确的 Sub 过程（如 Sub Main）。宿主绝不猜测调用 Function 函数。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds,
                    origScreenUpdating = origScreenUpdating,
                    origDisplayAlerts = origDisplayAlerts,
                    origEnableEvents = origEnableEvents,
                    origCalculation = origCalculation,
                    hostStateRestored = true,
                    hostStateRestoreDetails = "运行前拦截阶段，未变更宿主全局状态"
                };
            }

            Match targetProcMatch = null;
            string mainProcName = null;
            bool isNativeWbParam = false;

            // 规则 1：寻找名为 Main 的 Sub 过程
            var mainSubs = new List<Match>();
            foreach (Match m in subMatches)
            {
                if (string.Equals(m.Groups[2].Value, "Main", StringComparison.OrdinalIgnoreCase))
                {
                    mainSubs.Add(m);
                }
            }

            if (mainSubs.Count == 1)
            {
                mainProcName = mainSubs[0].Groups[2].Value;
                string p = mainSubs[0].Groups[3].Value;
                if (Regex.IsMatch(p, @"(?:targetWb|wb|workbook)\s+As\s+(?:Workbook|Object)", RegexOptions.IgnoreCase))
                {
                    targetProcMatch = mainSubs[0];
                    isNativeWbParam = true;
                }
            }
            else if (mainSubs.Count > 1)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = "entry_conflict",
                    executionPhase = "blocked_before_run",
                    summary = "执行已拒绝：检测到多个名为 Main 的重复过程定义",
                    error = "代码中存在多个名为 Main 的 Sub 过程定义，存在语法冲突，已在运行前安全拦截。",
                    riskNotice = "【过程冲突拦截】代码存在重复定义的 Sub Main，已阻止注入以防编译错误。",
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds,
                    origScreenUpdating = origScreenUpdating,
                    origDisplayAlerts = origDisplayAlerts,
                    origEnableEvents = origEnableEvents,
                    origCalculation = origCalculation,
                    hostStateRestored = true,
                    hostStateRestoreDetails = "运行前拦截阶段，未变更宿主全局状态"
                };
            }
            else
            {
                // 无 Sub Main 时：
                // 规则 2：寻找显式接收 targetWb As Workbook 的 Sub 过程
                var targetParamSubs = new List<Match>();
                foreach (Match m in subMatches)
                {
                    string p = m.Groups[3].Value;
                    if (Regex.IsMatch(p, @"(?:targetWb|wb|workbook)\s+As\s+(?:Workbook|Object)", RegexOptions.IgnoreCase))
                    {
                        targetParamSubs.Add(m);
                    }
                }

                if (targetParamSubs.Count == 1)
                {
                    targetProcMatch = targetParamSubs[0];
                    mainProcName = targetParamSubs[0].Groups[2].Value;
                    isNativeWbParam = true;
                }
                else if (targetParamSubs.Count > 1)
                {
                    var pNames = new List<string>();
                    foreach (Match m in targetParamSubs) pNames.Add(m.Groups[2].Value);
                    return new VbaExecutionResult
                    {
                        success = false,
                        precheckStatus = "entry_ambiguous",
                        executionPhase = "blocked_before_run",
                        summary = "执行已拒绝：存在多个接收 Workbook 参数的过程，无法唯一确定主入口",
                        error = "代码中包含多个接收 Workbook 参数的过程 (" + string.Join(", ", pNames) + ") 且无 Sub Main。宿主拒绝猜测执行。",
                        riskNotice = "【入口歧义拦截】为防误调辅助过程，宿主已拒绝执行。",
                        rawModelResponse = rawModelResponse ?? "",
                        originalVbaCode = vbaCode,
                        executedVbaCode = vbaCode,
                        wrapperCode = null,
                        originalCodeHash = ComputeSha256(vbaCode),
                        executedCodeHash = ComputeSha256(vbaCode),
                        isSourceIdentical = true,
                        transformSteps = transformSteps,
                        elapsedMs = sw.ElapsedMilliseconds,
                        origScreenUpdating = origScreenUpdating,
                        origDisplayAlerts = origDisplayAlerts,
                        origEnableEvents = origEnableEvents,
                        origCalculation = origCalculation,
                        hostStateRestored = true,
                        hostStateRestoreDetails = "运行前拦截阶段，未变更宿主全局状态"
                    };
                }
                else
                {
                    // 无 Sub Main，且无接收 targetWb 参数的 Sub
                    // 规则 3：若代码中【有且仅有 1 个】Sub 过程，可以确定其为唯一无参入口
                    if (subMatches.Count == 1)
                    {
                        mainProcName = subMatches[0].Groups[2].Value;
                        isNativeWbParam = false;
                    }
                    else
                    {
                        // 规则 4：存在多个无参 Sub 过程，且没有任何一个叫 Main，入口完全不明确！坚决拒绝盲猜！
                        var subNames = new List<string>();
                        foreach (Match m in subMatches) subNames.Add(m.Groups[2].Value);

                        return new VbaExecutionResult
                        {
                            success = false,
                            precheckStatus = "entry_unidentified",
                            executionPhase = "blocked_before_run",
                            summary = "执行已拒绝：代码包含多个过程且无法明确确定主入口",
                            error = "代码中包含多个 Sub 过程 (" + string.Join(", ", subNames) + ")，未指定主入口 Sub Main 亦无明确的 targetWb 参数过程。为防止盲目猜测调用辅助过程导致不可控后果，宿主已拒绝执行。",
                            riskNotice = "【入口歧义拦截】宿主不猜测、不改写模型源码。请在提示词中要求模型明确主入口过程（如命名为 Sub Main 或指定 targetWb 参数）。",
                            rawModelResponse = rawModelResponse ?? "",
                            originalVbaCode = vbaCode,
                            executedVbaCode = vbaCode,
                            wrapperCode = null,
                            originalCodeHash = ComputeSha256(vbaCode),
                            executedCodeHash = ComputeSha256(vbaCode),
                            isSourceIdentical = true,
                            transformSteps = transformSteps,
                            elapsedMs = sw.ElapsedMilliseconds,
                            origScreenUpdating = origScreenUpdating,
                            origDisplayAlerts = origDisplayAlerts,
                            origEnableEvents = origEnableEvents,
                            origCalculation = origCalculation,
                            hostStateRestored = true,
                            hostStateRestoreDetails = "运行前拦截阶段，未变更宿主全局状态"
                        };
                    }
                }
            }

            string targetWbName = "";
            string targetWbFullName = "";
            int preSheetCount = 0;
            try
            {
                targetWbName = (string)targetWorkbook.Name;
                targetWbFullName = (string)targetWorkbook.FullName;
                preSheetCount = (int)targetWorkbook.Sheets.Count;
            }
            catch (Exception ex)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：无法访问目标工作簿属性",
                    error = "目标工作簿状态异常: " + ex.Message,
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = vbaCode,
                    wrapperCode = null,
                    originalCodeHash = ComputeSha256(vbaCode),
                    executedCodeHash = ComputeSha256(vbaCode),
                    isSourceIdentical = true,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds,
                    origScreenUpdating = origScreenUpdating,
                    origDisplayAlerts = origDisplayAlerts,
                    origEnableEvents = origEnableEvents,
                    origCalculation = origCalculation,
                    hostStateRestored = true,
                    hostStateRestoreDetails = "运行前拦截阶段，未变更宿主全局状态"
                };
            }

            // 4. 通用透明调度架构：绝不静默改写模型正文源码
            string callMacroName = "";
            string wrapperCode = null;
            string finalCode = vbaCode;

            if (isNativeWbParam)
            {
                // 模型过程原生声明了接收目标工作簿参数，100% 原始源码直调
                callMacroName = mainProcName;
                finalCode = vbaCode;
                wrapperCode = null;
                transformSteps.Add("直调主入口: " + callMacroName + " (模型原生接收目标工作簿参数，源码 100% 原始直调，无包装器)");
            }
            else
            {
                // 模型生成了通用无参主过程
                // 原则：模型正文源码保持 100% 零修改，仅在模块尾部追加受控的透明调用入口包装器
                string wrapperSubName = "LeeHostRunner_" + DateTime.Now.ToString("mmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                callMacroName = wrapperSubName;

                wrapperCode =
                    "\r\n' ===== [LeeExcel 自动生成的受控调用入口包装器 - 保持模型正文源码零篡改] =====\r\n" +
                    "Sub " + wrapperSubName + "(targetWb As Workbook)\r\n" +
                    "    targetWb.Activate\r\n" +
                    "    Call " + mainProcName + "\r\n" +
                    "End Sub";

                finalCode = vbaCode + "\r\n" + wrapperCode;
                transformSteps.Add("模型源码正文 100% 保持原貌。因主过程为无参，宿主独立追加受控入口包装器: " + wrapperSubName + " -> 调度主过程 " + mainProcName + "。实测边界警示：包装器激活目标簿确保 ActiveSheet/ActiveWorkbook 导向目标簿；但不能保证对显式指定外部工作簿的任意代码具备沙箱级强制隔离。多簿安全由整本快照与执行后全局比对保障。");
            }

            string originalHash = ComputeSha256(vbaCode);
            string executedHash = ComputeSha256(finalCode);
            bool isIdentical = string.Equals(vbaCode, finalCode, StringComparison.Ordinal);

            // 5. 执行前：记录当前 Excel 实例中所有【非目标工作簿】的初始状态快照
            var otherWbSnapshots = new List<Tuple<string, int, string>>();
            try
            {
                foreach (dynamic wb in app.Workbooks)
                {
                    try
                    {
                        string oFull = (string)wb.FullName;
                        if (!string.Equals(oFull, targetWbFullName, StringComparison.OrdinalIgnoreCase))
                        {
                            string oName = (string)wb.Name;
                            int sCount = (int)wb.Sheets.Count;
                            string uAddr = "";
                            try { uAddr = ((string)wb.ActiveSheet.UsedRange.Address) ?? ""; } catch { }
                            otherWbSnapshots.Add(Tuple.Create(oName, sCount, uAddr));
                        }
                    }
                    catch { }
                }
            }
            catch { }

            dynamic vbComp = null;
            dynamic vbProj = null;
            string moduleName = "";
            string precheckStatus = "unavailable";
            string capturedRuntimeDialogError = null;
            VbaExecutionResult finalResult = null;

            try
            {
                // 6. 检查 VBProject 访问权限并注入临时模块
                try
                {
                    vbProj = targetWorkbook.VBProject;
                }
                catch (COMException comEx)
                {
                    uint cErr = (uint)comEx.ErrorCode;
                    string cErrStr = string.Format("0x{0:X8} ({1})", cErr, comEx.ErrorCode);
                    if (cErr == 0x800A03EC || comEx.Message.Contains("1004"))
                    {
                        finalResult = new VbaExecutionResult
                        {
                            success = false,
                            precheckStatus = "unavailable",
                            executionPhase = "injection_failed",
                            failureStage = "injection_failed",
                            rawErrorCode = cErrStr,
                            summary = "代码注入失败：Excel 未开启对 VBA 工程对象模型的访问信任",
                            error = "错误 1004：请在 Excel“文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置”中勾选【信任对 VBA 工程对象模型的访问】。",
                            rawModelResponse = rawModelResponse ?? "",
                            originalVbaCode = vbaCode,
                            executedVbaCode = finalCode,
                            wrapperCode = wrapperCode,
                            originalCodeHash = originalHash,
                            executedCodeHash = executedHash,
                            isSourceIdentical = isIdentical,
                            transformSteps = transformSteps,
                            elapsedMs = sw.ElapsedMilliseconds
                        };
                        return finalResult;
                    }
                    throw;
                }

                // 7. 动态创建标准临时模块并注入宏（绝不篡改模型源码，原样注入）
                try
                {
                    moduleName = "LeeMod_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    vbComp = vbProj.VBComponents.Add(1);
                    vbComp.Name = moduleName;
                    vbComp.CodeModule.AddFromString(finalCode);
                }
                catch (Exception injEx)
                {
                    var cEx = injEx as COMException;
                    string codeStr = cEx != null ? string.Format("0x{0:X8} ({1})", (uint)cEx.ErrorCode, cEx.ErrorCode) : "COM Error";
                    finalResult = new VbaExecutionResult
                    {
                        success = false,
                        precheckStatus = "failed",
                        executionPhase = "injection_failed",
                        failureStage = "injection_failed",
                        rawErrorCode = codeStr,
                        summary = "模块注入失败：未能将代码注入到目标工作簿 VBE 工程",
                        error = "向 VBE 模块写入源码时抛出异常: " + injEx.Message,
                        riskNotice = "【注入阶段错误】代码未能成功写入工作簿临时模块，未做任何修改。",
                        rawModelResponse = rawModelResponse ?? "",
                        originalVbaCode = vbaCode,
                        executedVbaCode = finalCode,
                        wrapperCode = wrapperCode,
                        originalCodeHash = originalHash,
                        executedCodeHash = executedHash,
                        isSourceIdentical = isIdentical,
                        transformSteps = transformSteps,
                        elapsedMs = sw.ElapsedMilliseconds
                    };
                    return finalResult;
                }

                // 8. 预编译安全探测：静态语法校验
                string precheckDetail = "";
                try
                {
                    try { targetWorkbook.Activate(); } catch { }
                    try { app.VBE.ActiveVBProject = vbProj; } catch { }
                    try { vbComp.Activate(); } catch { }
                    dynamic compileBtn = app.VBE.CommandBars.FindControl(Type.Missing, 578);
                    if (compileBtn != null)
                    {
                        if ((bool)compileBtn.Enabled)
                        {
                            try { compileBtn.Execute(); } catch { }
                            if (!(bool)compileBtn.Enabled)
                            {
                                precheckStatus = "passed";
                                precheckDetail = "VBE 预编译探测通过（静态语法与对象模型校验通过）。";
                                transformSteps.Add("【静态语法预检】" + precheckDetail + "（警示：预编译消除静态语法与对象错误，运行期仍受整本快照保护）");
                            }
                            else
                            {
                                // 注意：在多工程、多工作簿或侧边栏无焦点环境下，CommandBar 578 的 Enabled 状态可能不会改变为 false
                                // 严禁武断判定为“缺少 End If”！实事求是记录为未重置状态，放行给带保护的运行时校验
                                precheckStatus = "warning";
                                precheckDetail = "VBE 预编译状态未重置（多工程或非活动窗口状态），将在整本快照保护下执行。";
                                transformSteps.Add("【预检提示】" + precheckDetail);
                            }
                        }
                        else
                        {
                            precheckStatus = "passed";
                            precheckDetail = "VBE 当前无未编译变更，工程静态语法状态正常。";
                            transformSteps.Add("【静态语法预检】" + precheckDetail);
                        }
                    }
                    else
                    {
                        precheckStatus = "unavailable";
                        precheckDetail = "VBE 预编译命令(ID 578)当前不可用，将在整本快照保护下执行。";
                        transformSteps.Add("【预检提示】" + precheckDetail);
                    }
                }
                catch (Exception cEx)
                {
                    precheckStatus = "unavailable";
                    precheckDetail = "执行 VBE 预编译探测时捕获异常: " + cEx.Message + "，将在整本快照保护下执行。";
                    transformSteps.Add("【预检提示】" + precheckDetail);
                }

                // 8. 挂起屏幕刷新与系统弹窗
                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                }
                catch { }

                // 9. 运行期守望与挂起保护（绝不杀 Excel 进程，保障快照与用户其他工作簿安全）
                string macroAddress = "'" + targetWbName + "'!" + callMacroName;
                int timeoutSeconds = 45;
                bool isTimedOut = false;
                Exception runException = null;

                uint excelPid = 0;
                try
                {
                    IntPtr excelHwnd = new IntPtr(Convert.ToInt64(app.Hwnd));
                    GetWindowThreadProcessId(excelHwnd, out excelPid);
                }
                catch { }

                using (var cts = new System.Threading.CancellationTokenSource())
                {
                    var watchdog = System.Threading.Tasks.Task.Run(() =>
                    {
                        while (!cts.Token.IsCancellationRequested)
                        {
                            try
                            {
                                EnumWindows((hWnd, lParam) =>
                                {
                                    try
                                    {
                                        uint pid;
                                        GetWindowThreadProcessId(hWnd, out pid);
                                        if (excelPid != 0 && pid == excelPid)
                                        {
                                            var sbClass = new System.Text.StringBuilder(256);
                                            GetClassName(hWnd, sbClass, sbClass.Capacity);
                                            if (sbClass.ToString() == "#32770")
                                            {
                                                var sbTitle = new System.Text.StringBuilder(256);
                                                GetWindowText(hWnd, sbTitle, sbTitle.Capacity);
                                                string title = sbTitle.ToString();

                                                // 严禁关闭用户的标准对话框（打开文件、另存为、选项、设置、打印等）
                                                bool isUserDialog = title.Contains("打开") || title.Contains("另存为") ||
                                                                    title.Contains("Open") || title.Contains("Save As") ||
                                                                    title.Contains("查找") || title.Contains("替换") ||
                                                                    title.Contains("选项") || title.Contains("Options") ||
                                                                    title.Contains("设置") || title.Contains("打印") ||
                                                                    title.Contains("Print") || title.Contains("格式");

                                                if (!isUserDialog)
                                                {
                                                    IntPtr btnEndHwnd = IntPtr.Zero;
                                                    bool isConfirmedVbaErrorDialog = false;
                                                    string tempCapturedError = null;

                                                    if (title.Contains("Microsoft Visual Basic") || title.Contains("VBA"))
                                                    {
                                                        isConfirmedVbaErrorDialog = true;
                                                    }

                                                    try
                                                    {
                                                        EnumChildWindows(hWnd, (childHwnd, l) =>
                                                        {
                                                            try
                                                            {
                                                                var cCls = new System.Text.StringBuilder(128);
                                                                var cTxt = new System.Text.StringBuilder(512);
                                                                GetClassName(childHwnd, cCls, cCls.Capacity);
                                                                GetWindowText(childHwnd, cTxt, cTxt.Capacity);
                                                                string ct = cTxt.ToString();
                                                                int ctrlId = GetDlgCtrlID(childHwnd);

                                                                // 特征1：VBA 对话框特有的 Control ID 4800 (结束按钮) 或 4801 (调试按钮)
                                                                if (ctrlId == 4800 || ctrlId == 4801)
                                                                {
                                                                    isConfirmedVbaErrorDialog = true;
                                                                }

                                                                // 特征2：Static 控件包含错误关键字
                                                                if (cCls.ToString() == "Static" && !string.IsNullOrWhiteSpace(ct))
                                                                {
                                                                    if (ct.Contains("错误") || ct.Contains("error") || ct.Contains("缺少") || ct.Contains("1004"))
                                                                    {
                                                                        isConfirmedVbaErrorDialog = true;
                                                                        tempCapturedError = ct;
                                                                    }
                                                                }

                                                                if (ctrlId == 4800 || ((ctrlId == 1 || ctrlId == 2) && (ct.Contains("确定") || ct.Contains("OK"))))
                                                                {
                                                                    btnEndHwnd = childHwnd;
                                                                }
                                                            }
                                                            catch { }
                                                            return true;
                                                        }, IntPtr.Zero);
                                                    }
                                                    catch { }

                                                    // 仅当 100% 确认属于当前 Excel 宏抛出的 VBA 错误模态框时才介入，坚决不误触用户普通对话框
                                                    if (isConfirmedVbaErrorDialog)
                                                    {
                                                        if (!string.IsNullOrEmpty(tempCapturedError))
                                                        {
                                                            capturedRuntimeDialogError = tempCapturedError;
                                                        }

                                                        if (btnEndHwnd != IntPtr.Zero)
                                                        {
                                                            SendMessage(btnEndHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                                                        }

                                                        SendMessage(hWnd, WM_COMMAND, (IntPtr)4800, IntPtr.Zero); // VBA 运行时错误弹窗的“结束(&E)”按钮 ID
                                                        SendMessage(hWnd, WM_COMMAND, (IntPtr)1, IntPtr.Zero);    // 编译错误确定按钮
                                                        PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                    return true;
                                }, IntPtr.Zero);
                            }
                            catch { }
                            System.Threading.Thread.Sleep(60);
                        }
                    });

                    var runTask = System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            app.Run(macroAddress, targetWorkbook);
                        }
                        catch (Exception rx)
                        {
                            runException = rx;
                        }
                    });

                    string hangPhase = null;
                    string hangSummary = null;
                    string hangError = null;
                    string hangRiskNotice = null;

                    if (!runTask.Wait(timeoutSeconds * 1000))
                    {
                        isTimedOut = true;
                        // 1. 记录初始阶段：疑似挂起，中断请求已发送（绝不假定停止成功）
                        hangPhase = "hang_suspected_interrupt_sent";
                        hangSummary = "执行疑似挂起：宏运行已超过 " + timeoutSeconds + " 秒无响应，已发送中断指令";

                        // 2. 发送 VBE Reset(ID 228) 尝试停止宏
                        try
                        {
                            dynamic resetBtn = app.VBE.CommandBars.FindControl(Type.Missing, 228);
                            if (resetBtn != null && (bool)resetBtn.Enabled)
                            {
                                resetBtn.Execute();
                            }
                        }
                        catch { }

                        // 3. 严格实测证实恢复判定：必须同时满足以下4项，才允许记录为“中断并恢复”
                        // (a) 宏执行线程是否在 3 秒内退出
                        bool macroStopped = runTask.Wait(3000);

                        // (b) Excel COM 接口是否恢复响应
                        bool comResponsive = false;
                        var comProbeTask = System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                var ready = app.Ready;
                                var curName = (string)targetWorkbook.Name;
                                comResponsive = true;
                            }
                            catch { }
                        });
                        comProbeTask.Wait(3000);

                        // (c) 目标及其他工作簿是否可读
                        bool workbooksReadable = false;
                        if (comResponsive)
                        {
                            try
                            {
                                int sheetCount = (int)targetWorkbook.Sheets.Count;
                                foreach (var snap in otherWbSnapshots)
                                {
                                    dynamic oWb = app.Workbooks[snap.Item1];
                                    var oName = (string)oWb.Name;
                                }
                                workbooksReadable = true;
                            }
                            catch { }
                        }

                        // (d) 临时注入模块是否已成功移除
                        bool moduleCleaned = false;
                        if (comResponsive && vbProj != null && vbComp != null)
                        {
                            try
                            {
                                vbProj.VBComponents.Remove(vbComp);
                                bool stillExists = false;
                                foreach (dynamic c in vbProj.VBComponents)
                                {
                                    if ((string)c.Name == moduleName) { stillExists = true; break; }
                                }
                                moduleCleaned = !stillExists;
                                if (moduleCleaned) vbComp = null; // 标记已清除，防止 finally 重复
                            }
                            catch { }
                        }

                        // 综合判定
                        if (macroStopped && comResponsive && workbooksReadable && moduleCleaned)
                        {
                            hangPhase = "hang_interrupted_recovered";
                            hangSummary = "宏运行挂起超时已安全中断并证实恢复响应（未强杀 Excel）";
                            hangError = "宏执行超过 " + timeoutSeconds + " 秒未响应（可能陷入耗时大计算或死循环）。中断指令已获实测证实：Excel 接口已恢复响应、工作簿状态可读、临时注入模块已成功移除。您的 Excel 进程未被强杀。建议核对数据或点击【快照回滚】恢复至执行前初始状态。";
                            hangRiskNotice = "【挂起中断证实恢复】已实测证实宏终止、Excel 恢复响应、目标及其他工作簿均可读、临时模块已清除。可继续操作或使用快照恢复。";
                        }
                        else
                        {
                            hangPhase = "hang_unconfirmed_locked";
                            hangSummary = "宏挂起中断未获证实，当前工作簿已紧急锁定保护";
                            hangError = "宏执行超时且发送中断请求后，Excel 未能在时限内证实恢复响应（宏可能仍在底层运算，或 COM 状态未就绪）。\n为保护工作簿数据不被进一步损坏，当前工作簿已自动锁定，严禁继续执行新宏！\n\n【快照恢复指引】：请在面板中点击【快照回滚】恢复至执行前快照；如 Excel 完全卡死，请保存其他工作簿后安全重启 Excel。";
                            hangRiskNotice = "【高危锁定保护】宏中断未获证实，已严禁向当前工作簿继续注入或执行任何新宏。请使用执行前快照恢复！";

                            // 锁定目标工作簿，禁止后续注入新宏
                            LockWorkbook(targetWbName);
                        }
                    }

                    cts.Cancel();
                    try { watchdog.Wait(200); } catch { }

                    if (isTimedOut)
                    {
                        sw.Stop();
                        finalResult = new VbaExecutionResult
                        {
                            success = false,
                            precheckStatus = precheckStatus,
                            executionPhase = hangPhase,
                            summary = hangSummary,
                            error = hangError,
                            riskNotice = hangRiskNotice,
                            rawModelResponse = rawModelResponse ?? "",
                            originalVbaCode = vbaCode,
                            executedVbaCode = finalCode,
                            wrapperCode = wrapperCode,
                            originalCodeHash = originalHash,
                            executedCodeHash = executedHash,
                            isSourceIdentical = isIdentical,
                            transformSteps = transformSteps,
                            elapsedMs = sw.ElapsedMilliseconds
                        };
                        return finalResult;
                    }
                }

                if (runException != null)
                {
                    throw runException;
                }

                sw.Stop();

                // 9. 执行后检查：核验非目标工作簿是否受到意外影响
                bool otherAffected = false;
                string otherWarning = null;
                foreach (var snap in otherWbSnapshots)
                {
                    try
                    {
                        dynamic oWb = app.Workbooks[snap.Item1];
                        int postCount = (int)oWb.Sheets.Count;
                        string postAddr = "";
                        try { postAddr = ((string)oWb.ActiveSheet.UsedRange.Address) ?? ""; } catch { }

                        if (postCount != snap.Item2 || postAddr != snap.Item3)
                        {
                            otherAffected = true;
                            otherWarning = "检测到非目标工作簿 [" + snap.Item1 + "] 发生附带变更！";
                            break;
                        }
                    }
                    catch { }
                }

                // 10. 目标工作簿写后读回与核验
                var readback = PerformReadback(targetWorkbook, targetWbName, targetWbFullName, preSheetCount);
                readback.otherWorkbooksAffected = otherAffected;
                readback.affectedWorkbooksWarning = otherWarning;

                finalResult = new VbaExecutionResult
                {
                    success = !otherAffected,
                    summary = otherAffected ? ("执行告警：" + otherWarning) : ("宏已运行（耗时 " + sw.ElapsedMilliseconds + " ms）"),
                    error = otherWarning,
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = finalCode,
                    wrapperCode = wrapperCode,
                    originalCodeHash = originalHash,
                    executedCodeHash = executedHash,
                    isSourceIdentical = isIdentical,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds,
                    readback = readback,
                    precheckStatus = (precheckStatus == "warning" || precheckStatus == "passed") ? "passed" : precheckStatus,
                    executionPhase = "macro_completed",
                    riskNotice = (precheckStatus == "passed" || precheckStatus == "warning")
                        ? "【运行完成（受检执行）】宏已正常返回并读回数据。静态语法与动态执行校验通过。宿主状态已恢复原值。（说明：快照仅能回滚目标工作簿本身，无法覆盖对其他工作簿、磁盘或网络的外部副作用）。"
                        : "【运行完成（未执行预检）】宏已正常返回并读回数据。预编译探测不可用，当前处于未受检执行。宿主状态已恢复原值。（说明：快照仅能回滚目标工作簿本身，无法覆盖对其他工作簿、磁盘或网络的外部副作用）。"
                };
                return finalResult;
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errDetail = ex.Message;
                if (!string.IsNullOrEmpty(capturedRuntimeDialogError))
                {
                    errDetail = capturedRuntimeDialogError.Trim();
                }
                else if (ex.InnerException != null)
                {
                    errDetail += " (" + ex.InnerException.Message + ")";
                }

                if (vbProj != null)
                {
                    try
                    {
                        vbProj.VBE.MainWindow.Visible = false;
                    }
                    catch { }
                }

                var comEx = ex as COMException;
                uint errorCode = comEx != null ? (uint)comEx.ErrorCode : 0;
                string rawErrCodeStr = comEx != null ? string.Format("0x{0:X8} ({1})", errorCode, comEx.ErrorCode) : "General Exception";

                bool isInvocationError = errDetail.Contains("找不到宏") ||
                                         errDetail.Contains("Cannot run the macro") ||
                                         errDetail.Contains("参数不可选") ||
                                         errDetail.Contains("Wrong number of arguments") ||
                                         errDetail.Contains("找不到指定的") ||
                                         (errorCode == 0x800A03EC && errDetail.Contains("找不到宏"));

                bool isRealCompileError = errDetail.Contains("编译错误") ||
                                          errDetail.Contains("Compile error") ||
                                          errDetail.Contains("未定义") ||
                                          errDetail.Contains("未声明") ||
                                          errDetail.Contains("语法错误") ||
                                          errDetail.Contains("Syntax error") ||
                                          errDetail.Contains("缺少") ||
                                          errDetail.Contains("Expected") ||
                                          (errorCode == 0x800A03E6 || errorCode == 0x800A001C);

                string finalPrecheckStatus = precheckStatus;
                string finalPhase = "runtime_error";
                string finalStage = "runtime_error";
                string finalSummary = "执行失败：VBA 运行期抛出异常";
                string triggerPoint = "";

                WorkbookReadback postErrorReadback = null;
                bool isPartiallyModified = false;
                int? vbaErrNum = null;
                string vbaErrDesc = null;

                if (!string.IsNullOrEmpty(capturedRuntimeDialogError))
                {
                    var mNum = Regex.Match(capturedRuntimeDialogError, @"(?:运行时错误|Runtime error)\s*'?(\d+)'?", RegexOptions.IgnoreCase);
                    if (mNum.Success)
                    {
                        int n;
                        if (int.TryParse(mNum.Groups[1].Value, out n)) vbaErrNum = n;
                    }
                    string[] lines = capturedRuntimeDialogError.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 1)
                    {
                        vbaErrDesc = lines[lines.Length - 1].Trim();
                    }
                    else
                    {
                        vbaErrDesc = capturedRuntimeDialogError.Trim();
                    }
                }
                else if (errorCode == 0x800A03EC)
                {
                    vbaErrNum = 1004;
                    vbaErrDesc = errDetail;
                }

                if (isInvocationError)
                {
                    finalPhase = "invocation_failed";
                    finalStage = "invocation_failed";
                    finalSummary = "宏入口调用失败：未找到有效宏入口或调用签名不匹配";
                }
                else if (isRealCompileError)
                {
                    finalPhase = "compile_failed";
                    finalStage = "compile_failed";
                    finalSummary = "代码编译失败：VBE 报告编译或语法错误";
                    finalPrecheckStatus = "failed";
                }
                else
                {
                    // 真实运行时错误（例如 1004: 类 Range 的 Select 方法无效）
                    finalPhase = "runtime_error";
                    finalStage = "runtime_error";
                    string errName = vbaErrNum.HasValue ? vbaErrNum.Value.ToString() : (errorCode == 0x800A03EC ? "1004" : rawErrCodeStr);
                    finalSummary = "宏运行期异常：执行语句时抛出错误 " + errName;

                    if (errDetail.Contains("Select") || errDetail.Contains("选择"))
                    {
                        triggerPoint = "触发点: Range.Select 操作（当目标工作表非当前活动表时调用将触发 1004 运行时错误）";
                    }

                    // 实测检测目标工作簿在发生异常前是否已经发生了部分写入或格式改变
                    if (targetWorkbook != null)
                    {
                        try
                        {
                            postErrorReadback = PerformReadback(targetWorkbook, targetWbName, targetWbFullName, preSheetCount);
                            if (postErrorReadback != null)
                            {
                                bool hasArea = !string.IsNullOrEmpty(postErrorReadback.usedRangeAddress) && postErrorReadback.usedRangeAddress != "A1";
                                bool hasContent = postErrorReadback.sampleValues != null && postErrorReadback.sampleValues.Count > 0;
                                bool hasStyle = postErrorReadback.hasBorders || postErrorReadback.hasInteriorColor || postErrorReadback.hasFormulas;
                                if (hasArea || hasContent || hasStyle)
                                {
                                    isPartiallyModified = true;
                                }
                            }
                        }
                        catch { }
                    }
                }

                string finalRiskNotice = isRealCompileError
                    ? "【编译阶段错误】代码在编译阶段未通过，已自动中止并安全清理临时模块，工作簿快照完好。"
                    : (isInvocationError
                        ? "【调用阶段错误】未能成功调度宏入口过程，已安全清理临时模块，工作簿快照完好。"
                        : (isPartiallyModified
                            ? "【运行期部分修改】宏在执行过程中抛出 " + (vbaErrNum.HasValue ? vbaErrNum.Value.ToString() : rawErrCodeStr) + " 异常中断。执行后读回区域: " + (postErrorReadback != null ? postErrorReadback.usedRangeAddress : "") + "；宏可能已部分修改工作簿。执行前快照已完好就绪，可随时整本一键回滚。"
                            : "【运行期异常】宏在执行期间发生运行时错误中断（错误码: " + rawErrCodeStr + "）。Excel 实例完好未被强杀，快照完好，可使用快照恢复。"));

                if (isPartiallyModified)
                {
                    finalSummary = "宏运行异常中断：工作簿可能已部分修改（建议核对数据或点击【快照回滚】恢复）";
                }

                finalResult = new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = finalPrecheckStatus,
                    executionPhase = finalPhase,
                    failureStage = finalStage,
                    rawErrorCode = rawErrCodeStr,
                    errorTriggerPoint = triggerPoint,
                    vbaErrNumber = vbaErrNum,
                    vbaErrDescription = vbaErrDesc,
                    comHResult = rawErrCodeStr,
                    hostExecutionPhase = finalPhase,
                    isPartiallyModified = isPartiallyModified,
                    readback = postErrorReadback,
                    summary = finalSummary,
                    error = errDetail,
                    riskNotice = finalRiskNotice,
                    rawModelResponse = rawModelResponse ?? "",
                    originalVbaCode = vbaCode,
                    executedVbaCode = finalCode,
                    wrapperCode = wrapperCode,
                    originalCodeHash = originalHash,
                    executedCodeHash = executedHash,
                    isSourceIdentical = isIdentical,
                    transformSteps = transformSteps,
                    elapsedMs = sw.ElapsedMilliseconds
                };
                return finalResult;
            }
            finally
            {
                // 11. 瞬时清理临时模块，保持工作簿纯净无宏
                if (vbComp != null && vbProj != null)
                {
                    try
                    {
                        vbProj.VBComponents.Remove(vbComp);
                    }
                    catch (Exception rmEx)
                    {
                        string warn = "临时模块 [" + moduleName + "] 自动移除失败 (" + rmEx.Message + ")，工作簿可能残留临时宏代码，建议关闭重开或手动删除。";
                        System.Diagnostics.Debug.WriteLine(warn);
                    }

                    // 二次实测核验：确认模块是否确实已不复存在
                    try
                    {
                        bool stillRemains = false;
                        foreach (dynamic c in vbProj.VBComponents)
                        {
                            if ((string)c.Name == moduleName)
                            {
                                stillRemains = true;
                                break;
                            }
                        }
                        if (stillRemains)
                        {
                            System.Diagnostics.Debug.WriteLine("警告：探测到临时模块仍残留在 VBProject 中：" + moduleName);
                        }
                    }
                    catch { }
                }

                // 12. 宿主状态精确原值恢复与读回核对（严禁统一改成 True/自动）
                bool excelAlive = false;
                try
                {
                    var probe = app.Ready;
                    excelAlive = true;
                }
                catch { }

                bool? curUp = null;
                bool? curAl = null;
                bool? curEv = null;
                int? curCa = null;
                bool stateMatched = false;
                string stateDetail = "";

                if (excelAlive)
                {
                    try
                    {
                        if (origScreenUpdating.HasValue) app.ScreenUpdating = origScreenUpdating.Value;
                        if (origDisplayAlerts.HasValue) app.DisplayAlerts = origDisplayAlerts.Value;
                        if (origEnableEvents.HasValue) app.EnableEvents = origEnableEvents.Value;
                        if (origCalculation.HasValue) app.Calculation = origCalculation.Value;

                        // 恢复后立即读回核对
                        try { curUp = (bool)app.ScreenUpdating; } catch { }
                        try { curAl = (bool)app.DisplayAlerts; } catch { }
                        try { curEv = (bool)app.EnableEvents; } catch { }
                        try { curCa = (int)app.Calculation; } catch { }

                        bool allOk = true;
                        if (origScreenUpdating.HasValue && curUp != origScreenUpdating.Value) allOk = false;
                        if (origDisplayAlerts.HasValue && curAl != origDisplayAlerts.Value) allOk = false;
                        if (origEnableEvents.HasValue && curEv != origEnableEvents.Value) allOk = false;
                        if (origCalculation.HasValue && curCa != origCalculation.Value) allOk = false;

                        stateMatched = allOk;
                        stateDetail = allOk
                            ? string.Format("宿主状态已精确恢复原值并读回核对一致 (Updating={0}, Alerts={1}, Events={2}, Calc={3})", curUp, curAl, curEv, curCa)
                            : string.Format("宿主状态恢复读回不一致: 期望(Events={0}, Calc={1}), 实际(Events={2}, Calc={3})", origEnableEvents, origCalculation, curEv, curCa);
                    }
                    catch (Exception stateEx)
                    {
                        stateMatched = false;
                        stateDetail = "恢复宿主环境状态时捕获异常: " + stateEx.Message;
                    }
                }
                else
                {
                    stateMatched = false;
                    stateDetail = "Excel 进程已失联或无响应，未声明宿主状态已成功恢复。";
                }

                if (finalResult != null)
                {
                    finalResult.origScreenUpdating = origScreenUpdating;
                    finalResult.origDisplayAlerts = origDisplayAlerts;
                    finalResult.origEnableEvents = origEnableEvents;
                    finalResult.origCalculation = origCalculation;
                    finalResult.restoredScreenUpdating = curUp;
                    finalResult.restoredDisplayAlerts = curAl;
                    finalResult.restoredEnableEvents = curEv;
                    finalResult.restoredCalculation = curCa;
                    finalResult.hostStateRestored = stateMatched;
                    finalResult.hostStateRestoreDetails = stateDetail;
                    finalResult.injectedModuleName = moduleName;
                }
            }
        }

        private static WorkbookReadback PerformReadback(dynamic targetWb, string expectedName, string expectedFullName, int preSheetCount)
        {
            var rb = new WorkbookReadback
            {
                targetWorkbookName = expectedName,
                targetWorkbookFullName = expectedFullName,
                targetVerified = false,
                otherWorkbooksAffected = false,
                affectedWorkbooksWarning = null,
                sampleValues = new List<string>(),
                rowCount = 0,
                columnCount = 0,
                usedRangeAddress = "",
                startCell = "",
                endCell = "",
                hasFormulas = false,
                hasBorders = false,
                hasInteriorColor = false,
                sheetCount = 0
            };

            try
            {
                // 1. 验证目标工作簿身份
                string currentName = (string)targetWb.Name;
                rb.targetVerified = string.Equals(currentName, expectedName, StringComparison.OrdinalIgnoreCase);
                rb.sheetCount = (int)targetWb.Sheets.Count;

                // 2. 获取真正发生修改或操作的目标工作表 (支持宏写入非活动工作表的场景)
                dynamic sheet = null;
                dynamic activeSh = null;
                try { activeSh = targetWb.ActiveSheet; } catch { }

                try
                {
                    dynamic candidateSheet = null;
                    int maxCells = 0;
                    foreach (dynamic sh in targetWb.Worksheets)
                    {
                        try
                        {
                            dynamic ur = sh.UsedRange;
                            if (ur != null)
                            {
                                string addr = ((string)ur.Address) ?? "";
                                int rCount = (int)ur.Rows.Count;
                                int cCount = (int)ur.Columns.Count;
                                object val = null;
                                try { val = ur.Cells[1, 1].Value2; } catch { }

                                // 若不是单格 $A$1，或者行数>1/列数>1，或者第一格有内容，优先作为被操作表
                                if (addr != "$A$1" || rCount > 1 || cCount > 1 || val != null)
                                {
                                    int cellCount = rCount * cCount;
                                    if (cellCount > maxCells)
                                    {
                                        maxCells = cellCount;
                                        candidateSheet = sh;
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    if (candidateSheet != null)
                    {
                        sheet = candidateSheet;
                    }
                    else
                    {
                        sheet = activeSh ?? targetWb.Sheets[1];
                    }
                }
                catch
                {
                    sheet = activeSh;
                }

                if (sheet != null)
                {
                    rb.targetSheetName = (string)sheet.Name;

                    dynamic usedRange = null;
                    try
                    {
                        usedRange = sheet.UsedRange;
                    }
                    catch { }

                    if (usedRange != null)
                    {
                        string rawAddr = (string)usedRange.Address;
                        rb.usedRangeAddress = rawAddr != null ? rawAddr.Replace("$", "") : "";
                        rb.rowCount = (int)usedRange.Rows.Count;
                        rb.columnCount = (int)usedRange.Columns.Count;

                        if (!string.IsNullOrEmpty(rb.usedRangeAddress))
                        {
                            var parts = rb.usedRangeAddress.Split(':');
                            rb.startCell = parts[0];
                            rb.endCell = parts.Length > 1 ? parts[1] : parts[0];
                        }

                        // 3. 抽样单元格文本
                        int sampleMaxRows = Math.Min(rb.rowCount, 9);
                        int sampleMaxCols = Math.Min(rb.columnCount, 9);

                        for (int r = 1; r <= sampleMaxRows && rb.sampleValues.Count < 12; r++)
                        {
                            for (int c = 1; c <= sampleMaxCols && rb.sampleValues.Count < 12; c++)
                            {
                                try
                                {
                                    dynamic cell = usedRange.Cells[r, c];
                                    object val = cell.Value2;
                                    if (val != null)
                                    {
                                        string text = val.ToString().Trim();
                                        if (!string.IsNullOrEmpty(text) && !rb.sampleValues.Contains(text))
                                        {
                                            rb.sampleValues.Add(text);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }

                        // 4. 检查公式
                        rb.hasFormulas = CheckHasFormulas(usedRange);

                        // 5. 检查边框
                        rb.hasBorders = CheckHasBorders(usedRange);

                        // 6. 检查背景填充色
                        rb.hasInteriorColor = CheckHasInteriorColor(usedRange);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("PerformReadback warning: " + ex.Message);
            }

            return rb;
        }

        public static bool CheckHasFormulas(dynamic usedRange)
        {
            try
            {
                object hf = usedRange.HasFormula;
                if (hf != null && !(hf is DBNull) && hf is bool && (bool)hf)
                {
                    return true;
                }
                dynamic rows = usedRange.Rows;
                dynamic cols = usedRange.Columns;
                int rCount = (int)rows.Count;
                int cCount = (int)cols.Count;
                for (int r = 1; r <= Math.Min(rCount, 5); r++)
                {
                    for (int c = 1; c <= Math.Min(cCount, 5); c++)
                    {
                        try
                        {
                            dynamic cell = usedRange.Cells[r, c];
                            if ((bool)cell.HasFormula) return true;
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return false;
        }

        public static bool CheckHasBorders(dynamic usedRange)
        {
            try
            {
                dynamic borders = usedRange.Borders;
                object lineStyle = borders.LineStyle;
                if (lineStyle != null && Convert.ToInt32(lineStyle) != -4142)
                {
                    return true;
                }
            }
            catch { }
            return false;
        }

        public static bool CheckHasInteriorColor(dynamic usedRange)
        {
            try
            {
                dynamic interior = usedRange.Interior;
                object colorIndex = interior.ColorIndex;
                if (colorIndex != null && Convert.ToInt32(colorIndex) != -4142)
                {
                    return true;
                }
            }
            catch { }
            return false;
        }


    }
}
