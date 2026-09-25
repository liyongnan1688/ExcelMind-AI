using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;

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

            // 1. 结构完整性校验：必须包含过程定义与 End Sub (防止模型响应被截断导致编译报错)
            var subMatches = Regex.Matches(vbaCode, @"(?:Public\s+|Private\s+)?Sub\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*(?:\((.*?)\))?", RegexOptions.IgnoreCase);
            bool hasEndSub = Regex.IsMatch(vbaCode, @"End\s+Sub", RegexOptions.IgnoreCase);

            if (subMatches.Count == 0 || !hasEndSub)
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：VBA 代码不完整 (缺少 Sub 或 End Sub)",
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

            // 2. 识别主入口过程与参数特征
            // 扫描是否存在显式声明接收 Workbook 对象的入口过程
            Match targetSubMatch = null;
            foreach (Match m in subMatches)
            {
                string p = m.Groups[2].Value;
                if (Regex.IsMatch(p, @"(?:targetWb|wb|workbook)\s+As\s+(?:Workbook|Object)", RegexOptions.IgnoreCase))
                {
                    targetSubMatch = m;
                    break;
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
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 3. 通用透明调度架构：绝不使用正则篡改模型正文源码
            string callMacroName = "";
            string wrapperCode = null;
            string finalCode = vbaCode;

            if (targetSubMatch != null)
            {
                // 模型过程原生声明了接收目标工作簿参数，100% 原始源码直调
                callMacroName = targetSubMatch.Groups[1].Value;
                finalCode = vbaCode;
                wrapperCode = null;
                transformSteps.Add("直调主入口: " + callMacroName + " (模型原生接收目标工作簿参数，源码 100% 原始直调)");
            }
            else
            {
                // 模型生成了标准通用无参主过程或多过程结构
                // 原则：模型正文源码保持 100% 零修改，仅在模块尾部追加受控的透明调用入口包装器
                string mainSubName = subMatches[0].Groups[1].Value;
                string wrapperSubName = "LeeHostRunner_" + DateTime.Now.ToString("mmss");
                callMacroName = wrapperSubName;

                wrapperCode =
                    "\r\n' ===== [LeeExcel 自动生成的受控调用入口包装器 - 保持模型正文源码零篡改] =====\r\n" +
                    "Sub " + wrapperSubName + "(targetWb As Workbook)\r\n" +
                    "    targetWb.Activate\r\n" +
                    "    Call " + mainSubName + "\r\n" +
                    "End Sub";

                finalCode = vbaCode + "\r\n" + wrapperCode;
                transformSteps.Add("模型源码正文 100% 保持原貌，独立追加受控入口包装器: " + wrapperSubName + " -> 调度主过程 " + mainSubName);
            }

            string originalHash = ComputeSha256(vbaCode);
            string executedHash = ComputeSha256(finalCode);
            bool isIdentical = string.Equals(vbaCode, finalCode, StringComparison.Ordinal);

            // 4. 执行前：记录当前 Excel 实例中所有【非目标工作簿】的初始状态快照
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
            bool screenUpdated = false;
            string moduleName = "";
            string precheckStatus = "unavailable";

            try
            {
                // 5. 检查 VBProject 访问权限
                try
                {
                    vbProj = targetWorkbook.VBProject;
                }
                catch (COMException comEx)
                {
                    if ((uint)comEx.ErrorCode == 0x800A03EC || comEx.Message.Contains("1004"))
                    {
                        return new VbaExecutionResult
                        {
                            success = false,
                            precheckStatus = "unavailable",
                            executionPhase = "intercepted_before_run",
                            summary = "执行中断：Excel 未开启对 VBA 工程对象模型的访问信任",
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
                    }
                    throw;
                }

                // 6. 动态创建标准临时模块并注入宏
                moduleName = "LeeMod_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                vbComp = vbProj.VBComponents.Add(1);
                vbComp.Name = moduleName;
                vbComp.CodeModule.AddFromString(finalCode);

                // 7. 预编译安全探测：精确区分语法预检通过、失败与命令不可用
                string precheckDetail = "";
                bool compileIntercepted = false;

                try
                {
                    try { vbComp.Activate(); } catch { }
                    dynamic compileBtn = app.VBE.CommandBars.FindControl(Type.Missing, 578);
                    if (compileBtn != null)
                    {
                        if ((bool)compileBtn.Enabled)
                        {
                            compileBtn.Execute();
                            if ((bool)compileBtn.Enabled)
                            {
                                compileIntercepted = true;
                                precheckStatus = "failed";
                                precheckDetail = "VBE 预编译未通过：代码中存在未定义符号/变量、缺少类型引用、语法拼写错误或括号不匹配。";
                            }
                            else
                            {
                                precheckStatus = "passed";
                                precheckDetail = "VBE 预编译探测通过（命令 578 执行成功且状态重置就绪）。静态语法与对象模型校验通过。";
                            }
                        }
                        else
                        {
                            precheckStatus = "passed";
                            precheckDetail = "VBE 当前无未编译变更，工程静态语法状态正常。";
                        }
                    }
                    else
                    {
                        precheckStatus = "unavailable";
                        precheckDetail = "VBE 预编译命令(ID 578)当前不可用（可能受语言包或安全策略限制），未执行静态语法预检。";
                    }
                }
                catch (Exception cEx)
                {
                    precheckStatus = "unavailable";
                    precheckDetail = "执行 VBE 预编译探测时抛出异常: " + cEx.Message + "，未执行静态语法预检。";
                }

                if (compileIntercepted)
                {
                    return new VbaExecutionResult
                    {
                        success = false,
                        precheckStatus = "failed",
                        executionPhase = "syntax_failed",
                        summary = "代码预编译未通过：存在语法或引用错误，已在运行前安全拦截",
                        error = precheckDetail,
                        riskNotice = "【运行前安全拦截】代码未进入运行阶段，未对工作簿做任何更改，临时模块已安全清理。",
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
                }

                if (precheckStatus == "passed")
                {
                    transformSteps.Add("【静态语法预检】" + precheckDetail + "（警示：预编译仅消除静态语法与对象错误，不能消除除零/死循环等运行期挂起风险）");
                }
                else
                {
                    transformSteps.Add("【预编译不可用警示】" + precheckDetail + " 系统拒绝静默报告为受检执行，将在整本快照保护下执行。");
                }

                // 8. 挂起屏幕刷新与系统弹窗
                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    screenUpdated = true;
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
                    IntPtr excelHwnd = new IntPtr((int)app.Hwnd);
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

                                                if (!isUserDialog && (title.Contains("Microsoft Visual Basic") || title.Contains("Microsoft Excel")))
                                                {
                                                    // 仅当确定为宏执行期间弹出的 VBA 运行时错误弹窗时才进行受控关闭
                                                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                    return true;
                                }, IntPtr.Zero);
                            }
                            catch { }
                            System.Threading.Thread.Sleep(100);
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
                        return new VbaExecutionResult
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

                return new VbaExecutionResult
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
                    precheckStatus = precheckStatus,
                    executionPhase = "macro_completed",
                    riskNotice = precheckStatus == "passed"
                        ? "【运行完成（静态已受检）】宏已正常返回并读回数据。静态语法与对象已受检通过。（说明：快照仅能回滚目标工作簿本身，无法覆盖对其他工作簿、磁盘或网络的外部副作用）。"
                        : "【运行完成（未执行预检）】宏已正常返回并读回数据。预编译探测不可用，当前处于未受检执行。（说明：快照仅能回滚目标工作簿本身，无法覆盖对其他工作簿、磁盘或网络的外部副作用）。"
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errDetail = ex.Message;
                if (ex.InnerException != null) errDetail += " (" + ex.InnerException.Message + ")";

                if (vbProj != null)
                {
                    try
                    {
                        vbProj.VBE.MainWindow.Visible = false;
                    }
                    catch { }
                }

                return new VbaExecutionResult
                {
                    success = false,
                    precheckStatus = precheckStatus,
                    executionPhase = "runtime_error",
                    summary = "执行失败：VBA 运行期抛出异常",
                    error = errDetail,
                    riskNotice = "【运行期异常】宏在执行期间发生运行时错误中断。Excel 实例完好未被强杀，快照完好，可一键恢复。",
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
                }

                if (screenUpdated)
                {
                    try
                    {
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                    catch { }
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

                // 2. 获取当前活动工作表或新建的工作表
                dynamic sheet = null;
                try
                {
                    sheet = targetWb.ActiveSheet;
                }
                catch { }

                if (sheet == null)
                {
                    try
                    {
                        sheet = targetWb.Sheets[1];
                    }
                    catch { }
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
