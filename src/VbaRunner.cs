using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace LeeExcel
{
    public class VbaExecutionResult
    {
        public bool success { get; set; }
        public string summary { get; set; }
        public string error { get; set; }
        public string vbaCode { get; set; }
        public double elapsedMs { get; set; }
    }

    public class VbaRunner
    {
        private static string ExtractSubName(string code)
        {
            var m = Regex.Match(code, @"(?:Public\s+|Private\s+)?Sub\s+([a-zA-Z0-9_\u4e00-\u9fa5]+)\s*\(", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                return m.Groups[1].Value;
            }
            return null;
        }

        public static VbaExecutionResult RunVbaCode(dynamic app, dynamic targetWorkbook, string vbaCode)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            if (string.IsNullOrEmpty(vbaCode))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：传入的代码为空",
                    error = "未检测到有效 VBA 代码",
                    vbaCode = "",
                    elapsedMs = 0
                };
            }

            // 1. 严格校验代码是否包含完整的 Sub 与 End Sub
            if (!Regex.IsMatch(vbaCode, @"(?:Public\s+|Private\s+)?Sub\s+", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(vbaCode, @"End\s+Sub", RegexOptions.IgnoreCase))
            {
                return new VbaExecutionResult
                {
                    success = false,
                    summary = "执行中断：VBA 代码不完整 (缺少 Sub 或 End Sub)",
                    error = "模型输出的代码结构不完整，已在注入前安全拦截，避免导致 Excel 编译错误。",
                    vbaCode = vbaCode,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            // 2. 自动将 ThisWorkbook 安全规整为 ActiveWorkbook，确保加载项动态注入时精确作用于当前活动文档
            string finalCode = Regex.Replace(
                vbaCode,
                @"\bThisWorkbook\b",
                "ActiveWorkbook",
                RegexOptions.IgnoreCase
            );

            // 3. 将主过程入口标准化替换为纯英文 LeeTaskEntry，彻底根除中文过程名引发的 COM 0x800A9C68 异常
            string entrySubName = "LeeTaskEntry";
            finalCode = Regex.Replace(
                finalCode,
                @"(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(",
                "Sub " + entrySubName + "(",
                RegexOptions.IgnoreCase
            );

            dynamic vbComp = null;
            dynamic vbProj = null;
            bool screenUpdated = false;

            try
            {
                // 3. 检查 VBProject 访问权限
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
                            summary = "执行中断：Excel 未开启对 VBA 工程对象模型的访问信任",
                            error = "错误 1004：请在 Excel“文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置”中勾选【信任对 VBA 工程对象模型的访问】。",
                            vbaCode = vbaCode,
                            elapsedMs = sw.ElapsedMilliseconds
                        };
                    }
                    throw;
                }

                // 4. 动态创建标准临时模块 (1 = vbext_ct_StdModule)
                string moduleName = "LeeMod_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                vbComp = vbProj.VBComponents.Add(1);
                vbComp.Name = moduleName;
                vbComp.CodeModule.AddFromString(finalCode);

                // 5. 挂起屏幕刷新与告警
                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    screenUpdated = true;
                }
                catch { }

                // 6. 调用标准化英文宏
                string macroAddress = "'" + targetWorkbook.Name + "'!" + entrySubName;
                app.Run(macroAddress);

                sw.Stop();
                return new VbaExecutionResult
                {
                    success = true,
                    summary = "执行成功：已按指令完成当前工作簿操作",
                    error = null,
                    vbaCode = finalCode,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                string errDetail = ex.Message;
                if (ex.InnerException != null) errDetail += " (" + ex.InnerException.Message + ")";

                // 若触发了编译错误导致 VBE 弹窗，强制将 VBE 窗口隐藏
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
                    summary = "执行失败：VBA 运行期抛出异常",
                    error = errDetail,
                    vbaCode = finalCode,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }
            finally
            {
                // 7. 瞬时清理临时模块，保持工作簿纯净无宏
                if (vbComp != null && vbProj != null)
                {
                    try
                    {
                        vbProj.VBComponents.Remove(vbComp);
                    }
                    catch (Exception rmEx)
                    {
                        System.Diagnostics.Debug.WriteLine("Remove temp module warning: " + rmEx.Message);
                    }
                }

                // 恢复屏幕刷新
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
    }
}
