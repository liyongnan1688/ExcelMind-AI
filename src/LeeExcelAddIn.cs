using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace LeeExcel
{
    [ComVisible(true)]
    public class LeeExcelRibbon : ExcelRibbon
    {
        private static IRibbonUI _ribbon;

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public static void InvalidateRibbon()
        {
            try
            {
                if (_ribbon != null)
                {
                    _ribbon.Invalidate();
                }
            }
            catch { }
        }

        public override string GetCustomUI(string ribbonID)
        {
            return @"<customUI onLoad='OnRibbonLoad' xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='LeeExcelTab' label='AI 助手'>
        <group id='GroupAssistant' label='助手'>
          <button id='btnAssistant' size='large' getLabel='GetAssistantLabel' getImage='GetIcon'
                  keytip='W' onAction='OnTogglePane'
                  screentip='AI 助手' supertip='用中文描述，AI 直接改表' />
        </group>
        <group id='GroupMacro' label='宏'>
          <dynamicMenu id='menuFavoriteMacros' size='normal' label='收藏宏' getImage='GetIcon'
                       getContent='GetFavoriteMacrosContent' keytip='F'
                       screentip='收藏宏' supertip='一键再跑收藏的宏' />
          <button id='btnMacroLibrary' size='normal' label='宏库' getImage='GetIcon' keytip='L'
                  onAction='OnOpenMacroLibrary' screentip='宏库' supertip='查看、编辑、导入导出宏' />
        </group>
        <group id='GroupTasks' label='工具与任务'>
          <button id='btnDataTools' size='normal' label='工具' getImage='GetIcon' keytip='T'
                  onAction='OnOpenDataTools' screentip='数据工具' supertip='去重、对账、汇总、图表、导入' />
          <button id='btnBatchModal' size='normal' label='批量处理' getImage='GetIcon' keytip='B'
                  onAction='OnOpenBatch' screentip='批量处理' supertip='多个工作簿跑同一个宏' />
          <button id='btnWorkflowModal' size='normal' label='流水线' getImage='GetIcon' keytip='P'
                  onAction='OnOpenWorkflow' screentip='流水线' supertip='两步串起来自动跑' />
        </group>
        <group id='GroupSettings' label='设置'>
          <button id='btnApiSettings' size='normal' label='设置' getImage='GetIcon' keytip='S'
                  onAction='OnOpenSettings' screentip='设置' supertip='模型服务与参数' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public string GetAssistantLabel(IRibbonControl control)
        {
            return LeeExcelAddIn.IsTaskPaneVisible ? "收起工作台" : "打开工作台";
        }

        public System.Drawing.Bitmap GetIcon(IRibbonControl control)
        {
            try
            {
                if (control == null || string.IsNullOrEmpty(control.Id)) return null;
                if (control.Id == "btnAssistant")
                {
                    return BrandIconHelper.GetExcelMindIcon(32);
                }
                if (control.Id.StartsWith("fav_btn_"))
                {
                    return BrandIconHelper.GetSmallIcon("fav_run");
                }
                return BrandIconHelper.GetSmallIcon(control.Id);
            }
            catch
            {
                return null;
            }
        }

        public System.Drawing.Bitmap GetButtonImage(IRibbonControl control)
        {
            return GetIcon(control);
        }

        public string GetFavoriteMacrosContent(IRibbonControl control)
        {
            try
            {
                var all = ScriptManager.ListScripts();
                var favs = new System.Collections.Generic.List<ScriptInfo>();
                if (all != null)
                {
                    foreach (var s in all)
                    {
                        if (s.isFavorite) favs.Add(s);
                    }
                }

                var sb = new System.Text.StringBuilder();
                sb.Append("<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>");

                if (favs.Count == 0)
                {
                    sb.Append("<button id='fav_empty' label='（暂无收藏宏 - 请在宏库中点击⭐收藏）' enabled='false' getImage='GetIcon' />");
                }
                else
                {
                    for (int i = 0; i < favs.Count; i++)
                    {
                        var s = favs[i];
                        string cleanBtnId = "fav_btn_" + i;
                        string safeId = EscapeXml(s.id ?? s.fileName ?? "");
                        string safeLabel = EscapeXml(s.displayName ?? s.name ?? "未命名宏");
                        string tip = EscapeXml(string.Format("入口: {0} | 分类: {1}", s.entryPoint ?? "默认", s.category ?? "未分类"));

                        sb.Append(string.Format(
                            "<button id='{0}' tag='{1}' label='{2}' screentip='{2}' supertip='{3}' getImage='GetIcon' onAction='OnExecuteFavoriteMacro' />",
                            cleanBtnId, safeId, safeLabel, tip));
                    }
                }

                sb.Append("<menuSeparator id='fav_sep_manage' />");
                sb.Append("<button id='fav_open_lib' label='打开宏库管理...' getImage='GetIcon' onAction='OnOpenMacroLibrary' />");
                sb.Append("</menu>");

                return sb.ToString();
            }
            catch
            {
                return "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'><button id='fav_err' label='加载收藏宏失败' enabled='false' /></menu>";
            }
        }

        private static string EscapeXml(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            return str.Replace("&", "&amp;")
                      .Replace("<", "&lt;")
                      .Replace(">", "&gt;")
                      .Replace("\"", "&quot;")
                      .Replace("'", "&apos;");
        }

        public void OnExecuteFavoriteMacro(IRibbonControl control)
        {
            string scriptId = control.Tag;
            if (!string.IsNullOrEmpty(scriptId))
            {
                LeeExcelAddIn.PromptRunFavoriteMacro(scriptId);
            }
        }

        public void OnTogglePane(IRibbonControl control)
        {
            LeeExcelAddIn.ToggleTaskPane();
            InvalidateRibbon();
        }

        public void OnOpenMacroLibrary(IRibbonControl control)
        {
            LeeExcelAddIn.OpenMacroLibrary();
        }

        public void OnImportMacro(IRibbonControl control)
        {
            LeeExcelAddIn.OpenImportMacro();
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            LeeExcelAddIn.OpenSettings();
        }

        public void OnOpenDataTools(IRibbonControl control)
        {
            LeeExcelAddIn.OpenDataTools();
        }

        public void OnOpenBatch(IRibbonControl control)
        {
            LeeExcelAddIn.OpenBatch();
        }

        public void OnOpenWorkflow(IRibbonControl control)
        {
            LeeExcelAddIn.OpenWorkflow();
        }

        // 向后兼容旧入口
        public void OnOpenScripts(IRibbonControl control)
        {
            LeeExcelAddIn.OpenMacroLibrary();
        }

        public void OnMyMacros(IRibbonControl control)
        {
            LeeExcelAddIn.OpenMacroLibrary();
        }
    }

    public class LeeExcelAddIn : IExcelAddIn
    {
        private static CustomTaskPane _customTaskPane;
        private static TaskPaneControl _taskPaneControl;

        public void AutoOpen()
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                _taskPaneControl = new TaskPaneControl(app);

                _customTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_taskPaneControl, "ExcelMind AI");
                _customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _customTaskPane.Width = 440;
                _customTaskPane.Visible = true;

                // 监听工作簿切换、新建、打开
                try
                {
                    Action<object> notify = (wb) =>
                    {
                        if (_taskPaneControl != null) _taskPaneControl.NotifyWorkbookChanged();
                    };

                    Action relinquishFocus = () =>
                    {
                        if (_taskPaneControl != null) _taskPaneControl.RelinquishFocusToExcel();
                    };

                    try { app.WorkbookActivate += notify; } catch { }
                    try { app.WorkbookOpen += notify; } catch { }
                    try { app.NewWorkbook += notify; } catch { }
                    try { app.WindowActivate += new Action<object, object>((wb, wn) => notify(wb)); } catch { }
                    try { app.WorkbookAfterSave += new Action<object, bool>((wb, success) => notify(wb)); } catch { }
                    try { app.SheetActivate += new Action<object>((sh) => notify(null)); } catch { }

                    // 工作表交互双重保险：当用户在工作表触发选区或编辑操作时，若焦点仍残留于任务窗格，平滑归还给工作表
                    try { app.SheetSelectionChange += new Action<object, object>((sh, target) => relinquishFocus()); } catch { }
                    try { app.SheetBeforeDoubleClick += new Action<object, object, bool>((sh, target, cancel) => relinquishFocus()); } catch { }
                    try { app.SheetBeforeRightClick += new Action<object, object, bool>((sh, target, cancel) => relinquishFocus()); } catch { }
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("ExcelMind AI 启动异常: " + ex.Message);
            }
        }

        public void AutoClose()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = false;
                _customTaskPane.Delete();
                _customTaskPane = null;
            }
        }

        public static bool IsTaskPaneVisible
        {
            get
            {
                try
                {
                    return _customTaskPane != null && _customTaskPane.Visible;
                }
                catch
                {
                    return false;
                }
            }
        }

        public static void EnsureTaskPaneVisible()
        {
            if (_customTaskPane != null)
            {
                if (!_customTaskPane.Visible)
                {
                    _customTaskPane.Visible = true;
                    LeeExcelRibbon.InvalidateRibbon();
                }
                _customTaskPane.Width = 440;
                if (_taskPaneControl != null) _taskPaneControl.NotifyWorkbookChanged();
            }
        }

        public static void ToggleTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = !_customTaskPane.Visible;
                LeeExcelRibbon.InvalidateRibbon();
                if (_customTaskPane.Visible)
                {
                    _customTaskPane.Width = 440;
                    if (_taskPaneControl != null) _taskPaneControl.NotifyWorkbookChanged();
                }
            }
        }

        public static void OpenMacroLibrary()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenMacroLibrary();
            }
        }

        public static void OpenScripts()
        {
            OpenMacroLibrary();
        }

        public static void OpenSettings()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenSettings();
            }
        }

        public static void OpenImportMacro()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenImportMacro();
            }
        }

        public static void OpenMyMacros()
        {
            OpenMacroLibrary();
        }

        public static void PromptRunFavoriteMacro(string scriptId)
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.PromptRunFavoriteMacro(scriptId);
            }
        }

        public static void OpenDataTools()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenDataTools();
            }
        }

        public static void OpenBatch()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenBatch();
            }
        }

        public static void OpenWorkflow()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenWorkflow();
            }
        }
    }
}
