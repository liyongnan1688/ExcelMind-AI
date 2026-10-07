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
      <tab id='LeeExcelTab' label='ExcelMind AI'>
        <!-- 分组 1：ExcelMind AI -->
        <group id='GroupAssistant' label='ExcelMind AI'>
          <button id='btnAssistant'
                  size='large'
                  onAction='OnTogglePane'
                  getImage='GetButtonImage'
                  screentip='ExcelMind AI'
                  supertip='打开或收起 ExcelMind AI 工作台。' />
        </group>

        <!-- 分组 2：宏与数据 -->
        <group id='GroupMacroAndData' label='宏与数据'>
          <button id='btnMacroLibrary'
                  label='宏库'
                  size='normal'
                  onAction='OnOpenMacroLibrary'
                  imageMso='VisualBasic'
                  screentip='宏库'
                  supertip='查看、搜索、运行、重命名、导出和管理所有已保存的 VBA 宏与自动化脚本。' />
          <dynamicMenu id='menuFavoriteMacros'
                       label='收藏宏'
                       size='normal'
                       imageMso='Favorites'
                       getContent='GetFavoriteMacrosContent'
                       screentip='收藏宏'
                       supertip='快速从功能区选择并直接确认运行已收藏的常用自动化宏。' />
          <button id='btnDataTools'
                  label='数据工具'
                  size='normal'
                  onAction='OnOpenDataTools'
                  imageMso='Consolidate'
                  screentip='快捷数据工具箱'
                  supertip='打开数据去重、跨表对账、多文件汇总与免公式快捷图表工作区。' />
        </group>

        <!-- 分组 3：任务与设置 -->
        <group id='GroupTasksAndSettings' label='任务与设置'>
          <menu id='menuTasks'
                label='任务'
                size='normal'
                imageMso='DiagramProcessClassic'
                screentip='任务'
                supertip='管理跨工作簿批量宏处理与双步骤自动化流水线任务。'>
            <button id='btnBatchModal'
                    label='批量处理'
                    imageMso='AppointmentColorDialog'
                    onAction='OnOpenBatch'
                    screentip='批量处理'
                    supertip='跨多个 Excel 工作簿文件批量执行宏或自动化处理任务。' />
            <button id='btnWorkflowModal'
                    label='工作流'
                    imageMso='DiagramProcessClassic'
                    onAction='OnOpenWorkflow'
                    screentip='双步骤工作流'
                    supertip='将数据清洗与处理串联为两步流水线任务，全流程快照保护。' />
          </menu>
          <button id='btnApiSettings'
                  label='设置'
                  size='normal'
                  onAction='OnOpenSettings'
                  imageMso='ServerConnection'
                  screentip='设置与环境诊断'
                  supertip='配置大模型 API 密钥、接口地址及导出脱敏诊断包。' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public System.Drawing.Bitmap GetButtonImage(IRibbonControl control)
        {
            try
            {
                return BrandIconHelper.GetExcelMindIcon(32);
            }
            catch
            {
                return null;
            }
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
                    sb.Append("<button id='fav_empty' label='（暂无收藏宏 - 请在宏库中点击⭐收藏）' enabled='false' imageMso='Info' />");
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
                            "<button id='{0}' tag='{1}' label='{2}' screentip='{2}' supertip='{3}' imageMso='MacroPlay' onAction='OnExecuteFavoriteMacro' />",
                            cleanBtnId, safeId, safeLabel, tip));
                    }
                }

                sb.Append("<menuSeparator id='fav_sep_manage' />");
                sb.Append("<button id='fav_open_lib' label='打开宏库管理...' imageMso='VisualBasic' onAction='OnOpenMacroLibrary' />");
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

        public static void EnsureTaskPaneVisible()
        {
            if (_customTaskPane != null)
            {
                if (!_customTaskPane.Visible)
                {
                    _customTaskPane.Visible = true;
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
