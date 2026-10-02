using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace LeeExcel
{
    [ComVisible(true)]
    public class LeeExcelRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='LeeExcelTab' label='ExcelMind AI'>
        <!-- 分组 1：智能助手 -->
        <group id='GroupAssistant' label='智能助手'>
          <button id='btnAssistant'
                  label='ExcelMind AI'
                  size='large'
                  onAction='OnTogglePane'
                  getImage='GetButtonImage'
                  screentip='ExcelMind AI'
                  supertip='打开或收起 ExcelMind AI 侧边栏，支持自然语言对话、公式生成与工作簿自动化操作。' />
        </group>

        <!-- 分组 2：宏工具 -->
        <group id='GroupMacroTools' label='宏工具'>
          <button id='btnMacroLibrary'
                  label='宏库'
                  size='large'
                  onAction='OnOpenMacroLibrary'
                  imageMso='VisualBasic'
                  screentip='宏库'
                  supertip='查看、搜索、运行、重命名、导出和管理所有已保存的 VBA 宏与自动化脚本。' />
          <button id='btnImportMacro'
                  label='导入'
                  size='large'
                  onAction='OnImportMacro'
                  imageMso='ImportTextFile'
                  screentip='导入'
                  supertip='从本地 .bas/.vba/.txt 文件导入或直接粘贴 VBA 源码，命名保存至宏库。' />
        </group>

        <!-- 分组 3：设置 -->
        <group id='GroupSettings' label='设置'>
          <button id='btnApiSettings'
                  label='API设置'
                  size='large'
                  onAction='OnOpenSettings'
                  imageMso='ServerConnection'
                  screentip='API设置'
                  supertip='配置大模型 API 密钥、接口地址及模型服务参数。' />
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

                    try { app.WorkbookActivate += notify; } catch { }
                    try { app.WorkbookOpen += notify; } catch { }
                    try { app.NewWorkbook += notify; } catch { }
                    try { app.WindowActivate += new Action<object, object>((wb, wn) => notify(wb)); } catch { }
                    try { app.WorkbookAfterSave += new Action<object, bool>((wb, success) => notify(wb)); } catch { }
                    try { app.SheetActivate += new Action<object>((sh) => notify(null)); } catch { }
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
    }
}
