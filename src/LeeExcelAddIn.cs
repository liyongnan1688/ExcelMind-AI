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
        <group id='LeeExcelGroup' label='智能操作'>
          <button id='btnTogglePane' label='ExcelMind AI' size='large' onAction='OnTogglePane' getImage='GetButtonImage' />
          <button id='btnOpenScripts' label='自动化脚本' size='large' onAction='OnOpenScripts' imageMso='VisualBasic' />
          <button id='btnOpenSettings' label='API配置' size='large' onAction='OnOpenSettings' imageMso='ServerConnection' />
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

        public void OnOpenScripts(IRibbonControl control)
        {
            LeeExcelAddIn.OpenScripts();
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            LeeExcelAddIn.OpenSettings();
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

        public static void OpenScripts()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenScripts();
            }
        }

        public static void OpenSettings()
        {
            EnsureTaskPaneVisible();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.OpenSettings();
            }
        }
    }
}
