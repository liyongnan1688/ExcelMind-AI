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
      <tab id='LeeExcelTab' label='AI 助手'>
        <group id='LeeExcelGroup' label='智能操作'>
          <button id='btnTogglePane' label='打开 AI 任务窗格' size='large' onAction='OnTogglePane' imageMso='FileNewBlankDocument' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnTogglePane(IRibbonControl control)
        {
            LeeExcelAddIn.ToggleTaskPane();
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

                _customTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_taskPaneControl, "Lee-Excel AI 助手");
                _customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _customTaskPane.Width = 430;
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
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("LeeExcel 启动异常: " + ex.Message);
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

        public static void ToggleTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = !_customTaskPane.Visible;
                if (_customTaskPane.Visible)
                {
                    if (_taskPaneControl != null) _taskPaneControl.NotifyWorkbookChanged();
                }
            }
        }
    }
}
