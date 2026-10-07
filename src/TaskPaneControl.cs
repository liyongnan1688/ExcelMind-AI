using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace LeeExcel
{
    public class TaskPaneControl : UserControl
    {
        private WebView2 _webView;
        private dynamic _app;
        private string _pendingAction = null;
        private TaskPaneFocusMessageFilter _focusFilter = null;

        public TaskPaneControl(dynamic excelApp)
        {
            _app = excelApp;
            Dock = DockStyle.Fill;

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_webView);

            InitializeWebView();

            // 注册跨窗格/工作表平滑焦点交接过滤器
            try
            {
                _focusFilter = new TaskPaneFocusMessageFilter(this.Handle, _app);
                Application.AddMessageFilter(_focusFilter);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("AddMessageFilter notice: " + ex.Message);
            }
        }

        public void RelinquishFocusToExcel()
        {
            if (this.IsHandleCreated)
            {
                TaskPaneFocusHelper.RelinquishFocusToExcel(this.Handle, _app);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_focusFilter != null)
                {
                    try
                    {
                        Application.RemoveMessageFilter(_focusFilter);
                    }
                    catch { }
                    _focusFilter = null;
                }
            }
            base.Dispose(disposing);
        }

        private async void InitializeWebView()
        {
            try
            {
                // 1. 安全获取插件基础目录 (Excel-DNA 下 Assembly.Location 可能为空，优先使用 XllPath)
                string baseDir = null;
                try
                {
                    string xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                    if (!string.IsNullOrEmpty(xllPath))
                    {
                        baseDir = Path.GetDirectoryName(xllPath);
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(baseDir))
                {
                    try
                    {
                        string loc = Assembly.GetExecutingAssembly().Location;
                        if (!string.IsNullOrEmpty(loc))
                        {
                            baseDir = Path.GetDirectoryName(loc);
                        }
                    }
                    catch { }
                }

                if (string.IsNullOrEmpty(baseDir))
                {
                    baseDir = AppDomain.CurrentDomain.BaseDirectory;
                }

                // 2. 双架构 WebView2Loader.dll 动态适配 (x86 与 x64)
                if (!string.IsNullOrEmpty(baseDir))
                {
                    try
                    {
                        string archFolder = (IntPtr.Size == 8) ? "win-x64" : "win-x86";
                        string nativeLoaderDir = Path.Combine(baseDir, "runtimes", archFolder, "native");
                        if (Directory.Exists(nativeLoaderDir) && File.Exists(Path.Combine(nativeLoaderDir, "WebView2Loader.dll")))
                        {
                            CoreWebView2Environment.SetLoaderDllFolderPath(nativeLoaderDir);
                        }
                    }
                    catch (Exception loaderEx)
                    {
                        System.Diagnostics.Debug.WriteLine("SetLoaderDllFolderPath notice: " + loaderEx.Message);
                    }
                }

                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "LeeExcel",
                    "WebView2Profile"
                );
                if (!Directory.Exists(userDataFolder))
                {
                    Directory.CreateDirectory(userDataFolder);
                }

                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(env);

                _webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.NavigationCompleted += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(_pendingAction))
                    {
                        try
                        {
                            if (_pendingAction.StartsWith("{"))
                            {
                                _webView.CoreWebView2.PostWebMessageAsString(_pendingAction);
                            }
                            else
                            {
                                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"" + _pendingAction + "\"}");
                            }
                        }
                        catch { }
                        _pendingAction = null;
                    }
                };

                // 2. 查找前端打包目录并使用虚拟域名映射 (彻底解决 file:// 下 ES Module CORS 限制)
                string distFolder = Path.Combine(baseDir, "dist");
                if (!Directory.Exists(distFolder))
                {
                    try
                    {
                        string altDist = Path.GetFullPath(Path.Combine(baseDir, "..", "web", "dist"));
                        if (Directory.Exists(altDist))
                        {
                            distFolder = altDist;
                        }
                    }
                    catch { }
                }

                if (Directory.Exists(distFolder) && File.Exists(Path.Combine(distFolder, "index.html")))
                {
                    try
                    {
                        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "app.lee-excel",
                            distFolder,
                            CoreWebView2HostResourceAccessKind.Allow
                        );
                        _webView.CoreWebView2.Navigate("https://app.lee-excel/index.html");
                        return;
                    }
                    catch (Exception mapEx)
                    {
                        System.Diagnostics.Debug.WriteLine("VirtualHost mapping warning: " + mapEx.Message);
                    }

                    // 备用 file 协议导航
                    string fileUri = new Uri(Path.Combine(distFolder, "index.html")).AbsoluteUri;
                    _webView.CoreWebView2.Navigate(fileUri);
                }
                else
                {
                    // 若尚未检测到前端文件，展示优雅的初始化就绪引导页
                    string fallbackHtml = @"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'/>
    <style>
        body { font-family: 'Segoe UI', 微软雅黑, sans-serif; padding: 24px; color: #201F1E; background: #F3F2F1; }
        .card { background: white; padding: 20px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); border-left: 4px solid #107C41; }
        h2 { margin-top: 0; color: #107C41; }
        p { line-height: 1.5; color: #605E5C; }
        .status { display: inline-block; padding: 4px 8px; background: #E7F3EC; color: #107C41; border-radius: 4px; font-weight: 600; font-size: 12px; }
    </style>
</head>
<body>
    <div class='card'>
        <h2>ExcelMind AI</h2>
        <span class='status'>原生加载项 C# 与 WebView2 就绪</span>
        <p>WebView2 引擎与 Excel COM 宿主连接成功。<br/>正在编译加载前端 Office 经典绿界面...</p>
    </div>
</body>
</html>";
                    _webView.CoreWebView2.NavigateToString(fallbackHtml);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("WebView2 初始化异常: " + ex.Message, "ExcelMind AI 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string msg = null;
                try { msg = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrEmpty(msg))
                {
                    try { msg = e.WebMessageAsJson; } catch { }
                }
                if (string.IsNullOrEmpty(msg)) return;

                Console.WriteLine("    [TaskPane 收到前端 WebMessage] action=" + (msg.Contains("execute_vba") ? "execute_vba" : (msg.Contains("restore_snapshot") ? "restore_snapshot" : "other")));
                string reply = NativeBridge.Dispatch(msg, _app);
                Console.WriteLine("    [TaskPane 宿主处理完成，向 WebView2 发送回复] action=" + (reply.Contains("execute_vba") ? "execute_vba" : (reply.Contains("restore_snapshot") ? "restore_snapshot" : "other")));
                if (_webView != null && _webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsString(reply);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("    [TaskPane 异常] " + ex.Message);
                System.Diagnostics.Debug.WriteLine("OnWebMessageReceived error: " + ex.Message);
            }
        }

        public void NotifyWorkbookChanged()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(NotifyWorkbookChanged));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                string info = NativeBridge.Dispatch("{\"action\":\"get_workbook_info\"}", _app);
                _webView.CoreWebView2.PostWebMessageAsString(info);
            }
        }

        public void OpenMacroLibrary()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenMacroLibrary));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_macro_library\"}");
            }
            else
            {
                _pendingAction = "open_macro_library";
            }
        }

        public void OpenScripts()
        {
            OpenMacroLibrary();
        }

        public void OpenSettings()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenSettings));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_settings\"}");
            }
            else
            {
                _pendingAction = "open_settings";
            }
        }

        public void OpenImportMacro()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenImportMacro));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_import_macro\"}");
            }
            else
            {
                _pendingAction = "open_import_macro";
            }
        }

        public void OpenMyMacros()
        {
            OpenMacroLibrary();
        }

        public void PromptRunFavoriteMacro(string scriptId)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action<string>(PromptRunFavoriteMacro), scriptId);
                }
                catch { }
                return;
            }

            string payload = "{\"action\":\"prompt_run_macro\",\"scriptId\":\"" + (scriptId ?? "") + "\"}";
            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString(payload);
            }
            else
            {
                _pendingAction = payload;
            }
        }

        public void OpenDataTools()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenDataTools));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_data_tools\"}");
            }
            else
            {
                _pendingAction = "open_data_tools";
            }
        }

        public void OpenBatch()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenBatch));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_batch_modal\"}");
            }
            else
            {
                _pendingAction = "open_batch_modal";
            }
        }

        public void OpenWorkflow()
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(new Action(OpenWorkflow));
                }
                catch { }
                return;
            }

            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.PostWebMessageAsString("{\"action\":\"open_workflow_modal\"}");
            }
            else
            {
                _pendingAction = "open_workflow_modal";
            }
        }
    }
}
