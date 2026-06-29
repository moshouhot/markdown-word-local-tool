using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace MarkdownWordTool;

public partial class Form1 : Form
{
    private const string VirtualHostName = "app.local";
    private readonly WebView2 webView;

    public Form1()
    {
        Text = "Markdown → Word";
        MinimumSize = new Size(960, 640);
        StartPosition = FormStartPosition.CenterScreen;
        WindowState = FormWindowState.Maximized;

        webView = new WebView2
        {
            Dock = DockStyle.Fill,
            AllowExternalDrop = true,
        };
        Controls.Add(webView);

        Load += async (_, _) => await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            string appFolder = Path.Combine(AppContext.BaseDirectory, "app");
            string indexPath = Path.Combine(appFolder, "index.html");
            if (!File.Exists(indexPath))
            {
                ShowStartupError($"未找到前端入口文件：{indexPath}");
                return;
            }

            string userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MarkdownWordTool",
                "WebView2");
            Directory.CreateDirectory(userDataFolder);

            CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataFolder);

            await webView.EnsureCoreWebView2Async(environment);

            CoreWebView2 core = webView.CoreWebView2;
            core.SetVirtualHostNameToFolderMapping(
                VirtualHostName,
                appFolder,
                CoreWebView2HostResourceAccessKind.DenyCors);

            core.Settings.AreDevToolsEnabled = true;
            core.Settings.AreDefaultContextMenusEnabled = true;
            core.Settings.IsStatusBarEnabled = false;
            core.DocumentTitleChanged += (_, _) =>
            {
                if (!string.IsNullOrWhiteSpace(core.DocumentTitle))
                {
                    Text = core.DocumentTitle;
                }
            };
            core.NavigationCompleted += (_, args) =>
            {
                if (!args.IsSuccess)
                {
                    ShowStartupError($"页面加载失败：{args.WebErrorStatus}");
                }
            };
            core.NewWindowRequested += (_, args) =>
            {
                args.Handled = true;
                core.Navigate(args.Uri);
            };

            core.Navigate($"https://{VirtualHostName}/index.html");
        }
        catch (WebView2RuntimeNotFoundException ex)
        {
            ShowStartupError("未检测到 Microsoft Edge WebView2 Runtime。请安装 WebView2 Runtime 后重试。\n\n" + ex.Message);
        }
        catch (Exception ex)
        {
            ShowStartupError("启动 Markdown → Word 工具失败。\n\n" + ex.Message);
        }
    }

    private static void ShowStartupError(string message)
    {
        MessageBox.Show(message, "Markdown → Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
