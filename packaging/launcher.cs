using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("表格工具")]
[assembly: AssemblyProduct("表格检索 / 模板改造 / 数据核对")]
[assembly: AssemblyCompany("杨子江 by DeepSeek HARNESS")]
[assembly: AssemblyCopyright("杨子江 by DeepSeek HARNESS")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

static class Launcher
{
    const int PORT = 8765;
    const string URL = "http://127.0.0.1:8765/";
    const string TITLE = "\u8868\u683c\u68c0\u7d22 \u00b7 \u6a21\u677f\u6539\u9020 \u00b7 \u6570\u636e\u6838\u5bf9";

    static string appDir;
    static string pythonExe;
    static string logPath;
    static Process server;
    static Form form;
    static Label status;
    static Button openBtn;
    static Button stopBtn;
    static bool startedByUs;

    [STAThread]
    static void Main(string[] args)
    {
        appDir = Path.GetDirectoryName(Application.ExecutablePath);
        pythonExe = Path.Combine(appDir, "runtime", "python.exe");
        Directory.CreateDirectory(Path.Combine(appDir, "data", "logs"));
        logPath = Path.Combine(appDir, "data", "logs", "service.log");

        if (args.Length > 0 && args[0] == "--selftest")
        {
            RunSelfTest();
            return;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        BuildUi();

        if (HealthOk())
        {
            SetStatus("\u670d\u52a1\u5df2\u5728\u8fd0\u884c\uff0c\u53ef\u76f4\u63a5\u6253\u5f00\u754c\u9762\u3002", true);
            OpenBrowser();
        }
        else
        {
            StartServer();
        }

        Application.Run(form);
    }

    // 无界面自检：验证"启动 python 服务 → 健康检查 → 清理"整条链路（供打包脚本/CI 使用）
    static void RunSelfTest()
    {
        string resultPath = Path.Combine(appDir, "data", "logs", "selftest.txt");
        string result;
        try
        {
            if (HealthOk())
            {
                result = "PASS already-running";
            }
            else
            {
                var psi = new ProcessStartInfo();
                psi.FileName = pythonExe;
                psi.Arguments = "-m app.server --port " + PORT + " -n --log \"" + logPath + "\"";
                psi.WorkingDirectory = appDir;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                server = new Process();
                server.StartInfo = psi;
                server.Start();
                startedByUs = true;
                bool ok = false;
                for (int i = 0; i < 80; i++)
                {
                    if (server.HasExited) break;
                    if (HealthOk()) { ok = true; break; }
                    Thread.Sleep(500);
                }
                result = ok ? "PASS " + URL : "FAIL server-not-ready";
            }
        }
        catch (Exception ex)
        {
            result = "FAIL " + ex.GetType().Name + ": " + ex.Message;
        }
        try
        {
            File.WriteAllText(resultPath, result + Environment.NewLine, Encoding.UTF8);
        }
        catch (Exception) { }
        Shutdown();
        Environment.Exit(result.StartsWith("PASS") ? 0 : 1);
    }

    static void BuildUi()
    {
        form = new Form();
        form.Text = TITLE;
        form.ClientSize = new Size(520, 230);
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MaximizeBox = false;
        form.Font = new Font("Microsoft YaHei UI", 9f);

        var title = new Label();
        title.Text = TITLE;
        title.Font = new Font("Microsoft YaHei UI", 13f, FontStyle.Bold);
        title.AutoSize = true;
        title.Location = new Point(22, 18);
        form.Controls.Add(title);

        var sub = new Label();
        sub.Text = "\u672c\u5730\u5fae\u578b Web \u5e94\u7528\uff1a\u5168\u5b57\u6bb5\u68c0\u7d22 / \u5b57\u6bb5\u6620\u5c04\u8f6c\u6362 / \u53cc\u8868\u6838\u5bf9 / \u5355\u5143\u683c\u4fee\u6539";
        sub.ForeColor = Color.DimGray;
        sub.AutoSize = true;
        sub.Location = new Point(24, 52);
        form.Controls.Add(sub);

        status = new Label();
        status.Text = "\u6b63\u5728\u542f\u52a8\u670d\u52a1\u2026";
        status.AutoSize = false;
        status.Size = new Size(470, 44);
        status.Location = new Point(24, 84);
        status.ForeColor = Color.DimGray;
        form.Controls.Add(status);

        openBtn = new Button();
        openBtn.Text = "\u6253\u5f00\u754c\u9762";
        openBtn.Size = new Size(140, 36);
        openBtn.Location = new Point(24, 150);
        openBtn.Enabled = false;
        openBtn.Click += delegate { OpenBrowser(); };
        form.Controls.Add(openBtn);

        var dataBtn = new Button();
        dataBtn.Text = "\u6253\u5f00\u6570\u636e\u76ee\u5f55";
        dataBtn.Size = new Size(140, 36);
        dataBtn.Location = new Point(176, 150);
        dataBtn.Click += delegate
        {
            string d = Path.Combine(appDir, "data");
            if (!Directory.Exists(d)) return;
            try
            {
                // 用 ShellExecute 打开资源管理器，避免 explorer.exe 0xc0000142（DLL 初始化失败）
                var psi = new ProcessStartInfo();
                psi.FileName = "explorer.exe";
                psi.Arguments = "\"" + d + "\"";
                psi.UseShellExecute = true;
                Process.Start(psi);
            }
            catch (Exception) { }
        };
        form.Controls.Add(dataBtn);

        stopBtn = new Button();
        stopBtn.Text = "\u505c\u6b62\u5e76\u9000\u51fa";
        stopBtn.Size = new Size(140, 36);
        stopBtn.Location = new Point(328, 150);
        stopBtn.Click += delegate { form.Close(); };
        form.Controls.Add(stopBtn);

        form.FormClosing += delegate { Shutdown(); };
    }

    static void SetStatus(string text, bool ok)
    {
        if (form.InvokeRequired)
        {
            form.Invoke(new Action<string, bool>(SetStatus), new object[] { text, ok });
            return;
        }
        status.Text = text;
        status.ForeColor = ok ? Color.SeaGreen : Color.DimGray;
        openBtn.Enabled = ok;
    }

    static bool HealthOk()
    {
        try
        {
            var req = (HttpWebRequest)WebRequest.Create(URL + "api/health");
            req.Timeout = 1200;
            req.ReadWriteTimeout = 1200;
            using (var resp = (HttpWebResponse)req.GetResponse())
            {
                return resp.StatusCode == HttpStatusCode.OK;
            }
        }
        catch (Exception) { return false; }
    }

    static void StartServer()
    {
        if (!File.Exists(pythonExe))
        {
            SetStatus("\u672a\u627e\u5230 runtime\\python.exe\uff0c\u8bf7\u91cd\u65b0\u89e3\u538b\u5b8c\u6574\u5b89\u88c5\u5305\u3002", false);
            return;
        }
        try
        {
            var psi = new ProcessStartInfo();
            psi.FileName = pythonExe;
            psi.Arguments = "-m app.server --port " + PORT + " -n --log \"" + logPath + "\"";
            psi.WorkingDirectory = appDir;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            server = new Process();
            server.StartInfo = psi;
            server.Start();
            startedByUs = true;
            SetStatus("\u6b63\u5728\u542f\u52a8\u670d\u52a1\uff08\u9996\u6b21\u542f\u52a8\u7ea6 2\u20135 \u79d2\uff09\u2026", false);
            var th = new Thread(WaitReady);
            th.IsBackground = true;
            th.Start();
        }
        catch (Exception ex)
        {
            SetStatus("\u542f\u52a8\u5931\u8d25\uff1a" + ex.Message + "\n\u65e5\u5fd7\uff1a" + logPath, false);
        }
    }

    static void WaitReady()
    {
        for (int i = 0; i < 80; i++)
        {
            if (server != null && server.HasExited)
            {
                SetStatus("\u670d\u52a1\u8fdb\u7a0b\u5df2\u9000\u51fa\uff0c\u8bf7\u67e5\u770b\u65e5\u5fd7\uff1a" + logPath, false);
                return;
            }
            if (HealthOk())
            {
                SetStatus("\u670d\u52a1\u5df2\u5c31\u7eea\uff1a" + URL + "\uff08\u5173\u95ed\u672c\u7a97\u53e3\u5373\u505c\u6b62\uff09", true);
                OpenBrowser();
                return;
            }
            Thread.Sleep(500);
        }
        SetStatus("\u670d\u52a1\u542f\u52a8\u8d85\u65f6\uff0c\u8bf7\u67e5\u770b\u65e5\u5fd7\uff1a" + logPath, false);
    }

    static void OpenBrowser()
    {
        try
        {
            var psi = new ProcessStartInfo(URL);
            psi.UseShellExecute = true;
            Process.Start(psi);
        }
        catch (Exception)
        {
            SetStatus("\u670d\u52a1\u5df2\u5728\u8fd0\u884c\uff0c\u4f46\u65e0\u6cd5\u81ea\u52a8\u6253\u5f00\u6d4f\u89c8\u5668\u3002\u8bf7\u624b\u52a8\u8bbf\u95ee\uff1a" + URL, true);
        }
    }

    static void Log(string line)
    {
        if (line == null) return;
        try
        {
            File.AppendAllText(logPath, DateTime.Now.ToString("HH:mm:ss") + " " + line + Environment.NewLine, Encoding.UTF8);
        }
        catch (Exception) { }
    }

    static void Shutdown()
    {
        try
        {
            if (startedByUs && server != null && !server.HasExited)
            {
                server.Kill();
                server.WaitForExit(3000);
            }
        }
        catch (Exception) { }
    }
}
