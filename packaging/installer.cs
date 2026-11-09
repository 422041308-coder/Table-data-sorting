using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("表格工具 安装程序")]
[assembly: AssemblyProduct("表格检索 / 模板改造 / 数据核对")]
[assembly: AssemblyCompany("杨子江 by DeepSeek HARNESS")]
[assembly: AssemblyCopyright("杨子江 by DeepSeek HARNESS")]
[assembly: AssemblyDescription("安装包：程序 + Python 运行时 + 证书 + 卸载器")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

static class Installer
{
    [STAThread]
    static int Main(string[] args)
    {
        bool silent = false;
        foreach (string a in args)
        {
            if (a == "/S" || a == "/s" || a == "/silent") silent = true;
        }
        string temp = Path.Combine(Path.GetTempPath(), "tabletool_install_" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(temp);
            var asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                string shortName = name;
                int i = name.LastIndexOf('.');
                if (name.EndsWith(".payload", StringComparison.OrdinalIgnoreCase)) shortName = "payload.zip";
                else if (name.EndsWith(".install", StringComparison.OrdinalIgnoreCase)) shortName = "install.cmd";
                else if (name.EndsWith(".uninstall_ps1", StringComparison.OrdinalIgnoreCase)) shortName = "uninstall.ps1";
                else if (name.EndsWith(".uninstall_bat", StringComparison.OrdinalIgnoreCase)) shortName = "uninstall.bat";
                using (var s = asm.GetManifestResourceStream(name))
                using (var f = File.Create(Path.Combine(temp, shortName)))
                {
                    s.CopyTo(f);
                }
            }
            var psi = new ProcessStartInfo();
            psi.FileName = "cmd.exe";
            psi.Arguments = "/c \"" + Path.Combine(temp, "install.cmd") + "\"" + (silent ? " /S" : "");
            psi.WorkingDirectory = temp;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = !silent;
            var p = Process.Start(psi);
            p.WaitForExit();
            try
            {
                string dbg = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "install_debug.log");
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("exit=" + p.ExitCode);
                sb.AppendLine("temp=" + temp);
                sb.AppendLine("args=" + psi.Arguments);
                foreach (string f in Directory.GetFiles(temp))
                {
                    sb.AppendLine("file=" + Path.GetFileName(f) + " " + new FileInfo(f).Length);
                }
                File.WriteAllText(dbg, sb.ToString());
            }
            catch (Exception) { }
            return p.ExitCode;
        }
        catch (Exception ex)
        {
            try
            {
                File.WriteAllText(
                    Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "install_error.log"),
                    ex.ToString());
            }
            catch (Exception) { }
            if (!silent)
            {
                MessageBox.Show("安装失败：" + ex.Message, "表格工具 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return 1;
        }
        finally
        {
            try
            {
                string log = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "install_error.log");
                if (!File.Exists(log)) Directory.Delete(temp, true);
            }
            catch (Exception) { }
        }
    }
}
