using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

internal sealed class InstallerConfig
{
    public string manifest_sha256 { get; set; }
    public string version { get; set; }
    public string python { get; set; }
}

internal sealed class InstalledFile
{
    public string path { get; set; }
    public long bytes { get; set; }
    public string sha256 { get; set; }
}

internal sealed class InstallManifest
{
    public int schema { get; set; }
    public string version { get; set; }
    public string channel { get; set; }
    public InstalledFile[] files { get; set; }
}

internal static class InstalledLiX
{
    private static readonly string Root = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
    private static readonly string App = Path.Combine(Root, "app");
    private static readonly string Data = Path.Combine(Environment.GetEnvironmentVariable("LOCALAPPDATA")
        ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LiX");
    private static InstallerConfig config;
    private static string currentStage = "检查安装文件";
    private static string diagnostic;
    private static volatile bool cancelled;
    private static Form window;
    private static Label status;
    private static ProgressBar progress;
    private static Button action;
    private static bool preparing;
    private static bool starting;

    private static string Hash(byte[] data)
    {
        using (SHA256 sha = SHA256.Create())
            return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
    }

    private static string FileHash(string path)
    {
        using (SHA256 sha = SHA256.Create())
        using (FileStream stream = File.OpenRead(path))
            return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
    }

    private static string Safe(string path)
    {
        string full = Path.GetFullPath(path);
        if (full.StartsWith("\\\\", StringComparison.Ordinal)) throw new IOException("Local paths only");
        for (string item = full; !String.IsNullOrEmpty(item); item = Path.GetDirectoryName(item))
        {
            try
            {
                if ((File.GetAttributes(item) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Linked installation path");
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
        }
        return full;
    }

    private static void Validate()
    {
        using (Stream source = Assembly.GetExecutingAssembly().GetManifestResourceStream("LiX.Bootstrap"))
        using (StreamReader reader = new StreamReader(source, Encoding.UTF8))
            config = new JavaScriptSerializer().Deserialize<InstallerConfig>(reader.ReadToEnd());
        string manifestPath = Safe(Path.Combine(Root, "LiX.install.json"));
        if (new FileInfo(manifestPath).Length > 2 * 1024 * 1024 || FileHash(manifestPath) != config.manifest_sha256)
            throw new IOException("Installation manifest changed");
        InstallManifest manifest = new JavaScriptSerializer { MaxJsonLength = 2 * 1024 * 1024 }
            .Deserialize<InstallManifest>(File.ReadAllText(manifestPath, Encoding.UTF8));
        if (manifest.schema != 1 || manifest.version != config.version || manifest.channel != "network-preview"
                || manifest.files == null || manifest.files.Length > 5000 || manifest.files.Length == 0)
            throw new IOException("Invalid installation manifest");
        HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long total = 0;
        foreach (InstalledFile row in manifest.files)
        {
            if (String.IsNullOrEmpty(row.path) || row.path.Contains("\\") || row.path.Contains(":")
                    || row.path.StartsWith("/") || !seen.Add(row.path)) throw new IOException("Invalid installation file");
            foreach (string part in row.path.Split('/'))
                if (part == ".." || part == "." || part.Length == 0 || part.EndsWith(".") || part.EndsWith(" "))
                    throw new IOException("Invalid installation component");
            string path = Safe(Path.Combine(Root, row.path.Replace('/', Path.DirectorySeparatorChar)));
            if (!path.StartsWith(Root, StringComparison.OrdinalIgnoreCase) || row.bytes < 0
                    || row.bytes > 64 * 1024 * 1024 || (total += row.bytes) > 512 * 1024 * 1024
                    || new FileInfo(path).Length != row.bytes || FileHash(path) != row.sha256)
                throw new IOException("An installed file is missing or changed");
        }
        Safe(Data);
        Directory.CreateDirectory(Data);
        string logs = Safe(Path.Combine(Data, "SetupLogs"));
        Directory.CreateDirectory(logs);
        diagnostic = Path.Combine(logs, "prepare-" + DateTime.UtcNow.ToString("yyyyMMddTHHmmssfff") + "-" + Guid.NewGuid().ToString("N") + ".log");
        using (File.Create(diagnostic)) { }
    }

    private static SortedDictionary<string, string> EnvironmentForApp()
    {
        SortedDictionary<string, string> values = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry row in Environment.GetEnvironmentVariables())
        {
            string key = (string)row.Key;
            if (key.StartsWith("UV_", StringComparison.OrdinalIgnoreCase) || key.StartsWith("PYTHON", StringComparison.OrdinalIgnoreCase)
                    || key.Equals("VIRTUAL_ENV", StringComparison.OrdinalIgnoreCase)) continue;
            values[key] = (string)row.Value;
        }
        values["UV_CACHE_DIR"] = Safe(Path.Combine(Data, "Cache"));
        values["UV_PYTHON_INSTALL_DIR"] = Safe(Path.Combine(Data, "Runtime", "python"));
        values["UV_PROJECT_ENVIRONMENT"] = Safe(Path.Combine(Data, "Environments", config.manifest_sha256.Substring(0, 16)));
        values["UV_PYTHON_PREFERENCE"] = "only-managed";
        values["UV_PYTHON_DOWNLOADS"] = "automatic";
        values["UV_NO_ENV_FILE"] = "1";
        values["UV_NO_CONFIG"] = "1";
        values["UV_NO_PROGRESS"] = "1";
        values["UV_HTTP_TIMEOUT"] = "60";
        string temporary = Safe(Path.Combine(Data, "Temp"));
        Directory.CreateDirectory(temporary);
        values["TEMP"] = Safe(temporary);
        values["TMP"] = temporary;
        values["PYTHONUTF8"] = "1";
        values["PYTHONNOUSERSITE"] = "1";
        values["PYTHONDONTWRITEBYTECODE"] = "1";
        values["PYTHONNET_RUNTIME"] = "netfx";
        values["OFFICE_WEBVIEW_GUI"] = "edgechromium";
        values.Remove("OFFICE_LIBREOFFICE_PATH");
        return values;
    }

    private static void Note(string text)
    {
        if (diagnostic != null) File.AppendAllText(diagnostic, DateTime.UtcNow.ToString("o") + " " + text + "\r\n", Encoding.UTF8);
    }

    private static void Stage(string text)
    {
        currentStage = text;
        Note("stage=" + text);
        if (window != null && !window.IsDisposed) window.BeginInvoke((Action)(() => status.Text = text));
    }

    private static void Execute(string[] args, int seconds)
    {
        int code = OwnedCommand.Run(Safe(Path.Combine(Root, "runtime", "uv.exe")), args, App,
            EnvironmentForApp(), seconds, () => cancelled, empty => Note("owned_job_empty=" + empty));
        Note("exit_code=" + code);
        if (code != 0) throw new InvalidOperationException("Stage exit " + code);
    }

    private static string[] Python(string module)
    {
        return new[] { "run", "--frozen", "--no-sync", "--no-dev", "--python", config.python, "python", "-m", module };
    }

    private static string ReadyPath()
    {
        return Safe(Path.Combine(App, ".runtime", "installed-ready.txt"));
    }

    private static bool Ready()
    {
        string ready = ReadyPath();
        return File.Exists(ready) && new FileInfo(ready).Length == 64
            && File.ReadAllText(ready, Encoding.ASCII) == config.manifest_sha256
            && File.Exists(Safe(Path.Combine(App, ".runtime", "libreoffice", "program", "soffice.com")))
            && File.Exists(Safe(Path.Combine(App, "office_agent", "webui", "static", "vendor", "superdoc", "superdoc.min.js")))
            && File.Exists(Safe(Path.Combine(EnvironmentForApp()["UV_PROJECT_ENVIRONMENT"], "Scripts", "python.exe")));
    }

    private static void Prepare()
    {
        Stage("1 / 4  准备 Python 与锁定依赖");
        Execute(new[] { "sync", "--locked", "--no-dev", "--no-default-groups",
            "--python", config.python, "--link-mode", "copy" }, 900);
        Stage("2 / 4  核对文档编辑资源");
        Execute(Python("tools.install_frontend_assets"), 150);
        Stage("3 / 4  准备文档渲染器");
        Execute(Python("tools.installed_runtime"), 900);
        Stage("4 / 4  检查桌面运行环境");
        Execute(Python("tools.installed_preflight"), 60);
        string ready = ReadyPath();
        Directory.CreateDirectory(Path.GetDirectoryName(ready));
        if (File.Exists(ready)) File.Move(ready, ready + ".previous-" + Guid.NewGuid().ToString("N"));
        using (FileStream stream = new FileStream(ready, FileMode.CreateNew, FileAccess.Write))
        {
            byte[] bytes = Encoding.ASCII.GetBytes(config.manifest_sha256);
            stream.Write(bytes, 0, bytes.Length);
            stream.Flush(true);
        }
        Note("preparation_complete=true; model_requests=0");
    }

    private static void RunApp()
    {
        Stage("正在打开 LiX");
        if (window != null) window.BeginInvoke((Action)(() => { starting = true; window.Hide(); }));
        Execute(new[] { "run", "--frozen", "--no-sync", "--no-dev", "--python", config.python,
            "python", "-m", "office_agent.app", "--ui", "web" }, 0);
    }

    private static void Begin()
    {
        if (preparing) return;
        cancelled = false;
        preparing = true;
        action.Enabled = false;
        progress.Style = ProgressBarStyle.Marquee;
        Task.Factory.StartNew(() =>
        {
            try
            {
                Prepare();
                RunApp();
                window.BeginInvoke((Action)(() => { starting = true; window.Close(); }));
            }
            catch (Exception error)
            {
                Note("failed_stage=" + currentStage + "; type=" + error.GetType().Name);
                if (!window.IsDisposed) window.BeginInvoke((Action)(() =>
                {
                    starting = false;
                    preparing = false;
                    status.Text = cancelled ? "准备已停止，已下载内容保留。" : currentStage + "未完成；请查看准备记录，修复后再试。";
                    action.Text = "重新准备";
                    action.Enabled = true;
                    progress.Style = ProgressBarStyle.Blocks;
                    window.Show();
                }));
            }
        });
    }

    private static void ShowPreparation()
    {
        window = new Form { Text = "准备 LiX", ClientSize = new Size(540, 300), StartPosition = FormStartPosition.CenterScreen,
            MinimumSize = new Size(480, 320), BackColor = Color.FromArgb(23, 23, 23), ForeColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 10), MaximizeBox = false };
        window.Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
        Label heading = new Label { Text = "首次准备", Location = new Point(24, 24), Size = new Size(470, 32), Font = new Font(window.Font.FontFamily, 18) };
        Label detail = new Label { Text = "联网准备 Python、运行依赖和文档资源。\n无需安装开发工具，不会发送模型请求。\n请预留约 3 GB 空间；下载可能需要数分钟。", Location = new Point(26, 72), Size = new Size(480, 80) };
        status = new Label { Text = "点击后开始准备。", Location = new Point(26, 158), Size = new Size(480, 25) };
        progress = new ProgressBar { Location = new Point(26, 192), Size = new Size(480, 5), Style = ProgressBarStyle.Blocks };
        Button logs = new Button { Text = "准备记录", Location = new Point(26, 230), Size = new Size(110, 36), FlatStyle = FlatStyle.Flat };
        logs.Click += (sender, args) => Process.Start(new ProcessStartInfo { FileName = Path.GetDirectoryName(diagnostic), UseShellExecute = true });
        action = new Button { Text = "准备并启动", Location = new Point(360, 230), Size = new Size(146, 36), FlatStyle = FlatStyle.Flat };
        action.Click += (sender, args) => Begin();
        window.Controls.AddRange(new Control[] { heading, detail, status, progress, logs, action });
        window.MinimumSize = window.Size;
        detail.Anchor = status.Anchor = progress.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        logs.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        action.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        window.FormClosing += (sender, args) =>
        {
            if (preparing && !starting)
            {
                args.Cancel = true;
                if (MessageBox.Show(window, "停止本次准备？已下载内容会保留。", "LiX", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    cancelled = true;
            }
        };
        Application.Run(window);
    }

    private static void Wake()
    {
        try
        {
            string sid = WindowsIdentity.GetCurrent().User.Value;
            string name = "Local\\OfficeAgentOrb-" + Hash(Encoding.UTF8.GetBytes(sid)).Substring(0, 24) + "-wake";
            using (EventWaitHandle signal = EventWaitHandle.OpenExisting(name)) signal.Set();
        }
        catch (WaitHandleCannotBeOpenedException) { }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        bool owned;
        using (Mutex mutex = new Mutex(true, "LiXInstalledRuntime", out owned))
        {
            if (!owned) { if (args.Length == 0) { Wake(); return 0; } return 32; }
            try
            {
                if (args.Length > 1 || (args.Length == 1 && args[0] != "--check" && args[0] != "--prepare-only")) return 2;
                Validate();
                if (args.Length == 1 && args[0] == "--check") return Ready() ? 0 : 42;
                if (args.Length == 1 && args[0] == "--prepare-only") { Prepare(); return 0; }
                if (Ready()) RunApp();
                else ShowPreparation();
                return 0;
            }
            catch (Exception error)
            {
                Note("failed_stage=" + currentStage + "; type=" + error.GetType().Name);
                if (args.Length == 0) MessageBox.Show("LiX 未能完成“" + currentStage + "”。\n"
                    + "安装文件可能缺失，或运行环境尚未准备。请重新安装或查看本机 LiX\\SetupLogs。\n不会自动重试或调用模型。",
                    "LiX — 无法启动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }
            finally { mutex.ReleaseMutex(); }
        }
    }
}
