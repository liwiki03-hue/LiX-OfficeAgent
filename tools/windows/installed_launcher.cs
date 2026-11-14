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

internal sealed class RendererProgress
{
    public int schema { get; set; }
    public string state { get; set; }
    public string code { get; set; }
    public long received { get; set; }
    public long total { get; set; }
    public int attempt { get; set; }
    public int? system_code { get; set; }
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
    private static string rendererFailure;
    private static string activeRendererReport;

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
            bool bundledRenderer = row.path == "runtime/LibreOffice_26.2.6.2_Win_x86-64.msi"
                && row.bytes == 373293056 && row.sha256 == "788ce7d4b56460357f57552cb8cd848e8f2254f28f6e51fa61e0c65a18096373";
            if (!path.StartsWith(Root, StringComparison.OrdinalIgnoreCase) || row.bytes < 0
                    || (!bundledRenderer && row.bytes > 64 * 1024 * 1024) || (total += row.bytes) > 512 * 1024 * 1024
                    || new FileInfo(path).Length != row.bytes || FileHash(path) != row.sha256)
                throw new IOException("An installed file is missing or changed");
        }
        Safe(Data);
        Directory.CreateDirectory(Data);
        string logs = Safe(Path.Combine(Data, "SetupLogs"));
        Directory.CreateDirectory(logs);
        diagnostic = Path.Combine(logs, "prepare-" + DateTime.UtcNow.ToString("yyyyMMddTHHmmssfff") + "-" + Guid.NewGuid().ToString("N") + ".log");
        using (File.Create(diagnostic)) { }
        Note("build_version=" + config.version + "; manifest_sha256=" + config.manifest_sha256);
    }

    private static SortedDictionary<string, string> EnvironmentForApp(bool network = false)
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
        values["UV_HTTP_CONNECT_TIMEOUT"] = "10";
        values["UV_HTTP_RETRIES"] = "2";
        if (network) BootstrapNetwork.Configure(values, App, Safe);
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
        values.Remove("LIX_RENDERER_REPORT");
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

    private static string RendererMessage(string code)
    {
        switch (code)
        {
            case "download_timeout": return "下载超时，已保留进度；点击重新准备继续下载。";
            case "download_network": return "下载连接中断；检查网络后点击重新准备续传。";
            case "download_tls": return "下载证书验证失败；请检查系统时间与网络。";
            case "download_http": return "官方下载站暂不可用；稍后重试或使用含渲染器安装包。";
            case "download_range": return "下载站未正确响应续传；请使用含渲染器安装包。";
            case "download_incomplete": return "下载未完成；点击重新准备从已有进度继续。";
            case "download_size":
            case "download_checksum":
            case "partial_invalid": return "下载文件校验失败，未解压；请使用含渲染器安装包。";
            case "disk_permission": return "无法写入运行目录；请检查目录权限或磁盘空间。";
            case "extract_failed": return "渲染器解压失败；请检查磁盘空间及系统安装服务。";
            case "installer_service_unavailable": return "Windows Installer 服务不可访问（1601）；渲染器未准备完成，下载缓存已保留。";
            case "extract_timeout": return "渲染器解压超时；未标记准备完成，请查看准备记录。";
            case "runtime_integrity_or_extraction": return "渲染器文件或解压结果异常；请查看准备记录。";
            default: return "渲染器准备失败；请查看准备记录或使用含渲染器安装包。";
        }
    }

    private static void ReadRendererProgress(string path)
    {
        try
        {
            if (!File.Exists(Safe(path)) || new FileInfo(path).Length > 4096) return;
            RendererProgress row = new JavaScriptSerializer().Deserialize<RendererProgress>(File.ReadAllText(path, Encoding.UTF8));
            if (row == null || row.schema != 1) return;
            string text = null;
            if (row.state == "download" && row.total == 373293056 && row.received >= 0 && row.received <= row.total)
                text = "3 / 4  下载渲染器 " + (row.received / 1048576) + " / 356 MiB";
            else if (row.state == "retry") text = "3 / 4  连接中断，保留进度并重试（最多 3 次）";
            else if (row.state == "verify") text = "3 / 4  使用本地组件，正在校验";
            else if (row.state == "extract") text = "3 / 4  使用本地组件，正在提取";
            else if (row.state == "failed") rendererFailure = text = RendererMessage(row.code) +
                (row.system_code.HasValue ? " 系统错误 " + row.system_code.Value + "。" : "");
            if (text != null && window != null && !window.IsDisposed)
                window.BeginInvoke((Action)(() => { if (activeRendererReport == path) status.Text = text; }));
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        catch (ArgumentException) { }
        catch (InvalidOperationException) { }
    }

    private static void ReadPreparationProgress(string path, DateTime started)
    {
        try
        {
            string text = currentStage + " · 已等待 " + (int)(DateTime.UtcNow - started).TotalSeconds + " 秒";
            if (path != null && File.Exists(Safe(path)) && new FileInfo(path).Length <= 4096)
            {
                RendererProgress row = new JavaScriptSerializer().Deserialize<RendererProgress>(File.ReadAllText(path, Encoding.UTF8));
                if (row != null && row.schema == 1 && row.received >= 0 && row.total >= row.received)
                    text = "2 / 4  编辑资源 · " + row.received + (row.total > 0 ? " / " + row.total : "") + " 字节";
            }
            if (window != null && !window.IsDisposed) window.BeginInvoke((Action)(() => status.Text = text));
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        catch (ArgumentException) { }
        catch (InvalidOperationException) { }
    }

    private static void Execute(string[] args, int seconds, bool renderer = false)
    {
        bool frontend = Array.IndexOf(args, "tools.install_frontend_assets") >= 0;
        SortedDictionary<string, string> environment = EnvironmentForApp(renderer || frontend || (args.Length > 0 && args[0] == "sync"));
        string report = null;
        System.Threading.Timer timer = null;
        DateTime started = DateTime.UtcNow;
        if (renderer || frontend)
        {
            report = Safe(Path.Combine(Path.GetDirectoryName(diagnostic), "renderer-" + Guid.NewGuid().ToString("N") + ".json"));
            environment[renderer ? "LIX_RENDERER_REPORT" : "LIX_SETUP_REPORT"] = report;
            activeRendererReport = report;
            Note("renderer_report=" + Path.GetFileName(report));
            timer = new System.Threading.Timer(_ =>
            {
                if (renderer) ReadRendererProgress(report);
                else ReadPreparationProgress(report, started);
            }, null, 0, 1000);
        }
        else if (seconds > 0) timer = new System.Threading.Timer(_ => ReadPreparationProgress(null, started), null, 0, 1000);
        int code;
        try
        {
            Action<string> failure = args.Length > 0 && args[0] == "sync" ? (Action<string>)(value =>
            {
                Note("dependency_failure=" + value);
                switch (value)
                {
                    case "dependency_cache_denied": rendererFailure = "依赖缓存不可访问，原目录已保留；请在有访问权限的本机用户会话重试。"; break;
                    case "dependency_disk_full": rendererFailure = "运行磁盘空间不足，依赖未准备完成。"; break;
                    case "dependency_tls_failed": rendererFailure = "依赖下载证书验证失败；请检查系统时间与网络。"; break;
                    case "dependency_download_timeout": rendererFailure = "依赖下载超时；请检查网络后重新准备。"; break;
                    case "dependency_network_failed": rendererFailure = "依赖下载连接失败；请检查网络或代理。"; break;
                    case "dependency_lock_failed": rendererFailure = "锁定依赖未能解析；请保留准备记录并核对构建版本。"; break;
                    default: rendererFailure = "运行依赖未准备完成；请查看本次准备记录。"; break;
                }
            }) : null;
            code = OwnedCommand.Run(Safe(Path.Combine(Root, "runtime", "uv.exe")), args, App,
                environment, seconds, () => cancelled, empty => Note("owned_job_empty=" + empty), failure);
        }
        catch (TimeoutException)
        {
            Note("stage_failure=deadline_exceeded");
            rendererFailure = currentStage + "超时；本次进程已停止，请核对准备记录后重试。";
            throw;
        }
        catch (System.ComponentModel.Win32Exception error)
        {
            Note("process_start_system_code=" + error.NativeErrorCode);
            rendererFailure = "准备进程未能启动（系统错误 " + error.NativeErrorCode + "）；请保留准备记录。";
            throw;
        }
        finally
        {
            if (timer != null)
            {
                using (ManualResetEvent stopped = new ManualResetEvent(false))
                {
                    timer.Dispose(stopped);
                    stopped.WaitOne(5000);
                }
                if (renderer) ReadRendererProgress(report);
                activeRendererReport = null;
            }
        }
        Note("exit_code=" + code);
        if (rendererFailure != null) Note((renderer ? "renderer_failure=" : "stage_failure=") + rendererFailure);
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
        rendererFailure = null;
        Stage("1 / 4  准备 Python 与锁定依赖");
        Execute(new[] { "sync", "--locked", "--no-dev", "--no-default-groups",
            "--python", config.python, "--link-mode", "copy" }, 900);
        Stage("2 / 4  核对文档编辑资源");
        Execute(Python("tools.install_frontend_assets"), 210);
        Stage("3 / 4  准备文档渲染器");
        Execute(Python("tools.installed_runtime"), 900, true);
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
                    status.Text = cancelled ? "准备已停止，已下载内容保留。" : rendererFailure ?? currentStage + "未完成；请查看准备记录，修复后再试。";
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
        status = new Label { Text = "点击后开始准备。", Location = new Point(26, 152), Size = new Size(480, 50) };
        progress = new ProgressBar { Location = new Point(26, 208), Size = new Size(480, 5), Style = ProgressBarStyle.Blocks };
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
