using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class LiXLauncher
{
    [STAThread]
    private static int Main()
    {
        string root = AppDomain.CurrentDomain.BaseDirectory;
        if (!File.Exists(Path.Combine(root, "Start-LiX.cmd")) ||
            !File.Exists(Path.Combine(root, "Start-OfficeAgent.cmd")))
        {
            MessageBox.Show("启动文件不完整。请将 LiX.exe 保留在项目文件夹内，与 Start-LiX.cmd 放在一起。",
                "LiX — 无法启动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return 2;
        }
        try
        {
            ProcessStartInfo info = new ProcessStartInfo
            {
                FileName = Path.Combine(Environment.SystemDirectory, "cmd.exe"),
                Arguments = "/d /c Start-LiX.cmd",
                WorkingDirectory = root,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            if (String.IsNullOrWhiteSpace(info.EnvironmentVariables["UV_CACHE_DIR"]))
            {
                string cache = Environment.GetEnvironmentVariable("UV_CACHE_DIR", EnvironmentVariableTarget.User);
                if (String.IsNullOrWhiteSpace(cache))
                    cache = Environment.GetEnvironmentVariable("UV_CACHE_DIR", EnvironmentVariableTarget.Machine);
                if (!String.IsNullOrWhiteSpace(cache))
                    info.EnvironmentVariables["UV_CACHE_DIR"] = cache;
            }
            using (Process process = Process.Start(info))
            {
                if (process == null)
                    throw new InvalidOperationException();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.StandardInput.Close();
                process.WaitForExit();
                int code = process.ExitCode;
                if (code != 0)
                    MessageBox.Show("LiX 未能正常启动或已异常退出（代码 " + code +
                        "）。请在项目文件夹运行 Start-LiX.cmd 查看错误。",
                        "LiX — 启动失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return code;
            }
        }
        catch
        {
            MessageBox.Show("无法启动 LiX。请确认项目文件完整，并尝试双击 Start-LiX.cmd。",
                "LiX — 启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
