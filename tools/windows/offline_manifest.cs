using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Win32.SafeHandles;

internal sealed class OfflineComponent
{
    public string id { get; set; }
    public string path { get; set; }
    public long bytes { get; set; }
    public string sha256 { get; set; }
    public int file_count { get; set; }
    public long total_bytes { get; set; }
}

internal sealed class OfflineFiles
{
    public int schema { get; set; }
    public string id { get; set; }
    public InstalledFile[] files { get; set; }
}

internal sealed class OfflineReady
{
    public int schema { get; set; }
    public string manifest_sha256 { get; set; }
    public string lock_sha256 { get; set; }
    public string installation { get; set; }
    public string environment { get; set; }
    public string python_sha256 { get; set; }
    public string environment_python_sha256 { get; set; }
    public bool preflight_passed { get; set; }
    public string inventory_sha256 { get; set; }
}

internal static class OfflineManifest
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(string name, uint access, uint sharing, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFileAttributesW(string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileSizeEx(SafeFileHandle file, out long size);

    private static string Extended(string path)
    {
        if (path.StartsWith("\\\\") || !Path.IsPathRooted(path)) throw new IOException("offline_local_path");
        return "\\\\?\\" + path;
    }

    private static FileStream Open(string path)
    {
        var handle = CreateFileW(Extended(path), 0x80000000, 1, IntPtr.Zero, 3, 0x00200000, IntPtr.Zero);
        if (handle.IsInvalid) { handle.Dispose(); throw new IOException("offline_file_open", new Win32Exception(Marshal.GetLastWin32Error())); }
        return new FileStream(handle, FileAccess.Read);
    }

    internal static long Length(string path)
    {
        using (var file = Open(path))
        {
            long size;
            if (!GetFileSizeEx(file.SafeFileHandle, out size)) throw new IOException("offline_file_size");
            return size;
        }
    }

    internal static string Digest(string path)
    {
        using (var sha = SHA256.Create())
        using (var file = Open(path))
            return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "").ToLowerInvariant();
    }

    internal static string TextHash(string value)
    {
        using (var sha = SHA256.Create())
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant();
    }

    private static string PathFor(string root, string name, Func<string, string> safe)
    {
        if (String.IsNullOrEmpty(name) || name.Length > 240 || name.Contains("\\") || name.StartsWith("/")) throw new IOException("offline_path");
        foreach (string part in name.Split('/'))
        {
            if (part.Length == 0 || part == "." || part == ".." || part.EndsWith(".") || part.EndsWith(" ")
                    || part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) throw new IOException("offline_path");
            string stem = part.Split('.')[0].ToUpperInvariant();
            if (stem == "CON" || stem == "PRN" || stem == "AUX" || stem == "NUL"
                    || System.Text.RegularExpressions.Regex.IsMatch(stem, "^(COM|LPT)[1-9]$")) throw new IOException("offline_path");
        }
        string result = Path.Combine(safe(root), name.Replace('/', Path.DirectorySeparatorChar));
        if (!result.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new IOException("offline_path");
        for (string item = result; !String.IsNullOrEmpty(item); )
        {
            uint attributes = GetFileAttributesW(Extended(item));
            if (attributes == 0xFFFFFFFF || (attributes & 0x400) != 0) throw new IOException("offline_link_or_missing");
            if (item.Length <= 3) break;
            int split = item.LastIndexOf(Path.DirectorySeparatorChar);
            item = split <= 2 ? item.Substring(0, 3) : item.Substring(0, split);
        }
        return result;
    }

    internal static void Validate(string root, InstallManifest manifest, Func<string, string> safe)
    {
        if (manifest.schema != 2 || manifest.channel != "offline-preview" || manifest.deployment_mode != "offline"
                || manifest.platform != "windows-x64" || manifest.python != "3.12.14" || manifest.environment_mode != "assemble-local"
                || manifest.requires_first_run_network || !manifest.bundles_python_runtime || !manifest.bundles_wheelhouse
                || manifest.components == null || manifest.components.Length != 6
                || !System.Text.RegularExpressions.Regex.IsMatch(manifest.build_id ?? "", "^[a-z0-9][a-z0-9-]{0,63}$")) throw new IOException("offline_manifest");
        var budgets = new Dictionary<string, long[]> {
            { "application", new long[] { 5000, 256, 64 } }, { "python", new long[] { 15000, 768, 128 } },
            { "wheels", new long[] { 512, 2048, 512 } }, { "renderer", new long[] { 20000, 2048, 256 } },
            { "prerequisites", new long[] { 16, 768, 384 } }, { "licenses", new long[] { 2000, 128, 32 } }
        };
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long grand = 0;
        foreach (var component in manifest.components)
        {
            if (component == null || !budgets.ContainsKey(component.id) || !ids.Add(component.id)
                    || component.path != "manifests/" + component.id + ".json" || component.bytes <= 0 || component.bytes > 4 * 1048576) throw new IOException("offline_component");
            string path = PathFor(root, component.path, safe);
            if (Length(path) != component.bytes || Digest(path) != component.sha256) throw new IOException("offline_component_identity");
            var data = new JavaScriptSerializer { MaxJsonLength = 4 * 1048576 }.Deserialize<OfflineFiles>(File.ReadAllText(path, Encoding.UTF8));
            long[] budget = budgets[component.id];
            if (data == null || data.schema != 1 || data.id != component.id || data.files == null
                    || data.files.Length == 0 || data.files.Length > budget[0] || data.files.Length != component.file_count) throw new IOException("offline_component_count");
            long total = 0;
            foreach (var row in data.files)
            {
                if (row == null || !names.Add(row.path) || row.bytes < 0 || row.bytes > budget[2] * 1048576L
                        || !System.Text.RegularExpressions.Regex.IsMatch(row.sha256 ?? "", "^[a-f0-9]{64}$")) throw new IOException("offline_file_budget");
                string file = PathFor(root, row.path, safe);
                total = checked(total + row.bytes);
                if (total > budget[1] * 1048576L || Length(file) != row.bytes || Digest(file) != row.sha256) throw new IOException("offline_file_changed");
            }
            if (total != component.total_bytes) throw new IOException("offline_component_size");
            grand = checked(grand + total);
        }
        if (grand != manifest.total_bytes || grand > 6L * 1024 * 1048576
                || Digest(PathFor(root, "app/uv.lock", safe)) != manifest.lock_sha256) throw new IOException("offline_total_or_lock");
    }

    internal static bool Ready(string path, string root, string environment, string manifest, string locked, Func<string, string> safe)
    {
        try
        {
            path = safe(path);
            if (!File.Exists(path) || new FileInfo(path).Length > 4096) return false;
            var row = new JavaScriptSerializer().Deserialize<OfflineReady>(File.ReadAllText(path, Encoding.UTF8));
            bool matches = row != null && row.schema == 2 && row.manifest_sha256 == manifest && row.lock_sha256 == locked
                && row.preflight_passed && row.installation == root && row.environment == environment
                && Digest(safe(Path.Combine(root, "runtime", "python", "python.exe"))) == row.python_sha256
                && Digest(safe(Path.Combine(environment, "Scripts", "python.exe"))) == row.environment_python_sha256;
            if (!matches) return false;
            string inventory = safe(Path.Combine(environment, ".lix-inventory.json"));
            if (new FileInfo(inventory).Length > 4 * 1048576 || Digest(inventory) != row.inventory_sha256) return false;
            var data = new JavaScriptSerializer { MaxJsonLength = 4 * 1048576 }.Deserialize<OfflineFiles>(File.ReadAllText(inventory, Encoding.UTF8));
            if (data == null || data.schema != 1 || data.id != "environment" || data.files == null || data.files.Length == 0 || data.files.Length > 25000) return false;
            long size = 0;
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in data.files)
            {
                if (item == null || !names.Add(item.path) || item.bytes < 0 || item.bytes > 512L * 1048576) return false;
                size = checked(size + item.bytes);
                string file = PathFor(environment, item.path, safe);
                if (size > 2L * 1024 * 1048576 || Length(file) != item.bytes || Digest(file) != item.sha256) return false;
            }
            return true;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    internal static void Publish(string path, string root, string environment, string manifest, string locked, Func<string, string> safe)
    {
        var row = new OfflineReady { schema = 2, manifest_sha256 = manifest, lock_sha256 = locked,
            installation = root, environment = environment, preflight_passed = true,
            python_sha256 = Digest(safe(Path.Combine(root, "runtime", "python", "python.exe"))),
            environment_python_sha256 = Digest(safe(Path.Combine(environment, "Scripts", "python.exe"))),
            inventory_sha256 = Digest(safe(Path.Combine(environment, ".lix-inventory.json"))) };
        string temporary = safe(path + "." + Guid.NewGuid().ToString("N") + ".tmp");
        using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write))
        {
            byte[] data = Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(row));
            stream.Write(data, 0, data.Length);
            stream.Flush(true);
        }
        if (File.Exists(path)) throw new IOException("offline_ready_conflict");
        File.Move(temporary, path);
    }
}
