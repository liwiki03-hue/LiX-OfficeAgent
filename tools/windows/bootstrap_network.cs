using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Microsoft.Win32;

internal static class BootstrapNetwork
{
    private static readonly string[] Keys = { "OFFICE_NETWORK_MODE", "OFFICE_NETWORK_PROXY", "OFFICE_NO_PROXY" };

    private static string Proxy(string value)
    {
        if (String.IsNullOrEmpty(value)) return null;
        Uri uri;
        if (value.Length > 2048 || Regex.IsMatch(value, @"\s") || !Uri.TryCreate(value, UriKind.Absolute, out uri)
                || (uri.Scheme != "http" && uri.Scheme != "https") || String.IsNullOrEmpty(uri.Host)
                || uri.Port == 0 || (uri.AbsolutePath != "" && uri.AbsolutePath != "/")
                || uri.Query != "" || uri.Fragment != "") throw new InvalidDataException("network_proxy_invalid");
        return value;
    }

    private static Dictionary<string, string> Read(string root, Func<string, string> safe)
    {
        var selected = new Dictionary<string, string>(StringComparer.Ordinal);
        string path = safe(Path.Combine(root, ".env"));
        if (!File.Exists(path)) return selected;
        byte[] data = new byte[65537];
        int length = 0;
        using (var stream = File.OpenRead(path))
        {
            int received;
            while (length < data.Length && (received = stream.Read(data, length, data.Length - length)) > 0) length += received;
        }
        if (length > 65536) throw new InvalidDataException("network_settings_invalid");
        string content = new UTF8Encoding(false, true).GetString(data, 0, length).TrimStart('\uFEFF');
        foreach (string line in content.Split(new[] { '\r', '\n' }, StringSplitOptions.None))
        {
            var match = Regex.Match(line, @"^\s*(?:export\s+)?(OFFICE_NETWORK_MODE|OFFICE_NETWORK_PROXY|OFFICE_NO_PROXY)\s*=\s*(.*)$");
            if (!match.Success) continue;
            string key = match.Groups[1].Value, value = match.Groups[2].Value.Trim();
            if (selected.ContainsKey(key)) throw new InvalidDataException("network_settings_invalid");
            if (value.StartsWith("\"") || value.StartsWith("'"))
            {
                char quote = value[0];
                int end = value.LastIndexOf(quote);
                if (end == 0 || (value.Substring(end + 1).Trim().Length > 0 && !value.Substring(end + 1).Trim().StartsWith("#")))
                    throw new InvalidDataException("network_settings_invalid");
                value = quote == '\'' ? value.Substring(1, end - 1) : new JavaScriptSerializer().Deserialize<string>(value.Substring(0, end + 1));
            }
            else value = Regex.Replace(value, @"\s+#.*$", "").Trim();
            if (value == null || value.Length > 2048 || value.IndexOfAny(new[] { '\0', '\r', '\n' }) >= 0)
                throw new InvalidDataException("network_settings_invalid");
            selected[key] = value;
        }
        return selected;
    }

    private static string Value(IDictionary<string, string> settings, IDictionary<string, string> environment, string key, string fallback)
    {
        string value;
        if (settings.TryGetValue(key, out value) && !String.IsNullOrEmpty(value)) return value;
        return environment.TryGetValue(key, out value) && !String.IsNullOrEmpty(value) ? value : fallback;
    }

    private static Dictionary<string, string> SystemProxy()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", false))
        {
            if (key == null || !Object.Equals(key.GetValue("ProxyEnable"), 1)) return result;
            string raw = key.GetValue("ProxyServer") as string;
            if (String.IsNullOrEmpty(raw) || raw.Length > 4096) return result;
            foreach (string part in raw.Split(';'))
            {
                string scheme = "all", address = part.Trim();
                int separator = address.IndexOf('=');
                if (separator >= 0) { scheme = address.Substring(0, separator).Trim(); address = address.Substring(separator + 1).Trim(); }
                if (scheme != "http" && scheme != "https" && scheme != "all") continue;
                if (!address.Contains("://")) address = "http://" + address;
                result[scheme] = Proxy(address);
            }
        }
        return result;
    }

    internal static void Configure(IDictionary<string, string> environment, string root, Func<string, string> safe)
    {
        var settings = Read(root, safe);
        string mode = Value(settings, environment, Keys[0], "direct");
        if (mode != "direct" && mode != "auto" && mode != "manual") throw new InvalidDataException("network_mode_invalid");
        string manual = Value(settings, environment, Keys[1], "");
        if (!String.IsNullOrEmpty(manual)) Proxy(manual);
        if (mode == "manual" && String.IsNullOrEmpty(manual)) throw new InvalidDataException("network_proxy_invalid");
        string bypass = Value(settings, environment, Keys[2], Value(settings, environment, "NO_PROXY", ""));
        if (bypass.Length > 4096 || bypass.IndexOfAny(new[] { '\0', '\r', '\n' }) >= 0) throw new InvalidDataException("network_bypass_invalid");
        string http = Value(settings, environment, "HTTP_PROXY", ""), https = Value(settings, environment, "HTTPS_PROXY", ""), all = Value(settings, environment, "ALL_PROXY", "");
        foreach (string name in new[] { "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY" }) environment.Remove(name);
        if (mode == "manual") http = https = Proxy(manual);
        else if (mode == "auto")
        {
            var system = String.IsNullOrEmpty(http) && String.IsNullOrEmpty(https) && String.IsNullOrEmpty(all)
                ? SystemProxy() : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string proxy;
            if (String.IsNullOrEmpty(http)) http = !String.IsNullOrEmpty(all) ? all : system.TryGetValue("http", out proxy) ? proxy : system.TryGetValue("all", out proxy) ? proxy : null;
            if (String.IsNullOrEmpty(https)) https = !String.IsNullOrEmpty(all) ? all : system.TryGetValue("https", out proxy) ? proxy : system.TryGetValue("all", out proxy) ? proxy : null;
            http = Proxy(http); https = Proxy(https);
        }
        if (mode != "direct")
        {
            if (!String.IsNullOrEmpty(http)) environment["HTTP_PROXY"] = http;
            if (!String.IsNullOrEmpty(https)) environment["HTTPS_PROXY"] = https;
        }
        environment[Keys[0]] = mode;
        environment[Keys[1]] = manual;
        environment[Keys[2]] = bypass;
        environment["NO_PROXY"] = (String.IsNullOrEmpty(bypass) ? "" : bypass + ",") + "localhost,.localhost,127.0.0.1,::1,[::1]";
    }
}
