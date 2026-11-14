using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

internal static class OwnedCommand
{
    [StructLayout(LayoutKind.Sequential)]
    private struct StartInfo
    {
        public int size;
        public IntPtr reserved, desktop, title;
        public int x, y, width, height, charsX, charsY, fill, flags;
        public short show, reservedSize;
        public IntPtr reservedBytes, input, output, error;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ProcessInfo
    {
        public IntPtr process, thread;
        public uint pid, tid;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Limits
    {
        public long processTime, jobTime;
        public uint flags;
        public UIntPtr minimum, maximum;
        public uint processLimit;
        public UIntPtr affinity;
        public uint priority, scheduling;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ExtendedLimits
    {
        public Limits basic;
        public ulong readOps, writeOps, otherOps, readBytes, writeBytes, otherBytes;
        public UIntPtr processMemory, jobMemory, peakProcessMemory, peakJobMemory;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Accounting
    {
        public long user, kernel, periodUser, periodKernel;
        public uint faults, processes, active, terminated;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetInformationJobObject(IntPtr job, int kind, ref ExtendedLimits data, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool QueryInformationJobObject(IntPtr job, int kind, out Accounting data, uint length, IntPtr returned);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool TerminateJobObject(IntPtr job, uint code);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool TerminateProcess(IntPtr process, uint code);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CreateProcess(string application, StringBuilder command, IntPtr processAttributes,
        IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory,
        ref StartInfo startup, out ProcessInfo information);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetExitCodeProcess(IntPtr handle, out uint code);
    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr handle);

    internal static string Quote(string argument)
    {
        StringBuilder value = new StringBuilder("\"");
        int slashes = 0;
        foreach (char character in argument)
        {
            if (character == '\\') { slashes++; continue; }
            if (character == '"') value.Append('\\', slashes * 2 + 1).Append('"');
            else value.Append('\\', slashes).Append(character);
            slashes = 0;
        }
        return value.Append('\\', slashes * 2).Append('"').ToString();
    }

    internal static int Run(string executable, string[] arguments, string directory,
        SortedDictionary<string, string> environment, int seconds, Func<bool> cancelled, Action<bool> cleaned)
    {
        return Run(executable, arguments, directory, environment, seconds, cancelled, cleaned, null);
    }

    private static string FailureCode(string output, int code)
    {
        string text = output.ToLowerInvariant();
        if ((text.Contains("cache") || text.Contains("缓存")) && (text.Contains("os error 5") || text.Contains("access is denied") || text.Contains("拒绝访问"))) return "dependency_cache_denied";
        if (text.Contains("os error 112") || text.Contains("no space left")) return "dependency_disk_full";
        if (text.Contains("certificate") || text.Contains("tls")) return "dependency_tls_failed";
        if (text.Contains("timed out") || text.Contains("timeout")) return "dependency_download_timeout";
        if (text.Contains("connection") || text.Contains("dns") || text.Contains("network")) return "dependency_network_failed";
        if (text.Contains("no solution found") || text.Contains("failed to resolve") || text.Contains("lockfile")) return "dependency_lock_failed";
        return code < 0 ? "dependency_start_failed" : "dependency_prepare_failed";
    }

    internal static int Run(string executable, string[] arguments, string directory,
        SortedDictionary<string, string> environment, int seconds, Func<bool> cancelled, Action<bool> cleaned,
        Action<string> diagnostics)
    {
        IntPtr job = CreateJobObject(IntPtr.Zero, null);
        if (job == IntPtr.Zero) throw new Win32Exception();
        IntPtr environmentBlock = IntPtr.Zero;
        ProcessInfo process = new ProcessInfo();
        bool resumed = false;
        int completedCode = -1;
        AnonymousPipeServerStream output = null, input = null;
        Thread collector = null;
        StringBuilder captured = new StringBuilder();
        try
        {
            ExtendedLimits limits = new ExtendedLimits();
            limits.basic.flags = 0x2000 | 8;
            limits.basic.processLimit = 32;
            if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf(limits))) throw new Win32Exception();
            StringBuilder block = new StringBuilder();
            foreach (KeyValuePair<string, string> row in environment)
                block.Append(row.Key).Append('=').Append(row.Value).Append('\0');
            block.Append('\0');
            environmentBlock = Marshal.StringToHGlobalUni(block.ToString());
            StringBuilder command = new StringBuilder(Quote(executable));
            foreach (string argument in arguments) command.Append(' ').Append(Quote(argument));
            StartInfo startup = new StartInfo();
            startup.size = Marshal.SizeOf(startup);
            if (diagnostics != null)
            {
                output = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);
                input = new AnonymousPipeServerStream(PipeDirection.Out, HandleInheritability.Inheritable);
                startup.flags = 0x100;
                startup.output = startup.error = output.ClientSafePipeHandle.DangerousGetHandle();
                startup.input = input.ClientSafePipeHandle.DangerousGetHandle();
                collector = new Thread(() =>
                {
                    byte[] buffer = new byte[4096];
                    try
                    {
                        int count;
                        while ((count = output.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            captured.Append(Encoding.UTF8.GetString(buffer, 0, count));
                            if (captured.Length > 65536) captured.Remove(0, captured.Length - 65536);
                        }
                    }
                    catch (IOException) { }
                    catch (ObjectDisposedException) { }
                });
                collector.IsBackground = true;
                collector.Start();
            }
            if (!CreateProcess(executable, command, IntPtr.Zero, IntPtr.Zero, diagnostics != null,
                    4 | 0x400 | 0x08000000, environmentBlock, directory, ref startup, out process))
                throw new Win32Exception();
            if (output != null) output.DisposeLocalCopyOfClientHandle();
            if (input != null) { input.DisposeLocalCopyOfClientHandle(); input.Dispose(); input = null; }
            if (!AssignProcessToJobObject(job, process.process)) throw new Win32Exception();
            if (ResumeThread(process.thread) != 1) throw new Win32Exception();
            resumed = true;
            Stopwatch clock = Stopwatch.StartNew();
            while (WaitForSingleObject(process.process, 100) == 258)
            {
                if (cancelled()) throw new OperationCanceledException();
                if (seconds > 0 && clock.Elapsed.TotalSeconds > seconds) throw new TimeoutException();
            }
            uint exitCode;
            if (!GetExitCodeProcess(process.process, out exitCode)) throw new Win32Exception();
            completedCode = (int)exitCode;
            return completedCode;
        }
        finally
        {
            if (process.process != IntPtr.Zero && !resumed) TerminateProcess(process.process, 1);
            TerminateJobObject(job, 1);
            Stopwatch deadline = Stopwatch.StartNew();
            Accounting account;
            bool empty = false;
            do
            {
                if (QueryInformationJobObject(job, 1, out account, (uint)Marshal.SizeOf(typeof(Accounting)), IntPtr.Zero)
                        && account.active == 0) { empty = true; break; }
                Thread.Sleep(20);
            } while (deadline.ElapsedMilliseconds < 5000);
            if (process.thread != IntPtr.Zero) CloseHandle(process.thread);
            if (process.process != IntPtr.Zero) CloseHandle(process.process);
            if (environmentBlock != IntPtr.Zero) Marshal.FreeHGlobal(environmentBlock);
            CloseHandle(job);
            if (input != null) input.Dispose();
            bool readerStopped = collector == null || collector.Join(2000);
            if (output != null) output.Dispose();
            if (!readerStopped) readerStopped = collector.Join(1000);
            empty = empty && readerStopped;
            cleaned(empty);
            if (!empty) throw new InvalidOperationException("Owned process cleanup could not be confirmed");
            if (diagnostics != null && completedCode > 0 && !cancelled()) diagnostics(FailureCode(captured.ToString(), completedCode));
        }
    }
}
