<#
.SYNOPSIS
    Starts the processes of a lab run inside a Windows job object, so that a run can never stop a process that it did not start.

.DESCRIPTION
    A process is created suspended, put into a job of its own with the handle that its creation returned, and then resumed. Everything it starts
    joins the same job. Stopping the job stops its members and nothing else. No process is looked up by its id, its parent id or its name to stop
    it: Windows uses process ids again, and a walk over parent ids once reached an older process and ended the signed-in session of a lab client.

    The process object that is returned is opened while the creation handle is still open. Windows does not use an id again while a handle to its
    process is open, so the object names the process that was started and no other.

    The job is set to end its members when its last handle closes. The script that starts a run closes the job in a finally block; when the
    whole PowerShell process ends in any way, Windows closes the handle.

    The test runners (Invoke-FullRegression.ps1, Invoke-PackagerSmoke.ps1) start their children with it. The job helper is the one of App
    Packager Desktop (tools\lab-tests\ProcessJob.ps1, pull request 315) and of the suite harness. The file is plain ASCII.
#>
if (-not ('SuiteLab.JobProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace SuiteLab
{
    public sealed class JobProcess
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct BasicLimits
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount, ReadTransferCount, WriteTransferCount, OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ExtendedLimits
        {
            public BasicLimits Basic;
            public IoCounters Io;
            public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes
        {
            public int Length;
            public IntPtr Descriptor;
            public int InheritHandle;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct StartupInfo
        {
            public int cb;
            public string lpReserved, lpDesktop, lpTitle;
            public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
            public short wShowWindow, cbReserved2;
            public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ProcessInformation
        {
            public IntPtr hProcess, hThread;
            public int dwProcessId, dwThreadId;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr attributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr job, int infoClass, ref ExtendedLimits info, int length);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryInformationJobObject(IntPtr job, int infoClass, IntPtr info, int length, IntPtr returnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateProcess(IntPtr process, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint ResumeThread(IntPtr thread);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateFile(string name, uint access, uint share, ref SecurityAttributes attributes, uint disposition, uint flags, IntPtr template);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateProcess(string application, StringBuilder commandLine, IntPtr processAttributes, IntPtr threadAttributes,
            bool inheritHandles, uint flags, IntPtr environment, string directory, ref StartupInfo startup, out ProcessInformation information);

        [DllImport("kernel32.dll")]
        private static extern uint GetOEMCP();

        public static int OemCodePage { get { return (int)GetOEMCP(); } }

        private IntPtr _job;

        public Process Process { get; private set; }

        private static IntPtr OutputFile(string path)
        {
            var attributes = new SecurityAttributes();
            attributes.Length = Marshal.SizeOf(typeof(SecurityAttributes));
            attributes.InheritHandle = 1;
            // GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS
            IntPtr handle = CreateFile(path ?? "NUL", 0x40000000, 1, ref attributes, 2, 0x80, IntPtr.Zero);
            if (handle == new IntPtr(-1)) throw new Win32Exception(Marshal.GetLastWin32Error());
            return handle;
        }

        public static JobProcess Start(string commandLine, string standardOutput, string standardError, string directory)
        {
            IntPtr job = CreateJobObject(IntPtr.Zero, null); // no name: the job belongs to this process alone
            if (job == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            IntPtr output = IntPtr.Zero, error = IntPtr.Zero;
            var information = new ProcessInformation();
            bool resumed = false;
            try
            {
                var limits = new ExtendedLimits();
                limits.Basic.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                if (!SetInformationJobObject(job, 9, ref limits, Marshal.SizeOf(typeof(ExtendedLimits)))) throw new Win32Exception(Marshal.GetLastWin32Error());
                output = OutputFile(standardOutput);
                // One file for both streams: both use one handle, so the lines stay in the order of their writes.
                bool shared = standardError != null && string.Equals(standardError, standardOutput, StringComparison.OrdinalIgnoreCase);
                if (!shared) error = OutputFile(standardError);
                var startup = new StartupInfo();
                startup.cb = Marshal.SizeOf(typeof(StartupInfo));
                startup.dwFlags = 0x100; // STARTF_USESTDHANDLES
                startup.hStdOutput = output;
                startup.hStdError = shared ? output : error;
                // CREATE_SUSPENDED | CREATE_NO_WINDOW: the child gets a console of its own, so no console event is shared with the caller.
                if (!CreateProcess(null, new StringBuilder(commandLine), IntPtr.Zero, IntPtr.Zero, true, 0x4 | 0x08000000, IntPtr.Zero, directory, ref startup, out information))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                if (!AssignProcessToJobObject(job, information.hProcess)) throw new Win32Exception(Marshal.GetLastWin32Error());
                var process = System.Diagnostics.Process.GetProcessById(information.dwProcessId);
                IntPtr keep = process.Handle; // opened while the creation handle is open, so it is this process
                if (ResumeThread(information.hThread) == 0xFFFFFFFF) throw new Win32Exception(Marshal.GetLastWin32Error());
                resumed = true;
                var run = new JobProcess();
                run._job = job;
                run.Process = process;
                return run;
            }
            finally
            {
                if (!resumed)
                {
                    // Only the handle of the creation above is used here.
                    if (information.hProcess != IntPtr.Zero) TerminateProcess(information.hProcess, 1);
                    CloseHandle(job);
                }
                if (information.hThread != IntPtr.Zero) CloseHandle(information.hThread);
                if (information.hProcess != IntPtr.Zero) CloseHandle(information.hProcess);
                if (output != IntPtr.Zero) CloseHandle(output);
                if (error != IntPtr.Zero) CloseHandle(error);
            }
        }

        // The ids of the processes that run in the job now. For a report only: the job itself is what stops them.
        public int[] MemberIds()
        {
            if (_job == IntPtr.Zero) return new int[0];
            const int capacity = 1024;
            int length = 8 + capacity * IntPtr.Size;
            IntPtr buffer = Marshal.AllocHGlobal(length);
            try
            {
                // JobObjectBasicProcessIdList: two counts, then one pointer-sized id for each process.
                if (!QueryInformationJobObject(_job, 3, buffer, length, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
                int count = Marshal.ReadInt32(buffer, 4);
                var ids = new int[count];
                for (int i = 0; i < count; i++) ids[i] = (int)Marshal.ReadIntPtr(buffer, 8 + i * IntPtr.Size).ToInt64();
                return ids;
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        public void Stop()
        {
            if (_job == IntPtr.Zero) return;
            TerminateJobObject(_job, 1);
            // The call above does not wait. The caller reads and writes the files of the run next, so the members must be gone.
            var waited = Stopwatch.StartNew();
            while (MemberIds().Length > 0 && waited.ElapsedMilliseconds < 15000) System.Threading.Thread.Sleep(50);
            CloseHandle(_job);
            _job = IntPtr.Zero;
        }
    }
}
'@
}

# Quotes one argument for a Windows command line, by the rules that CommandLineToArgvW reads.
function ConvertTo-NativeArgument {
    param([AllowEmptyString()][string]$Value)
    '"' + (($Value -replace '(\\*)"', '$1$1\"') -replace '(\\+)$', '$1$1') + '"'
}

# Starts a program in a job of its own. The arguments are joined with spaces, so an argument with a space carries its own quotes.
function Start-JobProcess {
    param([Parameter(Mandatory)][string]$FilePath, [string[]]$ArgumentList = @(), [string]$StandardOutput, [string]$StandardError)
    $line = (@("`"$FilePath`"") + $ArgumentList) -join ' '
    $out = if ($StandardOutput) { $StandardOutput } else { [NullString]::Value }
    $err = if ($StandardError) { $StandardError } else { [NullString]::Value }
    # The child starts in the current folder of PowerShell; the folder of the process itself does not follow Set-Location.
    [SuiteLab.JobProcess]::Start($line, $out, $err, (Get-Location).ProviderPath)
}

# Stops every member of the job that still runs, then closes the job. Safe to call again, and for a process that has ended.
function Stop-JobProcess {
    param([Parameter(Mandatory)]$Run)
    $Run.Stop()
}

# The text of a redirected output file of a job member. The member has a console of its own, so it writes in the OEM code page of the
# computer. A member of an ended job may still hold the file for a moment, so the read allows other writers.
function Read-RunOutput {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { (New-Object IO.StreamReader($stream, [Text.Encoding]::GetEncoding([SuiteLab.JobProcess]::OemCodePage), $true)).ReadToEnd() }
    finally { $stream.Dispose() }
}
