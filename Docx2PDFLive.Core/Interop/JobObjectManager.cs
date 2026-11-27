using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Docx2PDFLive.Core.Interop;

/// <summary>
/// Manages a Windows Job Object configured with JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE.
/// When the parent application process terminates or crashes, the Windows kernel
/// automatically terminates all processes assigned to this job object, guaranteeing
/// that no orphaned background processes (e.g., WINWORD.EXE or soffice.exe) are left running.
/// </summary>
public static class JobObjectManager
{
    private static IntPtr _jobHandle = IntPtr.Zero;
    private static readonly object _lock = new();

    public static bool IsInitialized => _jobHandle != IntPtr.Zero;

    public static void InitializeGlobalJobObject()
    {
        lock (_lock)
        {
            if (_jobHandle != IntPtr.Zero) return;

            try
            {
                _jobHandle = CreateJobObject(IntPtr.Zero, null);
                if (_jobHandle == IntPtr.Zero)
                {
                    Debug.WriteLine($"Failed to create Job Object. Win32 error: {Marshal.GetLastWin32Error()}");
                    return;
                }

                var basicLimits = new JOBOBJECT_BASIC_LIMIT_INFORMATION
                {
                    LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                };

                var extendedLimits = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
                {
                    BasicLimitInformation = basicLimits
                };

                int length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
                IntPtr extendedLimitsPtr = Marshal.AllocHGlobal(length);
                try
                {
                    Marshal.StructureToPtr(extendedLimits, extendedLimitsPtr, false);

                    if (!SetInformationJobObject(
                        _jobHandle,
                        JobObjectInfoClass.JobObjectExtendedLimitInformation,
                        extendedLimitsPtr,
                        (uint)length))
                    {
                        Debug.WriteLine($"Failed to set Job Object info. Win32 error: {Marshal.GetLastWin32Error()}");
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(extendedLimitsPtr);
                }

                // Keep _jobHandle open in this process without assigning Docx2PDF itself to the job.
                // Only explicitly assigned worker processes (WINWORD.EXE, soffice.exe) enter the job,
                // so unrelated detached launches (such as the auto-updater) survive Docx2PDF exit.
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception initializing Job Object: {ex.Message}");
            }
        }
    }

    public static bool AssignProcess(IntPtr processHandle)
    {
        if (_jobHandle == IntPtr.Zero)
        {
            InitializeGlobalJobObject();
        }

        if (_jobHandle != IntPtr.Zero && processHandle != IntPtr.Zero)
        {
            try
            {
                return AssignProcessToJobObject(_jobHandle, processHandle);
            }
            catch
            {
                return false;
            }
        }

        return false;
    }

    public static bool AssignProcess(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                return AssignProcess(process.Handle);
            }
        }
        catch { }
        return false;
    }

    public static bool IsProcessAssigned(Process process)
    {
        try
        {
            if (_jobHandle == IntPtr.Zero || process.HasExited) return false;
            return IsProcessInJob(process.Handle, _jobHandle, out bool inJob) && inJob;
        }
        catch
        {
            return false;
        }
    }

    #region Win32 API Definitions

    private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;

    private enum JobObjectInfoClass
    {
        JobObjectExtendedLimitInformation = 9
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IO_COUNTERS
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
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
    private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryLimit;
        public UIntPtr PeakJobMemoryLimit;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetInformationJobObject(
        IntPtr hJob,
        JobObjectInfoClass JobObjectInformationClass,
        IntPtr lpJobObjectInformation,
        uint cbJobObjectInformationLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool IsProcessInJob(IntPtr hProcess, IntPtr hJob, out bool result);

    #endregion
}
