using System;
using System.Runtime.InteropServices;

namespace xClient.Core.Utilities.TaskScheduler.Legacy
{
    // ------ Types used in in the Task Scheduler Interfaces ------
    internal enum TaskTriggerType
    {
        TIME_TRIGGER_ONCE = 0,  // Ignore the Type field.
        TIME_TRIGGER_DAILY = 1,  // Use DAILY
        TIME_TRIGGER_WEEKLY = 2,  // Use WEEKLY
        TIME_TRIGGER_MONTHLYDATE = 3,  // Use MONTHLYDATE
        TIME_TRIGGER_MONTHLYDOW = 4,  // Use MONTHLYDOW
        EVENT_TRIGGER_ON_IDLE = 5,  // Ignore the Type field.
        EVENT_TRIGGER_AT_SYSTEMSTART = 6,  // Ignore the Type field.
        EVENT_TRIGGER_AT_LOGON = 7   // Ignore the Type field.
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Daily
    {
        public ushort DaysInterval;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Weekly
    {
        public ushort WeeksInterval;
        public ushort DaysOfTheWeek;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MonthlyDate
    {
        public uint Days;
        public ushort Months;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MonthlyDOW
    {
        public ushort WhichWeek;
        public ushort DaysOfTheWeek;
        public ushort Months;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct TriggerTypeData
    {
        [FieldOffset(0)]
        public Daily daily;
        [FieldOffset(0)]
        public Weekly weekly;
        [FieldOffset(0)]
        public MonthlyDate monthlyDate;
        [FieldOffset(0)]
        public MonthlyDOW monthlyDOW;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct TaskTrigger
    {
        public ushort TriggerSize;             // Structure size.
        public ushort Reserved1;               // Reserved. Must be zero.
        public ushort BeginYear;               // Trigger beginning date year.
        public ushort BeginMonth;              // Trigger beginning date month.
        public ushort BeginDay;                // Trigger beginning date day.
        public ushort EndYear;                 // Optional trigger ending date year.
        public ushort EndMonth;                // Optional trigger ending date month.
        public ushort EndDay;                  // Optional trigger ending date day.
        public ushort StartHour;               // Run bracket start time hour.
        public ushort StartMinute;             // Run bracket start time minute.
        public uint MinutesDuration;           // Duration of run bracket.
        public uint MinutesInterval;           // Run bracket repetition interval.
        public uint Flags;                     // Trigger flags.
        public TaskTriggerType Type;           // Trigger type.
        public TriggerTypeData Data;           // Trigger data peculiar to this type (union).
        public ushort Reserved2;               // Reserved. Must be zero.
        public ushort RandomMinutesInterval;   // Maximum number of random minutes after start time.
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SystemTime
    {
        public ushort Year;
        public ushort Month;
        public ushort DayOfWeek;
        public ushort Day;
        public ushort Hour;
        public ushort Minute;
        public ushort Second;
        public ushort Milliseconds;
    }

    // ----- Interfaces -----
    [Guid("148BD527-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITaskScheduler
    {
        void SetTargetComputer([In, MarshalAs(UnmanagedType.LPWStr)] string Computer);
        void GetTargetComputer(out System.IntPtr Computer);
        void Enum([Out, MarshalAs(UnmanagedType.Interface)] out IEnumWorkItems EnumWorkItems);
        void Activate([In, MarshalAs(UnmanagedType.LPWStr)] string Name, [In] ref System.Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void Delete([In, MarshalAs(UnmanagedType.LPWStr)] string Name);
        void NewWorkItem([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In] ref System.Guid rclsid, [In] ref System.Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void AddWorkItem([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In, MarshalAs(UnmanagedType.Interface)] ITask WorkItem);
        void IsOfType([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In] ref System.Guid riid);
    }

    [Guid("148BD528-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IEnumWorkItems
    {
        [PreserveSig()]
        int Next([In] uint RequestCount, [Out] out System.IntPtr Names, [Out] out uint Fetched);
        void Skip([In] uint Count);
        void Reset();
        void Clone([Out, MarshalAs(UnmanagedType.Interface)] out IEnumWorkItems EnumWorkItems);
    }

    [Guid("148BD524-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITask
    {
        void CreateTrigger([Out] out ushort NewTriggerIndex, [Out, MarshalAs(UnmanagedType.Interface)] out ITaskTrigger Trigger);
        void DeleteTrigger([In] ushort TriggerIndex);
        void GetTriggerCount([Out] out ushort Count);
        void GetTrigger([In] ushort TriggerIndex, [Out, MarshalAs(UnmanagedType.Interface)] out ITaskTrigger Trigger);
        void GetTriggerString([In] ushort TriggerIndex, out System.IntPtr TriggerString);
        void GetRunTimes([In, MarshalAs(UnmanagedType.Struct)] ref SystemTime Begin, [In, MarshalAs(UnmanagedType.Struct)] ref SystemTime End, ref ushort Count, [Out] out System.IntPtr TaskTimes);
        void GetNextRunTime([In, Out, MarshalAs(UnmanagedType.Struct)] ref SystemTime NextRun);
        void SetIdleWait([In] ushort IdleMinutes, [In] ushort DeadlineMinutes);
        void GetIdleWait([Out] out ushort IdleMinutes, [Out] out ushort DeadlineMinutes);
        void Run();
        void Terminate();
        void EditWorkItem([In] uint hParent, [In] uint dwReserved);
        void GetMostRecentRunTime([In, Out, MarshalAs(UnmanagedType.Struct)] ref SystemTime LastRun);
        void GetStatus([Out, MarshalAs(UnmanagedType.Error)] out int Status);
        void GetExitCode([Out] out uint ExitCode);
        void SetComment([In, MarshalAs(UnmanagedType.LPWStr)] string Comment);
        void GetComment(out System.IntPtr Comment);
        void SetCreator([In, MarshalAs(UnmanagedType.LPWStr)] string Creator);
        void GetCreator(out System.IntPtr Creator);
        void SetWorkItemData([In] ushort DataLen, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.U1)] byte[] Data);
        void GetWorkItemData([Out] out ushort DataLen, [Out] out System.IntPtr Data);
        void SetErrorRetryCount([In] ushort RetryCount);
        void GetErrorRetryCount([Out] out ushort RetryCount);
        void SetErrorRetryInterval([In] ushort RetryInterval);
        void GetErrorRetryInterval([Out] out ushort RetryInterval);
        void SetFlags([In] uint Flags);
        void GetFlags([Out] out uint Flags);
        void SetAccountInformation([In, MarshalAs(UnmanagedType.LPWStr)] string AccountName, [In] IntPtr Password);
        void GetAccountInformation(out System.IntPtr AccountName);
        void SetApplicationName([In, MarshalAs(UnmanagedType.LPWStr)] string ApplicationName);
        void GetApplicationName(out System.IntPtr ApplicationName);
        void SetParameters([In, MarshalAs(UnmanagedType.LPWStr)] string Parameters);
        void GetParameters(out System.IntPtr Parameters);
        void SetWorkingDirectory([In, MarshalAs(UnmanagedType.LPWStr)] string WorkingDirectory);
        void GetWorkingDirectory(out System.IntPtr WorkingDirectory);
        void SetPriority([In] uint Priority);
        void GetPriority([Out] out uint Priority);
        void SetTaskFlags([In] uint Flags);
        void GetTaskFlags([Out] out uint Flags);
        void SetMaxRunTime([In] uint MaxRunTimeMS);
        void GetMaxRunTime([Out] out uint MaxRunTimeMS);
    }

    [Guid("148BD52B-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITaskTrigger
    {
        void SetTrigger([In, Out, MarshalAs(UnmanagedType.Struct)] ref TaskTrigger Trigger);
        void GetTrigger([In, Out, MarshalAs(UnmanagedType.Struct)] ref TaskTrigger Trigger);
        void GetTriggerString(out System.IntPtr TriggerString);
    }
  
    // ------ Classes ------
    [ComImport, Guid("148BD52A-A2AB-11CE-B11F-00AA00530503")]
    internal class CTaskScheduler
    {
    }

    [ComImport, Guid("148BD520-A2AB-11CE-B11F-00AA00530503")]
    internal class CTask
    {
    }
}
