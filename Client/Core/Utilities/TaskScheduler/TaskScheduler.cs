using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace xClient.Core.Utilities.TaskScheduler
{
    public enum TaskInstancesPolicy
    {
        /// <summary>Starts new instance while an existing instance is running.</summary>
        Parallel,
        /// <summary>Starts a new instance of the task after all other instances of the task are complete.</summary>
        Queue,
        /// <summary>Does not start a new instance if an existing instance of the task is running.</summary>
        IgnoreNew,
        /// <summary>Stops an existing instance of the task before it starts a new instance.</summary>
        StopExisting
    }

    public enum TaskCompatibility
    {
        /// <summary>The task is compatible with the AT command.</summary>
        AT,
        /// <summary>The task is compatible with Task Scheduler 1.0 (Windows Server™ 2003, Windows® XP, or Windows® 2000).</summary>
        V1,
        /// <summary>The task is compatible with Task Scheduler 2.0 (Windows Vista™, Windows Server™ 2008).</summary>
        V2,
        /// <summary>The task is compatible with Task Scheduler 2.1 (Windows® 7, Windows Server™ 2008 R2).</summary>
        V2_1,
        /// <summary>The task is compatible with Task Scheduler 2.2 (Windows® 8.x, Windows Server™ 2012).</summary>
        V2_2,
        // <summary>The task is compatible with Task Scheduler 2.3 (Windows® 10, Windows Server™ 2016).</summary>
        // V2_3
    }

    public enum TaskActionType
    {
        /// <summary>This action fires a handler.</summary>
        ComHandler = 5,

        /// <summary>
        /// This action performs a command-line operation. For example, the action can run a script,
        /// launch an executable, or, if the name of a document is provided, find its associated
        /// application and launch the application with the document.
        /// </summary>
        Execute = 0,

        /// <summary>This action sends and e-mail.</summary>
        SendEmail = 6,

        /// <summary>This action shows a message box.</summary>
        ShowMessage = 7
    }

    public enum TaskLogonType
    {
        /// <summary>The logon method is not specified. Used for non-NT credentials.</summary>
        None,
        /// <summary>Use a password for logging on the user. The password must be supplied at registration time.</summary>
        Password,
        /// <summary>Use an existing interactive token to run a task. The user must log on using a service for user (S4U) logon. When an S4U logon is used, no password is stored by the system and there is no access to either the network or to encrypted files.</summary>
        S4U,
        /// <summary>User must already be logged on. The task will be run only in an existing interactive session.</summary>
        InteractiveToken,
        /// <summary>Group activation. The groupId field specifies the group.</summary>
        Group,
        /// <summary>Indicates that a Local System, Local Service, or Network Service account is being used as a security context to run the task.</summary>
        ServiceAccount,
        /// <summary>First use the interactive token. If the user is not logged on (no interactive token is available), then the password is used. The password must be specified when a task is registered. This flag is not recommended for new tasks because it is less reliable than Password.</summary>
        InteractiveTokenOrPassword
    }

    public enum TaskTriggerType
    {
        /// <summary>Triggers the task when a specific event occurs. Version 1.2 only.</summary>
        Event = 0,
        /// <summary>Triggers the task at a specific time of day.</summary>
        Time = 1,
        /// <summary>Triggers the task on a daily schedule.</summary>
        Daily = 2,
        /// <summary>Triggers the task on a weekly schedule.</summary>
        Weekly = 3,
        /// <summary>Triggers the task on a monthly schedule.</summary>
        Monthly = 4,
        /// <summary>Triggers the task on a monthly day-of-week schedule.</summary>
        MonthlyDOW = 5,
        /// <summary>Triggers the task when the computer goes into an idle state.</summary>
        Idle = 6,
        /// <summary>Triggers the task when the task is registered. Version 1.2 only.</summary>
        Registration = 7,
        /// <summary>Triggers the task when the computer boots.</summary>
        Boot = 8,
        /// <summary>Triggers the task when a specific user logs on.</summary>
        Logon = 9,
        /// <summary>Triggers the task when a specific user session state changes. Version 1.2 only.</summary>
        SessionStateChange = 11,
        /// <summary>Triggers the custom trigger. Version 1.3 only.</summary>
        Custom = 12
    }

    public enum TaskState
    {
        /// <summary>The state of the task is unknown.</summary>
        Unknown,
        /// <summary>The task is registered but is disabled and no instances of the task are queued or running. The task cannot be run until it is enabled.</summary>
        Disabled,
        /// <summary>Instances of the task are queued.</summary>
        Queued,
        /// <summary>The task is ready to be executed, but no instances are queued or running.</summary>
        Ready,
        /// <summary>One or more instances of the task is running.</summary>
        Running
    }

    public enum TaskRunLevel
    {
        /// <summary>Tasks will be run with the least privileges.</summary>
        LUA,
        /// <summary>Tasks will be run with the highest privileges.</summary>
        Highest
    }

    [Flags]
    public enum TaskRunFlags
    {
        /// <summary>The task is run with all flags ignored.</summary>
        NoFlags = 0,
        /// <summary>The task is run as the user who is calling the Run method.</summary>
        AsSelf = 1,
        /// <summary>The task is run regardless of constraints such as "do not run on batteries" or "run only if idle".</summary>
        IgnoreConstraints = 2,
        /// <summary>The task is run using a terminal server session identifier.</summary>
        UseSessionId = 4,
        /// <summary>The task is run using a security identifier.</summary>
        UserSID = 8
    }

    internal enum TaskEnumFlags
    {
        Hidden = 1
    }



    [ComImport, Guid("BAE54997-48B1-4CBE-9965-D6BE263EBEA4"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IAction
    {
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        TaskActionType Type { get; }
    }

    [ComImport, Guid("02820E19-7B98-4ED2-B2E8-FDCCCEFF619B"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IActionCollection : IEnumerable
    {
        int Count { get; }
        IAction this[int index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
        string XmlText { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        [return: MarshalAs(UnmanagedType.Interface)]
        IAction Create([In] TaskActionType Type);
        void Remove([In, MarshalAs(UnmanagedType.Struct)] object index);
        void Clear();
        string Context { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, Guid("2A9C35DA-D357-41F4-BBC1-207AC1B1F3CB"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IBootTrigger : ITrigger
    {
        new TaskTriggerType Type { get; }
        new string Id { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        new IRepetitionPattern Repetition { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        new string ExecutionTimeLimit { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        new string StartBoundary { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        new string EndBoundary { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        new bool Enabled { get; [param: In] set; }

        string Delay { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, Guid("4C3D624D-FD6B-49A3-B9B7-09CB3CD3F047"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IExecAction : IAction
    {
        new string Id { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        new TaskActionType Type { get; }

        string Path { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Arguments { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string WorkingDirectory { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, Guid("84594461-0053-4342-A8FD-088FABF11F32"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IIdleSettings
    {
        string IdleDuration { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string WaitTimeout { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        bool StopOnIdleEnd { get; [param: In] set; }
        bool RestartOnIdle { get; [param: In] set; }
    }

    [ComImport, Guid("D98D51E5-C9B4-496A-A9C1-18980261CF0F"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IPrincipal
    {
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string DisplayName { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string UserId { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        TaskLogonType LogonType { get; set; }
        string GroupId { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        TaskRunLevel RunLevel { get; set; }
    }

    [ComImport, Guid("9C86F320-DEE3-4DD1-B972-A303F26B061E"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity, DefaultMember("Path")]
    internal interface IRegisteredTask
    {
        string Name { [return: MarshalAs(UnmanagedType.BStr)] get; }
        string Path { [return: MarshalAs(UnmanagedType.BStr)] get; }
        TaskState State { get; }
        bool Enabled { get; set; }
        [return: MarshalAs(UnmanagedType.Interface)]
        IRunningTask Run([In, MarshalAs(UnmanagedType.Struct)] object parameters);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRunningTask RunEx([In, MarshalAs(UnmanagedType.Struct)] object parameters, [In] int flags, [In] int sessionID, [In, MarshalAs(UnmanagedType.BStr)] string user);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRunningTaskCollection GetInstances(int flags);
        DateTime LastRunTime { get; }
        int LastTaskResult { get; }
        int NumberOfMissedRuns { get; }
        DateTime NextRunTime { get; }
        ITaskDefinition Definition { [return: MarshalAs(UnmanagedType.Interface)] get; }
        string Xml { [return: MarshalAs(UnmanagedType.BStr)] get; }
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetSecurityDescriptor(int securityInformation);
        void SetSecurityDescriptor([In, MarshalAs(UnmanagedType.BStr)] string sddl, [In] int flags);
        void Stop(int flags);
    }

    [ComImport, Guid("86627EB4-42A7-41E4-A4D9-AC33A72F2D52"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IRegisteredTaskCollection : IEnumerable
    {
        int Count { get; }
        IRegisteredTask this[object index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
    }

    [ComImport, Guid("416D8B73-CB41-4EA1-805C-9BE9A5AC4A74"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IRegistrationInfo
    {
        string Description { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Author { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Version { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Date { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Documentation { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string XmlText { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string URI { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        object SecurityDescriptor { [return: MarshalAs(UnmanagedType.Struct)] get; [param: In, MarshalAs(UnmanagedType.Struct)] set; }
        string Source { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, Guid("7FB9ACF1-26BE-400E-85B5-294B9C75DFD6"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IRepetitionPattern
    {
        string Interval { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Duration { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        bool StopAtDurationEnd { get; [param: In] set; }
    }

    [ComImport, Guid("F5BC8FC5-536D-4F77-B852-FBC1356FDEB6"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITaskDefinition
    {
        IRegistrationInfo RegistrationInfo { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        ITriggerCollection Triggers { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        ITaskSettings Settings { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        string Data { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        IPrincipal Principal { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        IActionCollection Actions { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        string XmlText { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, Guid("8CFAC062-A080-4C15-9A88-AA7C2AF80DFC"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity, DefaultMember("Path")]
    internal interface ITaskFolder
    {
        string Name { [return: MarshalAs(UnmanagedType.BStr)] get; }
        string Path { [return: MarshalAs(UnmanagedType.BStr)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        ITaskFolder GetFolder([MarshalAs(UnmanagedType.BStr)] string Path);
        [return: MarshalAs(UnmanagedType.Interface)]
        ITaskFolderCollection GetFolders(int flags);
        [return: MarshalAs(UnmanagedType.Interface)]
        ITaskFolder CreateFolder([In, MarshalAs(UnmanagedType.BStr)] string subFolderName, [In, Optional, MarshalAs(UnmanagedType.Struct)] object sddl);
        void DeleteFolder([MarshalAs(UnmanagedType.BStr)] string subFolderName, [In] int flags);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRegisteredTask GetTask([MarshalAs(UnmanagedType.BStr)] string Path);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRegisteredTaskCollection GetTasks(int flags);
        void DeleteTask([In, MarshalAs(UnmanagedType.BStr)] string Name, [In] int flags);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRegisteredTask RegisterTask([In, MarshalAs(UnmanagedType.BStr)] string Path, [In, MarshalAs(UnmanagedType.BStr)] string XmlText, [In] int flags, [In, MarshalAs(UnmanagedType.Struct)] object UserId, [In, MarshalAs(UnmanagedType.Struct)] object password, [In] TaskLogonType LogonType, [In, Optional, MarshalAs(UnmanagedType.Struct)] object sddl);
        [return: MarshalAs(UnmanagedType.Interface)]
        IRegisteredTask RegisterTaskDefinition([In, MarshalAs(UnmanagedType.BStr)] string Path, [In, MarshalAs(UnmanagedType.Interface)] ITaskDefinition pDefinition, [In] int flags, [In, MarshalAs(UnmanagedType.Struct)] object UserId, [In, MarshalAs(UnmanagedType.Struct)] object password, [In] TaskLogonType LogonType, [In, Optional, MarshalAs(UnmanagedType.Struct)] object sddl);
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetSecurityDescriptor(int securityInformation);
        void SetSecurityDescriptor([In, MarshalAs(UnmanagedType.BStr)] string sddl, [In] int flags);
    }

    [ComImport, Guid("79184A66-8664-423F-97F1-637356A5D812"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITaskFolderCollection : IEnumerable
    {
        int Count { get; }
        ITaskFolder this[object index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
    }

    [ComImport, Guid("B4EF826B-63C3-46E4-A504-EF69E4F7EA4D"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITaskNamedValueCollection : IEnumerable
    {
        int Count { get; }
        ITaskNamedValuePair this[int index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
        [return: MarshalAs(UnmanagedType.Interface)]
        ITaskNamedValuePair Create([In, MarshalAs(UnmanagedType.BStr)] string Name, [In, MarshalAs(UnmanagedType.BStr)] string Value);
        void Remove([In] int index);
        void Clear();
    }

    [ComImport, Guid("39038068-2B46-4AFD-8662-7BB6F868D221"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity, DefaultMember("Name")]
    internal interface ITaskNamedValuePair
    {
        string Name { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string Value { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
    }

    [ComImport, TypeLibType((short)0x10c0), DefaultMember("TargetServer"), Guid("2FABA4C7-4DA9-4013-9697-20CC3FD40F85"), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITaskService
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        ITaskFolder GetFolder([In, MarshalAs(UnmanagedType.BStr)] string Path);
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        IRunningTaskCollection GetRunningTasks(int flags);
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        ITaskDefinition NewTask([In] uint flags);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        void Connect([In, Optional, MarshalAs(UnmanagedType.Struct)] object serverName, [In, Optional, MarshalAs(UnmanagedType.Struct)] object user, [In, Optional, MarshalAs(UnmanagedType.Struct)] object domain, [In, Optional, MarshalAs(UnmanagedType.Struct)] object password);
        [DispId(5)]
        bool Connected { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)] get; }
        [DispId(0)]
        string TargetServer { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)] get; }
        [DispId(6)]
        string ConnectedUser { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)] get; }
        [DispId(7)]
        string ConnectedDomain { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)] get; }
        [DispId(8)]
        uint HighestVersion { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)] get; }
    }

    [ComImport, CoClass(typeof(TaskSchedulerClass)), Guid("2FABA4C7-4DA9-4013-9697-20CC3FD40F85"), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface TaskScheduler : ITaskService
    {
    }

    [ComImport, Guid("653758FB-7B9A-4F1E-A471-BEEB8E9B834E"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity, DefaultMember("InstanceGuid")]
    internal interface IRunningTask
    {
        string Name { [return: MarshalAs(UnmanagedType.BStr)] get; }
        string InstanceGuid { [return: MarshalAs(UnmanagedType.BStr)] get; }
        string Path { [return: MarshalAs(UnmanagedType.BStr)] get; }
        TaskState State { get; }
        string CurrentAction { [return: MarshalAs(UnmanagedType.BStr)] get; }
        void Stop();
        void Refresh();
        uint EnginePID { get; }
    }

    [ComImport, Guid("6A67614B-6828-4FEC-AA54-6D52E8F1F2DB"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface IRunningTaskCollection : IEnumerable
    {
        int Count { get; }
        IRunningTask this[object index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
    }

    [ComImport, DefaultMember("TargetServer"), Guid("0F87369F-A4E5-4CFC-BD3E-73E6154572DD"), TypeLibType((short)2), ClassInterface((short)0), System.Security.SuppressUnmanagedCodeSecurity]
    internal class TaskSchedulerClass : TaskScheduler
    {
        // Methods
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        public virtual extern void Connect([In, Optional, MarshalAs(UnmanagedType.Struct)] object serverName, [In, Optional, MarshalAs(UnmanagedType.Struct)] object user, [In, Optional, MarshalAs(UnmanagedType.Struct)] object domain, [In, Optional, MarshalAs(UnmanagedType.Struct)] object password);
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        public virtual extern ITaskFolder GetFolder([In, MarshalAs(UnmanagedType.BStr)] string Path);
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        public virtual extern IRunningTaskCollection GetRunningTasks(int flags);
        [return: MarshalAs(UnmanagedType.Interface)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        public virtual extern ITaskDefinition NewTask([In] uint flags);

        // Properties
        [DispId(5)]
        public virtual extern bool Connected { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)] get; }
        [DispId(7)]
        public virtual extern string ConnectedDomain { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)] get; }
        [DispId(6)]
        public virtual extern string ConnectedUser { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)] get; }
        [DispId(8)]
        public virtual extern uint HighestVersion { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)] get; }
        [DispId(0)]
        public virtual extern string TargetServer { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)] get; }
    }

    [ComImport, Guid("8FD4711D-2D02-4C8C-87E3-EFF699DE127E"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITaskSettings
    {
        bool AllowDemandStart { get; [param: In] set; }
        string RestartInterval { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        int RestartCount { get; [param: In] set; }
        TaskInstancesPolicy MultipleInstances { get; [param: In] set; }
        bool StopIfGoingOnBatteries { get; [param: In] set; }
        bool DisallowStartIfOnBatteries { get; [param: In] set; }
        bool AllowHardTerminate { get; [param: In] set; }
        bool StartWhenAvailable { get; [param: In] set; }
        string XmlText { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        bool RunOnlyIfNetworkAvailable { get; [param: In] set; }
        string ExecutionTimeLimit { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        bool Enabled { get; [param: In] set; }
        string DeleteExpiredTaskAfter { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        int Priority { get; [param: In] set; }
        TaskCompatibility Compatibility { get; [param: In] set; }
        bool Hidden { get; [param: In] set; }
        IIdleSettings IdleSettings { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        bool RunOnlyIfIdle { get; [param: In] set; }
        bool WakeToRun { get; [param: In] set; }
    }

    [ComImport, Guid("09941815-EA89-4B5B-89E0-2A773801FAC3"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITrigger
    {
        TaskTriggerType Type { get; }
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        IRepetitionPattern Repetition { [return: MarshalAs(UnmanagedType.Interface)] get; [param: In, MarshalAs(UnmanagedType.Interface)] set; }
        string ExecutionTimeLimit { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string StartBoundary { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        string EndBoundary { [return: MarshalAs(UnmanagedType.BStr)] get; [param: In, MarshalAs(UnmanagedType.BStr)] set; }
        bool Enabled { get; [param: In] set; }
    }

    [ComImport, Guid("85DF5081-1B24-4F32-878A-D9D14DF4CB77"), InterfaceType(ComInterfaceType.InterfaceIsDual), System.Security.SuppressUnmanagedCodeSecurity]
    internal interface ITriggerCollection : IEnumerable
    {
        int Count { get; }
        ITrigger this[int index] { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        new IEnumerator GetEnumerator();
        [return: MarshalAs(UnmanagedType.Interface)]
        ITrigger Create([In] TaskTriggerType Type);
        void Remove([In, MarshalAs(UnmanagedType.Struct)] object index);
        void Clear();
    }
}
