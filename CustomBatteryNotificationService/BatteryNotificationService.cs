using System;
using System.Diagnostics;
using System.Management;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Speech.Synthesis;
using System.Threading;
using System.Threading.Tasks;

namespace BatteryNotificationService
{
    public partial class BatteryNotificationService : ServiceBase
    {
        private int _eventId = 1;
        private readonly string _logSource = "CustomBatteryNotificationService";
        private readonly string _logName = "CustomBatteryNotificationServiceLog";
        private readonly string _appName = "[CustomBatteryNotficationService]";
        //private Timer _stateTimer;
        private readonly TimeSpan _secondsOnStandby = TimeSpan.FromSeconds(5);
        private readonly TimeSpan _secondsCharging = TimeSpan.FromSeconds(0.5);
        private Timer _stateTImer;

        private bool _isCharging = false;
        private bool _fullyChargedMessageSent = false;
        private bool _lowBatteryMessageSent = false;
        private bool _batteryChargingLogged = false;
        private bool _timerInChargingState = false;

        public BatteryNotificationService()
        {
            InitializeComponent();
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists(_logSource))
            {
                EventLog.CreateEventSource(
                    _logSource, _logName);
            }
            else
            {
                //remove existing source and create a new one.
                EventLog.DeleteEventSource(_logSource);

                EventLog.CreateEventSource(
                    _logSource, _logName);
            }

            eventLog1.Source = _logSource;
            eventLog1.Log = _logName;
        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry("[CustomBatteryNotificationService] service started!", EventLogEntryType.Information, _eventId);
            RunPowerShellMsg("Service started successfully.");

            TimerCallback checkBatteryStatusCallback = new(CheckBatteryStatus);
            _stateTImer = new(checkBatteryStatusCallback, null, _secondsOnStandby, _secondsOnStandby);

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("service stopped.");
        }

        public async void CheckBatteryStatus(object sender)
        {
            ManagementClass systemBatteryInstance = new("Win32_Battery");
            ManagementObjectCollection systemBatteryProperties = systemBatteryInstance.GetInstances();
            (string, string) battery = ("", ""); //batteryPercentage, batteryStatus

            await Task.Run(() =>
            {
                foreach (ManagementObject p in systemBatteryProperties)
                {
                    battery.Item1 = p.GetPropertyValue("EstimatedChargeRemaining").ToString();
                    battery.Item2 = p.GetPropertyValue("BatteryStatus").ToString();
                    //BatteryStatus = 2 == charging
                    //BatteryStatus = 1 == standby
                }
            });
            
            int.TryParse(battery.Item1, out int batteryPercentage);

            if (battery.Item2 is "2") _isCharging = true;
            else _isCharging = false;

            if (_eventId == 9999)
            {
                eventLog1.Clear();
                _eventId = 0;
            }

            if (!_isCharging)
            {
                _batteryChargingLogged = false;
                eventLog1.WriteEntry($"Laptop battery is {(battery.Item2 is not "2" ? "ON STANDBY" : "CHARGING")} with {battery.Item1}% remaining.", EventLogEntryType.Information, _eventId++);
            }
            else if (_isCharging && !_batteryChargingLogged)
            {
                _batteryChargingLogged = true;
                eventLog1.WriteEntry($"Laptop battery is CHARGING with {battery.Item1}% remaining.", EventLogEntryType.Information, _eventId++);
            }

            if ( !_isCharging) //is not charging and low battery
            {
                if (batteryPercentage <= 40 && !_lowBatteryMessageSent)
                {
                    _lowBatteryMessageSent = true;
                    eventLog1.WriteEntry($"Low battery, {battery.Item1}% remaining. Please charge.", EventLogEntryType.Information, _eventId++);
                    RunPowerShellMsg($"Low battery, please plug in your charger.");
                    PlaySystemBeepSound();
                }
                else // is not charging and above 40 percent
                {
                    _fullyChargedMessageSent = false;
                    if (_timerInChargingState)
                    {
                        _timerInChargingState = false;
                        _stateTImer.Change(_secondsOnStandby, _secondsOnStandby);
                    }
                }
            }

            if (_isCharging)// is charging
            {
                if (!_timerInChargingState)
                {
                    _timerInChargingState = true;
                    _stateTImer.Change(_secondsCharging, _secondsCharging);
                }
                if (batteryPercentage == 94 && !_fullyChargedMessageSent)//is charging and fully charged
                {
                    _fullyChargedMessageSent = true;
                    RunPowerShellMsg("Your laptop battery is fully charged. Please unplug your charger.");
                    RunTextToSpeech("Battery Fully Charged");
                }
            }
        }

        //Was going to make a separate handler to check battery percentage
        //still has no idea how to do it though.
        /* public void CheckBatteryPercentage(object sender, ElapsedEventArgs args)
         {
             ManagementClass systemBatteryInstance = new("Win32_Battery");
             ManagementObjectCollection systemBatteryProperties = systemBatteryInstance.GetInstances();

             //item1 = battery percentage
             //item2 = battery status
             (string, string) battery = ("", "");

             //List<Property> propertiesToBind = new();
             foreach (ManagementObject p in systemBatteryProperties)
             {
                 battery.Item1 = p.GetPropertyValue("EstimatedChargeRemaining").ToString();
                 battery.Item2 = p.GetPropertyValue("BatteryStatus").ToString();
                 //BatteryStatus = 2 == charging
                 //BatteryStatus = 1 == standby
             }

             eventLog1.WriteEntry("Laptop battery is " + batteryStatus is not "2" ? "ON STANDBY" : "CHARGING", EventLogEntryType.Information, _eventId++);
             if (batteryStatus is "2")
             {
                 Timer timer = new()
                 {
                     Interval = 5000 // 5 seconds
                 };
                 timer.Elapsed += new ElapsedEventHandler(CheckBatteryStatus);
                 timer.Start();
             }
         }*/

        private class Property
        {
            internal string Name { get; set; }
            internal string Value { get; set; }
        }

        //All the random enums and services are still put in one file..
        #region Service State
        private enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);
        #endregion

        #region run powershell commands
        private void RunPowerShellMsg(string message)
        {
            PowerShell ps = PowerShell.Create();

            message = _appName + " " + message;
            message = message.Replace("True message", "");

            ps.AddScript($"msg * {message} | Out-Null");
            ps.Invoke();
        }

        private void RunTextToSpeech(string textToSpeech)
        {
            try
            {
                SpeechSynthesizer synth = new();
                synth.SelectVoice("Microsoft Zira Desktop");
                synth.Volume = 5;
                synth.Rate = 1;
                synth.Speak(textToSpeech);
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.ToString(), EventLogEntryType.Error, _eventId++);
                PlaySystemBeepSound();
            }
        }

        private void PlaySystemBeepSound()
        {
            PowerShell ps = PowerShell.Create();
            ps.AddScript("[System.Media.SystemSounds]::Beep.Play()");
            ps.Invoke();
        }
        #endregion
    }
}
