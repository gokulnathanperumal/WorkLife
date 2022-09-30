#region Imports
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;
#endregion

namespace WorkLife
{
    public partial class Form1 : Form
    {
        #region Read Me
        //1.Create Windows Forms App 'WorkLife' and add lable1 & notifyIcon1
        //2.Download alarm-clock.ico (icon-icons.com/icon/alarm-clock/60726) and place it in root folder (Right click File > Properties > Build Action > Embedded resource)
        //2.Choose alarm-clock Icon for notifyIcon1 and Project
        //3.Create WorkLife.exe shortcut in C:\Users\<username>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
        #endregion

        #region Windows API
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const UInt32 SWP_NOSIZE = 0x0001;
        private const UInt32 SWP_NOMOVE = 0x0002;
        private const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE;
        #endregion

        #region Fields
        string currentDate;
        string currentProcessName;
        string currentWindowTitle;
        ProcessConfig process;
        HoursLog hours;
        int dayHourCount;
        bool showDaySummary;
        #endregion

        #region Constructor and Events
        public Form1()
        {
            InitializeComponent();

            //WindowState Events
            this.Load += new System.EventHandler(this.Form1_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.notifyIcon1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseClick);

            this.label1.AutoSize = false;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Text = "Loading...";
            this.label1.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.DoubleClick += new System.EventHandler(this.label1_DoubleClick);

            this.BackColor = System.Drawing.Color.Black;
            this.ForeColor = System.Drawing.Color.Green;

            InitializeDay();

            Timer timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += TimerTick;
            timer.Start();
        }

        private void TimerTick(object sender, EventArgs e)
        {
            if (currentDate != DateTime.Now.ToString("yyyy-MM-dd"))
            {
                InitializeDay();
            }

            if (hours.TotalSeconds % 60 == 0)
            {
                WriteHoursLog();
            }

            IntPtr hwnd = GetForegroundWindow();
            string activeWindowTitle = GetActiveWindowTitle(hwnd);
            Process activeProcess = GetActiveProcess(hwnd);
            UpdateHours(activeProcess, activeWindowTitle);
            
            if (currentProcessName != activeProcess.ProcessName || currentWindowTitle != activeWindowTitle)
            {
                WriteActiveWindowLog(activeProcess, activeWindowTitle);
                currentProcessName = activeProcess.ProcessName;
                currentWindowTitle = activeWindowTitle;
            }

            if (hours.WorkSeconds > 0 && showDaySummary)
            {
                SummaryBalloonTip();
                showDaySummary = false;
            }

            if (hours.WorkSeconds / 3600 > dayHourCount)
            {
                HourlyBalloonTip();
                dayHourCount = hours.WorkSeconds / 3600;
            }
        }

        private void label1_DoubleClick(object sender, EventArgs e)
        {
            SummaryBalloonTip();
        }
        #endregion

        #region WindowState Events and Methods
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "WorkLife";
            this.notifyIcon1.Text = "WorkLife";
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;

            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS);

            WindowMinimize();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ReadWriteProcessCategory();

            if (e.CloseReason == CloseReason.UserClosing)
            {
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
                e.Cancel = true;
            }
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            ReadWriteProcessCategory();

            if (this.WindowState == FormWindowState.Minimized)
            {
                WindowNormal();
            }
            else
            {
                WindowMinimize();
            }
        }

        private void WindowNormal()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;

            Graphics graphics = this.CreateGraphics();
            float scalingFactorX = graphics.DpiX / 96;
            float scalingFactorY = graphics.DpiY / 96;

            Screen screen = Screen.FromPoint(this.Location);
            this.Width = (int)(250 * scalingFactorX);
            this.Height = (int)(300 * scalingFactorY);
            this.Location = new Point(screen.WorkingArea.Right - this.Width, screen.WorkingArea.Bottom - this.Height);
        }

        private void WindowMinimize()
        {
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }
        #endregion

        #region Models
        public class ProcessConfig
        {
            public List<string> DevProcess { get; set; } = new List<string>() { "devenv", "chrome", "*" };
            public List<string> TeamProcess { get; set; } = new List<string>() { "Teams", "OUTLOOK" };
            public List<string> MeetingProcess { get; set; } = new List<string>() { "Zoom", "CiscoCollabHost" };
            public List<string> IdleProcess { get; set; } = new List<string>() { "Idle", "LockApp", "WorkLife" };
        }

        public class HoursLog
        {
            public int TotalSeconds { get; set; } = 0;
            public int WorkSeconds { get; set; } = 0;
            public int IdleSeconds { get; set; } = 0;
            public int DevSeconds { get; set; } = 0;
            public int TeamSeconds { get; set; } = 0;
            public int MeetingSeconds { get; set; } = 0;
        }

        public class ActiveWindowLog
        {
            public string DateTime { get; set; }
            public string ProcessName { get; set; }
            public int Id { get; set; }
            public string MainWindowTitle { get; set; }
            public string WindowTitle { get; set; }
        }
        #endregion

        #region Methods
        private void InitializeDay()
        {
            currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            currentProcessName = null;
            currentWindowTitle = null;
            ReadWriteProcessCategory();
            ReadHoursLog();
            dayHourCount = 0;
            showDaySummary = true;
        }

        private void ReadWriteProcessCategory()
        {
            try
            {
                string currentDirectory = System.IO.Directory.GetCurrentDirectory();
                string path = Path.Combine(currentDirectory, "ProcessConfig.json");
                string config = File.Exists(path) ? File.ReadAllText(path) : null;
                if (config != null)
                {
                    var jsonOptions = new JsonSerializerOptions { AllowTrailingCommas = true };
                    process = JsonSerializer.Deserialize<ProcessConfig>(config, jsonOptions);
                }
                else
                {
                    process = new ProcessConfig();

                    var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                    string newConfig = JsonSerializer.Serialize(process, jsonOptions);
                    File.WriteAllText(path, newConfig);
                }
            }
            catch
            {
                notifyIcon1.BalloonTipTitle = "Process Config Error";
                notifyIcon1.BalloonTipText = @"Failed to read/write ProcessConfig.json";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
                notifyIcon1.ShowBalloonTip(5000);
            }
        }

        private void ReadHoursLog()
        {
            try
            {
                string currentDirectory = System.IO.Directory.GetCurrentDirectory();
                string path = Path.Combine(currentDirectory, @"HoursLog\HoursLog_" + currentDate + ".json");
                string log = File.Exists(path) ? File.ReadAllText(path) : null;
                if (log != null)
                {
                    var jsonOptions = new JsonSerializerOptions { AllowTrailingCommas = true };
                    hours = JsonSerializer.Deserialize<HoursLog>(log, jsonOptions);
                }
                else
                {
                    hours = new HoursLog();
                }
            }
            catch
            {
                notifyIcon1.BalloonTipTitle = "Hour Log Error";
                notifyIcon1.BalloonTipText = @"Failed to read HoursLog_" + currentDate + ".json";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
                notifyIcon1.ShowBalloonTip(5000);
            }
        }

        private void UpdateHours(Process activeProcess, string activeWindowTitle)
        {
            hours.TotalSeconds += 1;

            if (process.DevProcess.Contains(activeProcess.ProcessName))
            {
                hours.DevSeconds += 1;
                hours.WorkSeconds += 1;
            }
            else if (process.TeamProcess.Contains(activeProcess.ProcessName))
            {
                hours.TeamSeconds += 1;
                hours.WorkSeconds += 1;
            }
            else if (process.MeetingProcess.Contains(activeProcess.ProcessName))
            {
                hours.MeetingSeconds += 1;
                hours.WorkSeconds += 1;
            }
            else if (process.IdleProcess.Contains(activeProcess.ProcessName))
            {
                hours.IdleSeconds += 1;
            }
            else
            {
                if (process.DevProcess.Contains("*"))
                {
                    hours.DevSeconds += 1;
                    hours.WorkSeconds += 1;
                }
                else if (process.TeamProcess.Contains("*"))
                {
                    hours.TeamSeconds += 1;
                    hours.WorkSeconds += 1;
                }
                else if (process.MeetingProcess.Contains("*"))
                {
                    hours.MeetingSeconds += 1;
                    hours.WorkSeconds += 1;
                }
                else
                {
                    hours.IdleSeconds += 1;
                }
            }

            string message = string.Format("{0}: {1} \n\n\n{2,-14}: {3} \n\n{4,-14}: {5} \n\n{6,-14}: {7} \n\n\n{8,-14}: {9} \n\n{10,-14}: {11} \n\n{12,-14}: {13}",
                "Current Process", activeProcess.ProcessName,
                "Total Hours", TimeSpan.FromSeconds(hours.TotalSeconds).ToString(@"hh\:mm\:ss"),
                "Work Hours", TimeSpan.FromSeconds(hours.WorkSeconds).ToString(@"hh\:mm\:ss"),
                "Idle Hours", TimeSpan.FromSeconds(hours.IdleSeconds).ToString(@"hh\:mm\:ss"),
                "Dev Hours", TimeSpan.FromSeconds(hours.DevSeconds).ToString(@"hh\:mm\:ss"),
                "Team Hours", TimeSpan.FromSeconds(hours.TeamSeconds).ToString(@"hh\:mm\:ss"),
                "Meeting Hours", TimeSpan.FromSeconds(hours.MeetingSeconds).ToString(@"hh\:mm\:ss"));

            label1.Text = message;
            this.notifyIcon1.Text = TimeSpan.FromSeconds(hours.WorkSeconds).ToString(@"hh\:mm\:ss");
        }

        private void WriteActiveWindowLog(Process activeProcess, string activeWindowTitle)
        {
            try
            {
                ActiveWindowLog log = new ActiveWindowLog()
                {
                    DateTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"),
                    ProcessName = activeProcess.ProcessName,
                    Id = activeProcess.Id,
                    MainWindowTitle = activeProcess.MainWindowTitle,
                    WindowTitle = activeWindowTitle
                };

                string json = JsonSerializer.Serialize(log);
                string currentDirectory = System.IO.Directory.GetCurrentDirectory();
                string path = Path.Combine(currentDirectory, @"ActiveWindowLog\ActiveWindowLog_" + currentDate + ".json");
                if (!Directory.Exists(Path.Combine(currentDirectory, "ActiveWindowLog")))
                {
                    Directory.CreateDirectory(Path.Combine(currentDirectory, "ActiveWindowLog"));
                }
                File.AppendAllLines(path, new List<string>() { json });
            }
            catch
            {
                notifyIcon1.BalloonTipTitle = "Active Window Log Error";
                notifyIcon1.BalloonTipText = @"Failed to write ActiveWindowLog_" + currentDate + ".json";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
                notifyIcon1.ShowBalloonTip(5000);
            }
        }

        private void WriteHoursLog()
        {
            try
            {
                string json = JsonSerializer.Serialize(hours);
                string currentDirectory = System.IO.Directory.GetCurrentDirectory();
                string path = Path.Combine(currentDirectory, @"HoursLog\HoursLog_" + currentDate + ".json");
                if (!Directory.Exists(Path.Combine(currentDirectory, "HoursLog")))
                {
                    Directory.CreateDirectory(Path.Combine(currentDirectory, "HoursLog"));
                }
                File.WriteAllText(path, json);
            }
            catch
            {
                notifyIcon1.BalloonTipTitle = "Hour Log Error";
                notifyIcon1.BalloonTipText = @"Failed to write HoursLog_" + currentDate + ".json";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
                notifyIcon1.ShowBalloonTip(5000);
            }
        }

        private string GetActiveWindowTitle(IntPtr hwnd)
        {
            string activeWindowTitle = string.Empty;
            int textLength = GetWindowTextLength(hwnd) + 1;
            StringBuilder stringBuilder = new StringBuilder(textLength);
            if (GetWindowText(hwnd, stringBuilder, textLength) > 0)
            {
                activeWindowTitle = stringBuilder.ToString();
            }
            return activeWindowTitle;
        }

        private Process GetActiveProcess(IntPtr hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            Process activeProcess = Process.GetProcessById((int)pid);
            return activeProcess;
        }

        private void HourlyBalloonTip()
        {
            notifyIcon1.BalloonTipTitle = $"{hours.WorkSeconds / 3600} Work Hours Completed";
            notifyIcon1.BalloonTipText = $"Dev Hours: {hours.DevSeconds / 3600} hr {(hours.DevSeconds % 3600) / 60} min"
                + $"\nTeam Hours: {hours.TeamSeconds / 3600} hr {(hours.TeamSeconds % 3600) / 60} min"
                + $"\nMeeting Hours: {hours.MeetingSeconds / 3600} hr {(hours.MeetingSeconds % 3600) / 60} min";
            notifyIcon1.BalloonTipIcon = (hours.WorkSeconds / 3600) > 8 ? ToolTipIcon.Warning : ToolTipIcon.Info;
            notifyIcon1.ShowBalloonTip(5000);
        }

        private void SummaryBalloonTip()
        {
            notifyIcon1.BalloonTipTitle = "Work Hours Summary";
            notifyIcon1.BalloonTipText = $"Yesterday: {GetHoursLog("Yesterday").WorkSeconds / 3600} hr {(GetHoursLog("Yesterday").WorkSeconds % 3600) / 60} min"
                + $"\nThis Week: {GetHoursLog("ThisWeek").WorkSeconds / 3600} hr {(GetHoursLog("ThisWeek").WorkSeconds % 3600) / 60} min"
                + $"\nLast Week: {GetHoursLog("LastWeek").WorkSeconds / 3600} hr {(GetHoursLog("LastWeek").WorkSeconds % 3600) / 60} min"
                + $"\nThis Month: {GetHoursLog("ThisMonth").WorkSeconds / 3600} hr {(GetHoursLog("ThisMonth").WorkSeconds % 3600) / 60} min"
                + $"\nLast Month: {GetHoursLog("LastMonth").WorkSeconds / 3600} hr {(GetHoursLog("LastMonth").WorkSeconds % 3600) / 60} min";
            notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
            notifyIcon1.ShowBalloonTip(5000);
        }

        private HoursLog GetHoursLog(string category)
        {
            HoursLog hoursLog = new HoursLog();

            var dayOfWeek = (int)DateTime.Now.DayOfWeek;
            var thisMonday = DateTime.Now.AddDays(dayOfWeek == 0 ? -6 : -(dayOfWeek - 1)).Date;
            var thisMonthStart = DateTime.Now.AddDays(-(DateTime.Now.Day - 1)).Date;
            var today = DateTime.Now.Date;

            var startDate = today;
            var endDate = today;

            if (category == "Yesterday")
            {
                startDate = today.AddDays(-1);
                endDate = today.AddDays(-1);
            }
            else if (category == "ThisWeek")
            {
                startDate = thisMonday;
                endDate = today;
            }
            else if (category == "LastWeek")
            {
                startDate = thisMonday.AddDays(-7);
                endDate = thisMonday.AddDays(-1);
            }
            else if (category == "ThisMonth")
            {
                startDate = thisMonthStart;
                endDate = today;
            }
            else if (category == "LastMonth")
            {
                startDate = thisMonthStart.AddMonths(-1);
                endDate = thisMonthStart.AddDays(-1);
            }

            for (DateTime i = startDate; i <= endDate; i = i.AddDays(1))
            {
                var hours = GetHoursLog(i);

                if (hours != null)
                {
                    hoursLog.TotalSeconds += hours.TotalSeconds;
                    hoursLog.WorkSeconds += hours.WorkSeconds;
                    hoursLog.IdleSeconds += hours.IdleSeconds;
                    hoursLog.DevSeconds += hours.DevSeconds;
                    hoursLog.TeamSeconds += hours.TeamSeconds;
                    hoursLog.MeetingSeconds += hours.MeetingSeconds;
                }
            }

            return hoursLog;
        }

        private HoursLog GetHoursLog(DateTime date)
        {
            HoursLog hoursLog = null;

            try
            {
                string currentDirectory = System.IO.Directory.GetCurrentDirectory();
                string path = Path.Combine(currentDirectory, @"HoursLog\HoursLog_" + date.ToString("yyyy-MM-dd") + ".json");
                string log = File.Exists(path) ? File.ReadAllText(path) : null;
                if (log != null)
                {
                    hoursLog = JsonSerializer.Deserialize<HoursLog>(log);
                }
            }
            catch (Exception) { }

            return hoursLog;
        }
        #endregion 
    }
}