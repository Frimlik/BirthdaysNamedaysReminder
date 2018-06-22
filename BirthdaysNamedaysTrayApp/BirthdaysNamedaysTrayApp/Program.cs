using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BirthdaysNamedaysTrayApp
{
    public class SysTrayApp : Form
    {
        [STAThread]
        public static void Main()
        {
            Application.Run(new SysTrayApp());
        }

        private static string appFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\BirthdaysNamedays";
        private string logFile = appFolder + @"\log.txt";
        private string birthdaysFile = appFolder + @"\birthdays.txt";
        private string namedaysFile = appFolder + @"\namedays.txt";
        private string outputFile = appFolder + @"\output.txt";
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
            
        public SysTrayApp()
        {
            // Create log file
            if (!Directory.Exists(appFolder))
            {
                Directory.CreateDirectory(appFolder);
            }

            using (StreamWriter sr = new StreamWriter(logFile, true))
            {
                sr.WriteLine("{0}: Application started.", DateTime.Now);
            }

            // Create a simple tray menu with only one item.
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);
            
            trayIcon = new NotifyIcon();
            trayIcon.BalloonTipIcon = ToolTipIcon.Info;
            trayIcon.BalloonTipTitle = "Oslavenci:";
            trayIcon.Text = "Narozeniny a svátky";
            trayIcon.Icon = new Icon(@"C:\Users\Jirka\Pictures\Pokeball.ico");

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            trayIcon.Click += OnTrayIconClick;

            SystemEvents.PowerModeChanged += OnPowerChange;
            SystemEvents.SessionSwitch += OnSessionSwitch;
            trayIcon.BalloonTipClicked += new EventHandler(trayIcon_BalloonTipClicked);

            CheckBirthdaysNamedays();
        }

        private void trayIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            Process.Start(outputFile);
        }

        private void OnTrayIconClick(object sender, EventArgs e)
        {
            CheckBirthdaysNamedays();
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                trayIcon.Dispose();
            }

            base.Dispose(isDisposing);
        }

        // Happens at sleep/hibernate/power off
        private void OnPowerChange(object s, PowerModeChangedEventArgs e)
        {
            string message = "";
            switch (e.Mode)
            {
                case PowerModes.Resume:
                    message = string.Format("{0}: System resumed. Application continued.", DateTime.Now);
                    break;
                case PowerModes.Suspend:
                    message = string.Format("{0}: System suspended. Application paused.", DateTime.Now);
                    break;
            }

            using (StreamWriter sr = new StreamWriter(logFile, true))
            {
                sr.WriteLine(message);
            }
        }

        // Happens at login or logout
        private void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            string message = "";
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                //I left my desk
                message = string.Format("{0}: I left my desk", DateTime.Now);
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                //I returned to my desk
                message = string.Format("{0}: I returned to my desk", DateTime.Now);
                CheckBirthdaysNamedays();
            }

            using (StreamWriter sr = new StreamWriter(logFile, true))
            {
                sr.WriteLine(message);
            }
        }

        private void CheckBirthdaysNamedays()
        {
            StringBuilder birthdays = new StringBuilder();
            StringBuilder namedays = new StringBuilder();
            StringBuilder completeText = new StringBuilder();

            string line;
            string[] data;
            DateTime today = DateTime.Today;
            bool somebodyHasBirthday = false;
            using (StreamReader sr = new StreamReader(birthdaysFile))
            {
                DateTime birthDate;
                string name, surname, nick;

                StringBuilder yesterdayBirth = new StringBuilder();
                StringBuilder todayBirth = new StringBuilder();
                StringBuilder tomorrowBirth = new StringBuilder();
                StringBuilder afterTomorrowBirth = new StringBuilder();

                while ((line = sr.ReadLine()) != null)
                {
                    data = line.Split(';');
                    birthDate = Convert.ToDateTime(data[0]);
                    name = data[1];
                    surname = data[2];
                    nick = data[3];

                    string celebrantLine;
                    if (nick != "")
                        celebrantLine = string.Format("      {0} '{1}' {2} - {3}", name, nick, surname, (birthDate.Year == 0) ? "neznámý věk" : (today.Year - birthDate.Year).ToString());
                    else
                        celebrantLine = string.Format("      {0} {1} - {2}", name, surname, (birthDate.Year == 0) ? "neznámý věk" : (today.Year - birthDate.Year).ToString());

                    if (birthDate.Day == today.Subtract(new TimeSpan(1, 0, 0, 0)).Day && birthDate.Month == today.Subtract(new TimeSpan(1, 0, 0, 0)).Month)
                    {
                        yesterdayBirth.Append(celebrantLine);
                        yesterdayBirth.Append(Environment.NewLine);
                    }
                    if (birthDate.Day == today.Day && birthDate.Month == today.Month)
                    {
                        todayBirth.Append(celebrantLine);
                        todayBirth.Append(Environment.NewLine);
                    }
                    else if (birthDate.Day == today.AddDays(1).Day && birthDate.Month == today.AddDays(1).Month)
                    {
                        tomorrowBirth.Append(celebrantLine);
                        tomorrowBirth.Append(Environment.NewLine);
                    }
                    else if (birthDate.Day == today.AddDays(2).Day && birthDate.Month == today.AddDays(2).Month)
                    {
                        afterTomorrowBirth.Append(celebrantLine);
                        afterTomorrowBirth.Append(Environment.NewLine);
                    }
                }

                if (yesterdayBirth.Length > 0)
                {
                    somebodyHasBirthday = true;
                    birthdays.Append("   Včera:");
                    birthdays.Append(Environment.NewLine);
                    birthdays.Append(yesterdayBirth);
                }
                if (todayBirth.Length > 0)
                {
                    somebodyHasBirthday = true;
                    birthdays.Append("   Dneska:");
                    birthdays.Append(Environment.NewLine);
                    birthdays.Append(todayBirth);
                }
                if (tomorrowBirth.Length > 0)
                {
                    somebodyHasBirthday = true;
                    birthdays.Append("   Zítra:");
                    birthdays.Append(Environment.NewLine);
                    birthdays.Append(tomorrowBirth);
                }
                if (afterTomorrowBirth.Length > 0)
                {
                    somebodyHasBirthday = true;
                    birthdays.Append("   Pozítří:");
                    birthdays.Append(Environment.NewLine);
                    birthdays.Append(afterTomorrowBirth);
                }
            }

            bool somebodyHasNameday = false;
            using (StreamReader sr = new StreamReader(namedaysFile))
            {
                DateTime nameDate;
                int day, month;
                string name, surname, nick;

                StringBuilder yesterdayName = new StringBuilder();
                StringBuilder todayName = new StringBuilder();
                StringBuilder tomorrowName = new StringBuilder();
                StringBuilder afterTomorrowName = new StringBuilder();

                while ((line = sr.ReadLine()) != null)
                {
                    data = line.Split(';');
                    day = Convert.ToInt32(data[0]);
                    month = Convert.ToInt32(data[1]);
                    nameDate = new DateTime(1, month, day);
                    name = data[2];
                    surname = data[3];
                    nick = data[4];

                    string celebrantLine;
                    if (nick != "")
                        celebrantLine = string.Format("      {0} '{1}' {2}", name, nick, surname);
                    else
                        celebrantLine = string.Format("      {0} {1}", name, surname);

                    if (nameDate.Day == today.Subtract(new TimeSpan(1, 0, 0, 0)).Day && nameDate.Month == today.Subtract(new TimeSpan(1, 0, 0, 0)).Month)
                    {
                        yesterdayName.Append(celebrantLine);
                        yesterdayName.Append(Environment.NewLine);
                    }
                    if (nameDate.Day == today.Day && nameDate.Month == today.Month)
                    {
                        todayName.Append(celebrantLine);
                        todayName.Append(Environment.NewLine);
                    }
                    else if (nameDate.Day == today.AddDays(1).Day && nameDate.Month == today.AddDays(1).Month)
                    {
                        tomorrowName.Append(celebrantLine);
                        tomorrowName.Append(Environment.NewLine);
                    }
                    else if (nameDate.Day == today.AddDays(2).Day && nameDate.Month == today.AddDays(2).Month)
                    {
                        afterTomorrowName.Append(celebrantLine);
                        afterTomorrowName.Append(Environment.NewLine);
                    }
                }

                if (yesterdayName.Length > 0)
                {
                    somebodyHasNameday = true;
                    namedays.Append("   Včera:");
                    namedays.Append(Environment.NewLine);
                    namedays.Append(yesterdayName);
                }
                if (todayName.Length > 0)
                {
                    somebodyHasNameday = true;
                    namedays.Append("   Dneska:");
                    namedays.Append(Environment.NewLine);
                    namedays.Append(todayName);
                }
                if (tomorrowName.Length > 0)
                {
                    somebodyHasNameday = true;
                    namedays.Append("   Zítra:");
                    namedays.Append(Environment.NewLine);
                    namedays.Append(tomorrowName);
                }
                if (afterTomorrowName.Length > 0)
                {
                    somebodyHasBirthday = true;
                    namedays.Append("   Pozítří:");
                    namedays.Append(Environment.NewLine);
                    namedays.Append(afterTomorrowName);
                }
            }

            if (somebodyHasBirthday)
            {
                completeText.Append("Narozeniny:");
                completeText.Append(Environment.NewLine);
                completeText.Append(birthdays);
                completeText.Append(Environment.NewLine);
            }
            if (somebodyHasNameday)
            {
                completeText.Append("Svátky:");
                completeText.Append(Environment.NewLine);
                completeText.Append(namedays);
            }
            
            if (!somebodyHasBirthday && !somebodyHasNameday)
            {
                completeText.Append("Žádný oslavenec v dohledu...");
                completeText.Append(Environment.NewLine);
                completeText.Append(Environment.NewLine);
            }

            using (StreamWriter sw = new StreamWriter(outputFile))
            {
                sw.Write(completeText.ToString());
            }

            // pokud delka textu presahuje 255 znaku (limit pro bublinu)
            // zakonci text u nejblizsiho drivejsiho oslavence (odradkovani)
            if (completeText.Length > 255)
            {
                string extraCelebrants = Environment.NewLine + Environment.NewLine + "Pro zobrazení dalších oslavenců klikni na bublinu.";
                completeText.Remove(255 - extraCelebrants.Length, completeText.Length - (255 - extraCelebrants.Length));
                int lastNewLineIndex = completeText.ToString().LastIndexOf(Environment.NewLine);
                completeText.Remove(lastNewLineIndex, completeText.Length - lastNewLineIndex);
                
                while (completeText[completeText.Length - 1] == ':')
                {
                    lastNewLineIndex = completeText.ToString().LastIndexOf(Environment.NewLine);
                    completeText.Remove(lastNewLineIndex, completeText.Length - lastNewLineIndex);
                }
                
                completeText.Append(extraCelebrants);
            }

            trayIcon.BalloonTipText = completeText.ToString();
            trayIcon.ShowBalloonTip(20000);
        }
    }
}