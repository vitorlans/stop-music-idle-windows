using System.ComponentModel;
using System.Timers;
using System.Runtime.InteropServices;

namespace StopSoundWin
{
    public partial class Form1 : Form
    {
        //EventLog eventLog1 = new EventLog();

        [DllImport("user32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);

        [StructLayout(LayoutKind.Sequential)]
        struct LASTINPUTINFO
        {
            public static readonly int SizeOf =
                   Marshal.SizeOf(typeof(LASTINPUTINFO));

            [MarshalAs(UnmanagedType.U4)]
            public int cbSize;
            [MarshalAs(UnmanagedType.U4)]
            public int dwTime;
        }

        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int VK_MEDIA_PREV_TRACK = 0xB1;
        public const int VK_MEDIA_STOP = 0xB2;

        public const int VK_VOLUME_DOWN = 0xAE;

        private int eventId = 1;

        public static int IdleTime() //In seconds
        {
            int idleTime = 0;
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = Marshal.SizeOf(lastInputInfo);
            lastInputInfo.dwTime = 0;

            int envTicks = Environment.TickCount;

            if (GetLastInputInfo(ref lastInputInfo))
            {
                int lastInputTick = lastInputInfo.dwTime;
                idleTime = envTicks - lastInputTick;
            }

            int a;

            if (idleTime > 0)
                a = idleTime / 1000;
            else
                a = idleTime;

            return a;
        }


        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            string log = $"Monitoring the System " +
                $"\nIdleTime: {IdleTime()} " +
                $"\nCurrent time: {DateTimeOffset.Now} " +
                $"\nLast input time: {InputTimer.GetLastInputTime()} " +
                $"\nIdle time: {InputTimer.GetInputIdleTime()}" +
                $"\nIs Windows Playing Sound: {IsWindowsPlayingSound()} ";


            Invoke(new Action(() =>
                        {
                            label1.Text = log;
                        }));

            //eventLog1.WriteEntry(log,
            //    EventLogEntryType.Information, eventId++);

            if (IdleTime() > 120 && IsWindowsPlayingSound())
            {
                keybd_event(VK_MEDIA_STOP, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);

                Invoke(new Action(() =>
                {
                    label2.Text = $"Latest Sound Off: {DateTimeOffset.Now}";
                }));
            }
        }

        public static class InputTimer
        {
            public static TimeSpan GetInputIdleTime()
            {
                var plii = new NativeMethods.LastInputInfo();
                plii.cbSize = (UInt32)Marshal.SizeOf(plii);

                if (NativeMethods.GetLastInputInfo(ref plii))
                {
                    return TimeSpan.FromMilliseconds(Environment.TickCount - plii.dwTime);
                }
                else
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
            }

            public static DateTimeOffset GetLastInputTime()
            {
                return DateTimeOffset.Now.Subtract(GetInputIdleTime());
            }

            private static class NativeMethods
            {
                public struct LastInputInfo
                {
                    public UInt32 cbSize;
                    public UInt32 dwTime;
                }

                [DllImport("user32.dll")]
                public static extern bool GetLastInputInfo(ref LastInputInfo plii);
            }
        }


        public Form1()
        {
            InitializeComponent();
            //eventLog1 = new System.Diagnostics.EventLog();
            //if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            //{
            //    System.Diagnostics.EventLog.CreateEventSource(
            //        "MySource", "MyNewLog");
            //}
            //eventLog1.Source = "MySource";
            //eventLog1.Log = "MyNewLog";
        }

        System.Timers.Timer timer = new System.Timers.Timer();


        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            //notifyIcon1.ContextMenuStrip.Items.Add("Configuration", null, new EventHandler(ShowConfig));
            notifyIcon1.ContextMenuStrip.Items.Add("Exit", null, new EventHandler(Exit));
            notifyIcon1.ContextMenuStrip.Items.Add("Pause", null, new EventHandler(Pause));


            //eventLog1.WriteEntry("In OnStart.");

            // Set up a timer that triggers every time.
            timer.Interval = 20000; // 20 seconds
            timer.Elapsed += new ElapsedEventHandler(OnTimer);
            timer.Start();

        }
        void ShowConfig(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
        void Pause(object sender, EventArgs e)
        {

            if (timer.Enabled)
            {
                notifyIcon1.ContextMenuStrip.Items[1].Text = "Unpause";
                label1.Text = "Watch Paused";
                timer.Stop();
            }
            else
            {
                notifyIcon1.ContextMenuStrip.Items[1].Text = "Pause";
                label1.Text = "";
                timer.Start();
            }
        }

        void Exit(object sender, EventArgs e)
        {
            // We must manually tidy up and remove the icon before we exit.
            // Otherwise it will be left behind until the user mouses over.
            notifyIcon1.Visible = false;
            Application.Exit();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.WindowsShutDown
                || e.CloseReason == CloseReason.ApplicationExitCall
                || e.CloseReason == CloseReason.TaskManagerClosing)
            {
                return;
            }
            e.Cancel = true;
            //assuming you want the close-button to only hide the form, 
            //and are overriding the form's OnFormClosing method:
            this.Hide();
        }

        public static bool IsWindowsPlayingSound()
        {
            var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            var speakers = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            var meter = (IAudioMeterInformation)speakers.Activate(typeof(IAudioMeterInformation).GUID, 0, IntPtr.Zero);
            var value = meter.GetPeakValue();

            // this is a bit tricky. 0 is the official "no sound" value
            // but for example, if you open a video and plays/stops with it (w/o killing the app/window/stream),
            // the value will not be zero, but something really small (around 1E-09)
            // so, depending on your context, it is up to you to decide
            // if you want to test for 0 or for a small value
            return value > 1E-08;
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator
        {
        }

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        private interface IMMDeviceEnumerator
        {
            void NotNeeded();
            IMMDevice GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        private interface IMMDevice
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")]
        private interface IAudioMeterInformation
        {
            float GetPeakValue();
            // the rest is not defined/needed
        }
    }


}