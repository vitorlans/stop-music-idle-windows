using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace IdleStopSound
{
    public partial class MyNewService : ServiceBase
    {
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

        public MyNewService()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";
        }
        private LowLevelKeyboardListener _listener;
        List<string> keys = new List<string>();

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart.");

            // Set up a timer that triggers every minute.
            Timer timer = new Timer();
            timer.Interval = 10000; // 10 seconds
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
            _listener = new LowLevelKeyboardListener();
            _listener.OnKeyPressed += _listener_OnKeyPressed;

            _listener.HookKeyboard();
        }


        void _listener_OnKeyPressed(object sender, KeyPressedArgs e)
        {
            keys.Add(e.KeyPressed.ToString());
        }

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


            eventLog1.WriteEntry($"Monitoring the System " +
                $"\nIdleTime: {IdleTime()} " +
                $"\nCurrent time: {DateTimeOffset.Now} " +
                $"\nLast input time: {InputTimer.GetLastInputTime()} " +
                $"\nIdle time: {InputTimer.GetInputIdleTime()}" +
                $"\nKEYS: {String.Join(", ", keys)}",
                EventLogEntryType.Information, eventId++);

            //if (IdleTime() > 120)
            //{
            //    keybd_event(VK_VOLUME_DOWN, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            //}
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");
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


        public class LowLevelKeyboardListener
        {
            private const int WH_KEYBOARD_LL = 13;
            private const int WM_KEYDOWN = 0x0100;
            private const int WM_SYSKEYDOWN = 0x0104;

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);

            public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

            public event EventHandler<KeyPressedArgs> OnKeyPressed;

            private LowLevelKeyboardProc _proc;
            private IntPtr _hookID = IntPtr.Zero;

            public LowLevelKeyboardListener()
            {
                _proc = HookCallback;
            }

            public void HookKeyboard()
            {
                _hookID = SetHook(_proc);
            }

            public void UnHookKeyboard()
            {
                UnhookWindowsHookEx(_hookID);
            }

            private IntPtr SetHook(LowLevelKeyboardProc proc)
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }

            private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN)
                {
                    int vkCode = Marshal.ReadInt32(lParam);

                    if (OnKeyPressed != null) { OnKeyPressed(this, new KeyPressedArgs(KeyInterop.KeyFromVirtualKey(vkCode))); }
                }

                return CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }

        public class KeyPressedArgs : EventArgs
        {
            public Key KeyPressed { get; private set; }

            public KeyPressedArgs(Key key)
            {
                KeyPressed = key;
            }
        }
    }
}
