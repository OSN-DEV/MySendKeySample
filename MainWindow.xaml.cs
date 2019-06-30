using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Threading;

namespace MySendKeySample {
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window {

        #region Declaration
        public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);
        private static class NativeMethods {
            internal const uint MAPVK_VK_TO_VSC = 0;

            [StructLayout(LayoutKind.Sequential, Pack = 1)]
            internal struct MOUSEINPUT32_WithSkip {
                public uint __Unused0; // See INPUT32 structure

                public int X;
                public int Y;
                public uint MouseData;
                public uint Flags;
                public uint Time;
                public IntPtr ExtraInfo;
            }

            [StructLayout(LayoutKind.Sequential, Pack = 1)]
            internal struct KEYBDINPUT32_WithSkip {
                public uint __Unused0; // See INPUT32 structure

                public ushort VirtualKey;
                public ushort ScanCode;
                public uint Flags;
                public uint Time;
                public IntPtr ExtraInfo;
            }

            [StructLayout(LayoutKind.Explicit)]
            internal struct INPUT32 {
                [FieldOffset(0)]
                public uint Type;
                [FieldOffset(0)]
                public MOUSEINPUT32_WithSkip Mouse;
                [FieldOffset(0)]
                public KEYBDINPUT32_WithSkip Keyboard;
            }

            // INPUT.KI (40). vk: 8, sc: 10, fl: 12, t: 16, ex: 24
            [StructLayout(LayoutKind.Explicit, Size = 40)]
            internal struct INPUT64 {
                [FieldOffset(0)]
                public uint Type;
                [FieldOffset(8)]
                public ushort VirtualKey;
                [FieldOffset(10)]
                public ushort ScanCode;
                [FieldOffset(12)]
                public uint Flags;
                [FieldOffset(16)]
                public uint Time;
                [FieldOffset(24)]
                public IntPtr ExtraInfo;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public extern static bool EnumWindows(EnumWindowsDelegate lpEnumFunc, IntPtr lparam);

            [DllImport("user32")]
            public static extern bool IsWindowVisible(IntPtr hWnd);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetWindowTextLength(IntPtr hWnd);

            [DllImport("User32.dll")]
            internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, [Out] out uint lpdwProcessId);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("User32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("User32.dll", EntryPoint = "SendInput", SetLastError = true)]
            internal static extern uint SendInput32(uint nInputs, INPUT32[] pInputs, int cbSize);

            [DllImport("User32.dll", EntryPoint = "SendInput", SetLastError = true)]
            internal static extern uint SendInput64(uint nInputs, INPUT64[] pInputs, int cbSize);

            [DllImport("user32.dll")]
            public static extern IntPtr GetMessageExtraInfo();

            [DllImport("User32.dll")]
            internal static extern IntPtr GetKeyboardLayout(uint idThread);

            [DllImport("User32.dll")]
            private static extern uint MapVirtualKey(uint uCode, uint uMapType);
            [DllImport("User32.dll")]
            private static extern uint MapVirtualKeyEx(uint uCode, uint uMapType,IntPtr hKL);
            internal static uint MapVirtualKey3(uint uCode, uint uMapType, IntPtr hKL) {
                if (hKL == IntPtr.Zero) {
                    return MapVirtualKey(uCode, uMapType);
                } else {
                    return MapVirtualKeyEx(uCode, uMapType, hKL);
                }
            }
        }

        private static class KeyStroke {
            public const int KeyDown = 0x100;
            public const int KeyUp = 0x101;
            public const int SysKeyDown = 0x104;
            public const int SysKeyup = 0x105;
        }
        private static class Flags {
            public const int None = 0x00;
            public const int KeyDown = 0x00;
            public const int KeyUp = 0x02;
            public const int ExtendeKey = 0x01; 
            public const int Unicode = 0x04;
            public const int ScanCode = 0x08;
        }
        private class KeySet {
            public ushort VirtualKey;
            public ushort ScanCode;
            public uint Flag;
            public KeySet(byte[] pair) : this(pair, Flags.None) {
            }
            public KeySet(byte[] pair, uint flag) {
                this.VirtualKey = KeySetPair.VirtualKey(pair);
                this.ScanCode = KeySetPair.ScanCode(pair);
                this.Flag = flag;
            }
        }
        private static class ExtraInfo {
            public const int SendKey = 1;
        }

        private static class InputType {
            public const uint Mouse = 0;
            public const uint Keyboard = 1;
            public const uint Hardware = 2;
        }

        private static string _className;
        private static bool _findSelf = false;
        private static IntPtr _targetWIndows = IntPtr.Zero;
        private static MainWindow _self;
        private static IntPtr _keyboardLayout = IntPtr.Zero;
        private delegate bool SendVKeyNativeDelegate(uint keyStroke, KeySet keyset);
        #endregion

        #region Constructor
        public MainWindow() {
            InitializeComponent();
            _self = this;
            _className = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;

            uint uPID;
            uint uTID = NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out uPID);
            _keyboardLayout = NativeMethods.GetKeyboardLayout(uTID);
        }
        #endregion

        #region Event
        private void Button_Click(object sender, RoutedEventArgs e) {
            NativeMethods.EnumWindows(new EnumWindowsDelegate(EnumWindowCallBack), IntPtr.Zero);
        }
        #endregion

        #region Public Method
        public void SendKey(IntPtr hWnd) {
            NativeMethods.SetForegroundWindow(hWnd);
            DoEvents();

            SendVKeyNativeDelegate sendKey;
            if (IntPtr.Size == 4) {
                sendKey = this.SendVKeyNative32;
            } else {
                sendKey = this.SendVKeyNative64;
            }

            // https://stackoverflow.com/questions/12761169/send-keys-through-sendinput-in-user32-dll
            sendKey(KeyStroke.KeyDown, new KeySet(KeySetPair.A));
            sendKey(KeyStroke.KeyDown, new KeySet(KeySetPair.Tab));
            sendKey(KeyStroke.KeyDown, new KeySet(KeySetPair.B));
            sendKey(KeyStroke.KeyDown, new KeySet(KeySetPair.B));
            DoEvents();
            sendKey(KeyStroke.KeyDown, new KeySet(KeySetPair.Enter));
        }
        #endregion

        #region Private Method
        private static bool EnumWindowCallBack(IntPtr hWnd, IntPtr lparam) {
            if (!NativeMethods.IsWindowVisible(hWnd)) {
                return true;
            }

            int length = NativeMethods.GetWindowTextLength(hWnd);
            if (0 < length) {
                uint processId;
                NativeMethods.GetWindowThreadProcessId(hWnd, out processId);
                var proccess = Process.GetProcessById((int)processId);

                if (_findSelf) {
                    _self.SendKey(hWnd);
                    _findSelf = false;
                    return false;
                } else if (proccess.ProcessName == _className) {
                    _findSelf = true;
                }
            }
            return true;
        }

        private static void DoEvents() {
            DispatcherFrame frame = new DispatcherFrame();
            var callback = new DispatcherOperationCallback(ExitFrames);
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, callback, frame);
            Dispatcher.PushFrame(frame);
        }
        private static object ExitFrames(object obj) {
            ((DispatcherFrame)obj).Continue = false;
            return null;
        }

        private bool SendVKeyNative32(uint keyStroke, KeySet keyset) {
            var input = new NativeMethods.INPUT32();
            input.Type = InputType.Keyboard;
            input.Keyboard.Flags = keyStroke | keyset.Flag;
            input.Keyboard.VirtualKey = keyset.VirtualKey;
            // input.Keyboard.ScanCode =(ushort)(NativeMethods.MapVirtualKey3((uint)keyset.VirtualKey, NativeMethods.MAPVK_VK_TO_VSC, _keyboardLayout) & 0xFFU); ;
            input.Keyboard.ScanCode = 0;
            input.Keyboard.Time = 0;
            input.Keyboard.ExtraInfo = NativeMethods.GetMessageExtraInfo();
            NativeMethods.INPUT32[] inputs = { input };
            if (NativeMethods.SendInput32(1, inputs, Marshal.SizeOf(typeof(NativeMethods.INPUT32))) != 1)
                return false;

            return true;
        }

        private bool SendVKeyNative64(uint keyStroke, KeySet keyset) {
            var input = new NativeMethods.INPUT64();
            input.Type = InputType.Keyboard;
            input.Flags = keyStroke | keyset.Flag;
            input.VirtualKey = keyset.VirtualKey;
            // input.ScanCode = (ushort)(NativeMethods.MapVirtualKey3((uint)keyset.VirtualKey, NativeMethods.MAPVK_VK_TO_VSC, _keyboardLayout) & 0xFFU);
            input.ScanCode = 0;
            input.Time = 0;
            input.ExtraInfo = NativeMethods.GetMessageExtraInfo();
            NativeMethods.INPUT64[] inputs = { input };
            if (NativeMethods.SendInput64(1, inputs, Marshal.SizeOf(typeof(NativeMethods.INPUT64))) != 1)
                return false;

            return true;
        }
        #endregion

    }
}
