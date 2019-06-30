using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySendKeySample { 
    
    // Keyset : Virtual Key, ScanCode
    internal class KeySetPair {
        public static class Index {
            public const int VirtualKey = 0;
            public const int ScanCode = 1;
        }

        /// <summary>
        ///  get virtual key
        /// </summary>
        /// <param name="keyset">keyset</param>
        /// <returns>virtual key</returns>
        public static byte VirtualKey(byte[] keyset) {
            return keyset[Index.VirtualKey];
        }

        /// <summary>
        ///  get scan code
        /// </summary>
        /// <param name="keyset">keyset</param>
        /// <returns>scan code</returns>
        public static byte ScanCode(byte[] keyset) {
            return keyset[Index.ScanCode];
        }

        // modifiy
        // public readonly static byte[] Henkan = { 0xFF, 0x79 };      // 変換
        public readonly static byte[] Henkan = { 0xF0, 0x70 };      // 変換
        public readonly static byte[] Muhenkan = { 0xEB, 0x7B };    // 無変換
        public readonly static byte[] ShiftL = { 0xA0, 0x2A };      // 左Shift
        public readonly static byte[] ShiftR = { 0xA1, 0x36 };      // 右Shift
        public readonly static byte[] ControlL = { 0xA2, 0x1D };    // 左Control
        public readonly static byte[] AltL = { 0xA4, 0x38 };        // 左Alt

        // function
        public readonly static byte[] F1 = { 0x70, 0x3B };
        public readonly static byte[] F2 = { 0x71, 0x3C };
        public readonly static byte[] F3 = { 0x72, 0x3D };
        public readonly static byte[] F4 = { 0x73, 0x3E };
        public readonly static byte[] F5 = { 0x74, 0x3F };
        public readonly static byte[] F6 = { 0x75, 0x40 };
        public readonly static byte[] F7 = { 0x76, 0x41 };
        public readonly static byte[] F8 = { 0x77, 0x42 };
        public readonly static byte[] F9 = { 0x78, 0x43 };
        public readonly static byte[] F10 = { 0x79, 0x44 };
        public readonly static byte[] F11 = { 0x7A, 0x57 };
        public readonly static byte[] F12 = { 0x7B, 0x58 };
        public readonly static byte[] F13 = { 0x7C, 0x64 };
        public readonly static byte[] F14 = { 0x7D, 0x65 };
        public readonly static byte[] F15 = { 0x7E, 0x66 };

        // Number
        public readonly static byte[] Num0 = { 0x30, 0x0B };
        public readonly static byte[] Num1 = { 0x31, 0x02 };
        public readonly static byte[] Num2 = { 0x32, 0x03 };
        public readonly static byte[] Num3 = { 0x33, 0x04 };
        public readonly static byte[] Num4 = { 0x34, 0x05 };
        public readonly static byte[] Num5 = { 0x35, 0x06 };
        public readonly static byte[] Num6 = { 0x36, 0x07 };
        public readonly static byte[] Num7 = { 0x37, 0x08 };
        public readonly static byte[] Num8 = { 0x38, 0x09 };
        public readonly static byte[] Num9 = { 0x39, 0x0A };

        // Alpha
        public readonly static byte[] A = { 0x41, 0x1E };
        public readonly static byte[] B = { 0x42, 0x30 };
        public readonly static byte[] C = { 0x43, 0x2E };
        public readonly static byte[] D = { 0x44, 0x20 };
        public readonly static byte[] E = { 0x45, 0x12 };
        public readonly static byte[] F = { 0x46, 0x21 };
        public readonly static byte[] G = { 0x47, 0x22 };
        public readonly static byte[] H = { 0x48, 0x23 };
        public readonly static byte[] I = { 0x49, 0x17 };
        public readonly static byte[] J = { 0x4A, 0x24 };
        public readonly static byte[] K = { 0x4B, 0x25 };
        public readonly static byte[] L = { 0x4C, 0x26 };
        public readonly static byte[] M = { 0x4D, 0x32 };
        public readonly static byte[] N = { 0x4E, 0x31 };
        public readonly static byte[] O = { 0x4F, 0x18 };
        public readonly static byte[] P = { 0x50, 0x19 };
        public readonly static byte[] Q = { 0x51, 0x10 };
        public readonly static byte[] R = { 0x52, 0x13 };
        public readonly static byte[] S = { 0x53, 0x1F };
        public readonly static byte[] T = { 0x54, 0x14 };
        public readonly static byte[] U = { 0x55, 0x16 };
        public readonly static byte[] V = { 0x56, 0x2F };
        public readonly static byte[] W = { 0x57, 0x11 };
        public readonly static byte[] X = { 0x58, 0x2D };
        public readonly static byte[] Y = { 0x59, 0x15 };
        public readonly static byte[] Z = { 0x5A, 0x2C };

        // Tenkey
        public readonly static byte[] TenKeyNum0 = { 0x60, 0x52 };
        public readonly static byte[] TenKeyNum1 = { 0x61, 0x4F };
        public readonly static byte[] TenKeyNum2 = { 0x62, 0x50 };
        public readonly static byte[] TenKeyNum3 = { 0x63, 0x51 };
        public readonly static byte[] TenKeyNum4 = { 0x64, 0x4B };
        public readonly static byte[] TenKeyNum5 = { 0x65, 0x4C };
        public readonly static byte[] TenKeyNum6 = { 0x66, 0x4D };
        public readonly static byte[] TenKeyNum7 = { 0x67, 0x47 };
        public readonly static byte[] TenKeyNum8 = { 0x68, 0x48 };
        public readonly static byte[] TenKeyNum9 = { 0x69, 0x49 };
        public readonly static byte[] TenKeyMulply = { 0x6A, 0x37 };       // *
        public readonly static byte[] TenKeyAdd = { 0x6B, 0x4E };          // +
        public readonly static byte[] TenKeySubtract = { 0x6D, 0x4A };     // -
        public readonly static byte[] TenKeySeparator = { 0x6C, 0x35 };    // /
        
 
        // Others
        public readonly static byte[] Astarsk = { 0xBA, 0x28 };
        public readonly static byte[] AtMark = { 0xC0, 0x1A };
        public readonly static byte[] BackSlash = { 0xE2, 0x73 };
        public readonly static byte[] BackSpace = { 0x08, 0x0E };
        public readonly static byte[] BracketsL = { 0xDB, 0x1B };
        public readonly static byte[] BracketsR = { 0xDD, 0x2B };
        public readonly static byte[] CapsLock = { 0x00, 0x3A };
        public readonly static byte[] Caret = { 0xDE, 0x0D };
        public readonly static byte[] Colon = { 0xBA, 0x28 };
        public readonly static byte[] Delete = { 0x2E, 0x53 };
        public readonly static byte[] Down = { 0x28, 0x50 };
        public readonly static byte[] End = { 0x23, 0x4F };
        public readonly static byte[] Enter = { 0x0D, 0x1C };
        public readonly static byte[] Equal = { 0x00, 0x00 };
        public readonly static byte[] Escape = { 0x1B, 0x01 };
        public readonly static byte[] GreaterThan = { 0xBE, 0x34 };      // >
        public readonly static byte[] Home = { 0x24, 0x47 };
        public readonly static byte[] Insert = { 0x26, 0x48 };
        public readonly static byte[] Left = { 0x25, 0x70 };
        public readonly static byte[] Kana = { 0xF2, 0x70 };
        public readonly static byte[] Kanji = { 0xF3, 0x4B };            // 半角/英数・漢字
        public readonly static byte[] LessThan = { 0xBC, 0x33 };         // <
        public readonly static byte[] Minus = { 0xBD, 0x0C };
        public readonly static byte[] PageDown = { 0x22, 0x51 };
        public readonly static byte[] PageUp = { 0x21, 0x21 };
        public readonly static byte[] Pause = { 0x13, 0x45 };
        public readonly static byte[] PrintScreen = { 0x2C, 0x37 };
        public readonly static byte[] Right = { 0x27, 0x4D };
        public readonly static byte[] ScrollLock = { 0x91, 0x46 };
        public readonly static byte[] SemiColon = { 0xBB, 0x27 };
        public readonly static byte[] SingleQuote = { 0x00, 0x00 };
        public readonly static byte[] Slash = { 0xBF, 0x35 };
        public readonly static byte[] Tab = { 0x09, 0x0F };
        public readonly static byte[] Underscore = { 0xE2, 0x73 };       // \が２つあるのでShiftとの組み合わせの名称
        public readonly static byte[] Up = { 0x26, 0x48 };
        public readonly static byte[] Yen = { 0xDC, 0x7D };
        public readonly static byte[] WinR = { 0x5C, 0x5C };
    }

    internal class ScanCode {
        // modifiy
        public static byte Muhenkan { get { return KeySetPair.ScanCode(KeySetPair.Muhenkan); } }
        public static byte Henkan { get { return KeySetPair.ScanCode(KeySetPair.Henkan); } }
        public static byte AltL { get { return KeySetPair.ScanCode(KeySetPair.AltL); } }

        // Function
        public static byte F1 { get { return KeySetPair.ScanCode(KeySetPair.F1); } }
        public static byte F2 { get { return KeySetPair.ScanCode(KeySetPair.F2); } }
        public static byte F3 { get { return KeySetPair.ScanCode(KeySetPair.F3); } }
        public static byte F4 { get { return KeySetPair.ScanCode(KeySetPair.F4); } }
        public static byte F5 { get { return KeySetPair.ScanCode(KeySetPair.F5); } }
        public static byte F6 { get { return KeySetPair.ScanCode(KeySetPair.F6); } }
        public static byte F7 { get { return KeySetPair.ScanCode(KeySetPair.F7); } }
        public static byte F8 { get { return KeySetPair.ScanCode(KeySetPair.F8); } }
        public static byte F9 { get { return KeySetPair.ScanCode(KeySetPair.F9); } }
        public static byte F10 { get { return KeySetPair.ScanCode(KeySetPair.F10); } }
        public static byte F11 { get { return KeySetPair.ScanCode(KeySetPair.F11); } }
        public static byte F12 { get { return KeySetPair.ScanCode(KeySetPair.F12); } }
        public static byte F13 { get { return KeySetPair.ScanCode(KeySetPair.F13); } }
        public static byte F14 { get { return KeySetPair.ScanCode(KeySetPair.F14); } }
        public static byte F15 { get { return KeySetPair.ScanCode(KeySetPair.F15); } }


        // Number
        public static byte Num0 { get { return KeySetPair.ScanCode(KeySetPair.Num0); } }
        public static byte Num1 { get { return KeySetPair.ScanCode(KeySetPair.Num1); } }
        public static byte Num2 { get { return KeySetPair.ScanCode(KeySetPair.Num2); } }
        public static byte Num3 { get { return KeySetPair.ScanCode(KeySetPair.Num3); } }
        public static byte Num4 { get { return KeySetPair.ScanCode(KeySetPair.Num4); } }
        public static byte Num5 { get { return KeySetPair.ScanCode(KeySetPair.Num5); } }
        public static byte Num6 { get { return KeySetPair.ScanCode(KeySetPair.Num6); } }
        public static byte Num7 { get { return KeySetPair.ScanCode(KeySetPair.Num7); } }
        public static byte Num8 { get { return KeySetPair.ScanCode(KeySetPair.Num8); } }
        public static byte Num9 { get { return KeySetPair.ScanCode(KeySetPair.Num9); } }

        // Alpha
        public static byte A { get { return KeySetPair.ScanCode(KeySetPair.A); } }
        public static byte C { get { return KeySetPair.ScanCode(KeySetPair.C); } }
        public static byte D { get { return KeySetPair.ScanCode(KeySetPair.D); } }
        public static byte E { get { return KeySetPair.ScanCode(KeySetPair.E); } }
        public static byte F { get { return KeySetPair.ScanCode(KeySetPair.F); } }
        public static byte G { get { return KeySetPair.ScanCode(KeySetPair.G); } }
        public static byte H { get { return KeySetPair.ScanCode(KeySetPair.H); } }
        public static byte I { get { return KeySetPair.ScanCode(KeySetPair.I); } }
        public static byte J { get { return KeySetPair.ScanCode(KeySetPair.J); } }
        public static byte K { get { return KeySetPair.ScanCode(KeySetPair.K); } }
        public static byte L { get { return KeySetPair.ScanCode(KeySetPair.L); } }
        public static byte M { get { return KeySetPair.ScanCode(KeySetPair.M); } }
        public static byte N { get { return KeySetPair.ScanCode(KeySetPair.N); } }
        public static byte O { get { return KeySetPair.ScanCode(KeySetPair.O); } }
        public static byte P { get { return KeySetPair.ScanCode(KeySetPair.P); } }
        public static byte Q { get { return KeySetPair.ScanCode(KeySetPair.Q); } }
        public static byte R { get { return KeySetPair.ScanCode(KeySetPair.R); } }
        public static byte S { get { return KeySetPair.ScanCode(KeySetPair.S); } }
        public static byte T { get { return KeySetPair.ScanCode(KeySetPair.T); } }
        public static byte U { get { return KeySetPair.ScanCode(KeySetPair.U); } }
        public static byte V { get { return KeySetPair.ScanCode(KeySetPair.V); } }
        public static byte W { get { return KeySetPair.ScanCode(KeySetPair.W); } }
        public static byte X { get { return KeySetPair.ScanCode(KeySetPair.X); } }
        public static byte Y { get { return KeySetPair.ScanCode(KeySetPair.Y); } }
        public static byte Z { get { return KeySetPair.ScanCode(KeySetPair.Z); } }


        // Others
        public static byte AtMark { get { return KeySetPair.ScanCode(KeySetPair.AtMark); } }
        public static byte BackSpace { get { return KeySetPair.ScanCode(KeySetPair.BackSpace); } }
        public static byte BracketsR { get { return KeySetPair.ScanCode(KeySetPair.BracketsR); } }
        public static byte CapsLock { get { return KeySetPair.ScanCode(KeySetPair.CapsLock); } }
        public static byte Caret { get { return KeySetPair.ScanCode(KeySetPair.Caret); } }
        public static byte Colon { get { return KeySetPair.ScanCode(KeySetPair.Colon); } }
        public static byte Delete { get { return KeySetPair.ScanCode(KeySetPair.Delete); } }
        public static byte Down { get { return KeySetPair.ScanCode(KeySetPair.Down); } }
        public static byte End { get { return KeySetPair.ScanCode(KeySetPair.End); } }
        public static byte Escape { get { return KeySetPair.ScanCode(KeySetPair.Escape); } }
        public static byte Enter { get { return KeySetPair.ScanCode(KeySetPair.Enter); } }
        public static byte GreaterThan { get { return KeySetPair.ScanCode(KeySetPair.GreaterThan); } }
        public static byte Kana { get { return KeySetPair.ScanCode(KeySetPair.Kana); } }
        public static byte Home { get { return KeySetPair.ScanCode(KeySetPair.Home); } }
        public static byte Insert { get { return KeySetPair.ScanCode(KeySetPair.Insert); } }
        public static byte Left { get { return KeySetPair.ScanCode(KeySetPair.Left); } }
        public static byte LessThan { get { return KeySetPair.ScanCode(KeySetPair.LessThan); } }
        public static byte Minus { get { return KeySetPair.ScanCode(KeySetPair.Minus); } }
        public static byte PageDown { get { return KeySetPair.ScanCode(KeySetPair.PageDown); } }
        public static byte PageUp { get { return KeySetPair.ScanCode(KeySetPair.PageUp); } }
        public static byte Pause { get { return KeySetPair.ScanCode(KeySetPair.Pause); } }
        public static byte PrintScreen { get { return KeySetPair.ScanCode(KeySetPair.PrintScreen); } }
        public static byte Right { get { return KeySetPair.ScanCode(KeySetPair.Right); } }
        public static byte ScrollLock { get { return KeySetPair.ScanCode(KeySetPair.ScrollLock); } }
        public static byte SemiColon { get { return KeySetPair.ScanCode(KeySetPair.SemiColon); } }
        public static byte Slash { get { return KeySetPair.ScanCode(KeySetPair.Slash); } }
        public static byte Tab { get { return KeySetPair.ScanCode(KeySetPair.Tab); } }
        public static byte Underscore { get { return KeySetPair.ScanCode(KeySetPair.Underscore); } }
        public static byte Up { get { return KeySetPair.ScanCode(KeySetPair.Up); } }
        public static byte Yen { get { return KeySetPair.ScanCode(KeySetPair.Yen); } }
        public static byte WinR { get { return KeySetPair.ScanCode(KeySetPair.WinR); } }
    }
}
