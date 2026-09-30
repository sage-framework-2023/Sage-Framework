using System;
using System.Runtime.InteropServices;
using System.Text;

namespace SageVBE
{
    static class Native
    {
        public const int WM_NCDESTROY = 0x0082;
        public const int WM_ERASEBKGND = 0x0014;
        public const int WM_SYSCOLORCHANGE = 0x0015;
        public const int WM_THEMECHANGED = 0x031A;
        public const int WM_CTLCOLOREDIT = 0x0133;
        public const int WM_CTLCOLORLISTBOX = 0x0134;
        public const int WM_CTLCOLORSTATIC = 0x0138;

        public const int WH_CBT = 5;
        public const int HCBT_CREATEWND = 3;

        public const uint GA_ROOT = 2;
        public const uint GA_ROOTOWNER = 3;

        public const uint RDW_INVALIDATE = 0x1, RDW_ERASE = 0x4, RDW_ALLCHILDREN = 0x80, RDW_UPDATENOW = 0x100, RDW_FRAME = 0x400;

        public const uint PAGE_READWRITE = 0x04;

        public const int TVM_SETBKCOLOR = 0x1100 + 29;
        public const int TVM_SETTEXTCOLOR = 0x1100 + 30;
        public const int TVM_SETLINECOLOR = 0x1100 + 40;

        public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        public const int DWMWA_BORDER_COLOR = 34;
        public const int DWMWA_CAPTION_COLOR = 35;
        public const int DWMWA_TEXT_COLOR = 36;
        public const int DWMWA_COLOR_DEFAULT = unchecked((int)0xFFFFFFFF);

        public const int WM_PAINT = 0x000F;
        public const int WM_NCPAINT = 0x0085;
        public const int WM_GETFONT = 0x0031;
        public const int GWL_STYLE = -16, GWL_EXSTYLE = -20;
        public const int WS_CHILD = 0x40000000;
        public const uint SRCCOPY = 0x00CC0020;

        public const int TCM_GETITEMCOUNT = 0x1304, TCM_GETITEMRECT = 0x130A, TCM_GETCURSEL = 0x130B, TCM_GETITEMW = 0x133C;
        public const uint TCIF_TEXT = 1;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct PAINTSTRUCT
        {
            public IntPtr hdc;
            public int fErase;
            public RECT rcPaint;
            public int fRestore, fIncUpdate;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] public byte[] rgbReserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TCITEM
        {
            public uint mask, dwState, dwStateMask;
            public IntPtr pszText;
            public int cchTextMax, iImage;
            public IntPtr lParam;
        }

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hwnd, int index);
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromDC(IntPtr hdc);
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hwnd);
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);
        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
        [DllImport("user32.dll")]
        public static extern int FrameRect(IntPtr dc, ref RECT rect, IntPtr brush);
        [DllImport("user32.dll")]
        public static extern IntPtr BeginPaint(IntPtr hwnd, out PAINTSTRUCT ps);
        [DllImport("user32.dll")]
        public static extern bool EndPaint(IntPtr hwnd, ref PAINTSTRUCT ps);
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, out RECT lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, ref TCITEM lParam);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateRectRgn(int left, int top, int right, int bottom);
        [DllImport("gdi32.dll")]
        public static extern bool FillRgn(IntPtr hdc, IntPtr region, IntPtr brush);
        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr obj);
        [DllImport("user32.dll")]
        public static extern int GetUpdateRgn(IntPtr hwnd, IntPtr region, bool erase);
        [DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr dest, int x, int y, int w, int h, IntPtr src, int sx, int sy, uint rop);

        public delegate IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData);
        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate uint GetSysColorProc(int index);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate IntPtr GetSysColorBrushProc(int index);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate int SetColorProc(IntPtr hdc, int color);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate IntPtr CreateSolidBrushProc(int color);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate IntPtr GetStockObjectProc(int index);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate int FillRectProc(IntPtr hdc, IntPtr rect, IntPtr brush);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        public delegate IntPtr CreatePenProc(int style, int width, int color);

        [DllImport("comctl32.dll")]
        public static extern bool SetWindowSubclass(IntPtr hwnd, SubclassProc proc, UIntPtr id, UIntPtr refData);
        [DllImport("comctl32.dll")]
        public static extern bool RemoveWindowSubclass(IntPtr hwnd, SubclassProc proc, UIntPtr id);
        [DllImport("comctl32.dll")]
        public static extern IntPtr DefSubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc proc, IntPtr hMod, uint threadId);
        [DllImport("user32.dll")]
        public static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        public static extern IntPtr GetProcAddress(IntPtr module, string name);
        [DllImport("kernel32.dll")]
        public static extern bool VirtualProtect(IntPtr address, UIntPtr size, uint newProtect, out uint oldProtect);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc proc, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern bool EnumThreadWindows(uint threadId, EnumWindowsProc proc, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hwnd);
        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
        [DllImport("user32.dll")]
        public static extern int FillRect(IntPtr hdc, ref RECT rect, IntPtr brush);
        [DllImport("user32.dll")]
        public static extern bool RedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateSolidBrush(int color);
        [DllImport("gdi32.dll")]
        public static extern int SetTextColor(IntPtr hdc, int color);
        [DllImport("gdi32.dll")]
        public static extern int SetBkColor(IntPtr hdc, int color);

        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        public static extern int SetWindowTheme(IntPtr hwnd, string subAppName, string subIdList);

        public static string ClassName(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(64);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }
}
