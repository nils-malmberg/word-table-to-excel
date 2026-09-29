using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WordTableToExcel.App
{
    internal static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>OBJID_NATIVEOM : modèle objet natif de l'application propriétaire de la fenêtre (Word.Window).</summary>
        public const uint ObjIdNativeOm = 0xFFFFFFF0;
        public const int SwRestore = 9;

        public static readonly Guid IidIDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int command);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        public static string ClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(64);
            return GetClassName(hWnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
        }

        /// <summary>Met une fenêtre au premier plan (et la restaure si elle est réduite).</summary>
        public static void BringToFront(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            try
            {
                if (IsIconic(hWnd)) ShowWindow(hWnd, SwRestore);
                SetForegroundWindow(hWnd);
            }
            catch (Exception)
            {
                // Sans conséquence.
            }
        }
    }
}
