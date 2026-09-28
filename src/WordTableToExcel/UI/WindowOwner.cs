using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordTableToExcel.UI
{
    /// <summary>Fenêtre Word utilisée comme propriétaire des boîtes de dialogue (qui restent ainsi au premier plan).</summary>
    internal sealed class WindowOwner : IWin32Window
    {
        private WindowOwner(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; private set; }

        public static IWin32Window FromWord(object application)
        {
            IntPtr handle = IntPtr.Zero;
            try
            {
                dynamic app = application;
                handle = new IntPtr(Convert.ToInt64(app.ActiveWindow.Hwnd)); // Word 2013+
            }
            catch (Exception)
            {
                handle = IntPtr.Zero;
            }
            if (handle == IntPtr.Zero) handle = GetForegroundWindow();
            if (handle == IntPtr.Zero)
            {
                try
                {
                    handle = Process.GetCurrentProcess().MainWindowHandle;
                }
                catch (Exception)
                {
                    handle = IntPtr.Zero;
                }
            }
            return handle == IntPtr.Zero ? null : new WindowOwner(handle);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }
}
