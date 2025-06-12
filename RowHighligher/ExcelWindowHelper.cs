using System;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace RowHighligher
{
    public static class ExcelWindowHelper
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // GetWindow commands
        private const uint GW_OWNER = 4;
        private const uint GW_HWNDFIRST = 0;
        private const uint GW_HWNDNEXT = 2;

        public static bool IsExcelDialog(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
                return false;

            try
            {
                // Check for standard dialog class directly
                StringBuilder className = new StringBuilder(256);
                GetClassName(windowHandle, className, className.Capacity);
                string currentClassName = className.ToString();

                // Check for common dialog classes that Excel uses
                bool isStandardDialog = currentClassName == "#32770";
                bool isExcelDialog = currentClassName.Contains("Excel") && currentClassName.Contains("Dialog");
                bool isMessageBox = currentClassName == "MessageBoxEx";
                bool isTaskDialog = currentClassName == "TaskDialog";

                if (isStandardDialog || isExcelDialog || isMessageBox || isTaskDialog)
                {
                    // Verify it belongs to Excel by checking the ownership chain
                    IntPtr ownerWindow = GetWindow(windowHandle, GW_OWNER);
                    if (ownerWindow != IntPtr.Zero)
                    {
                        className.Clear();
                        GetClassName(ownerWindow, className, className.Capacity);
                        string ownerClass = className.ToString();
                        if (ownerClass.StartsWith("EXCEL") || ownerClass.Contains("Excel"))
                        {
                            return true;
                        }
                    }
                }

                // Additional checks for Excel modeless dialogs
                IntPtr parent = GetParent(windowHandle);
                if (parent != IntPtr.Zero)
                {
                    className.Clear();
                    GetClassName(parent, className, className.Capacity);
                    string parentClass = className.ToString();
                    if (parentClass.StartsWith("EXCEL") || parentClass.Contains("Excel"))
                    {
                        return currentClassName.Contains("Excel") || 
                               isStandardDialog || 
                               isMessageBox || 
                               isTaskDialog;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in IsExcelDialog: {ex.Message}");
            }

            return false;
        }

        public static bool IsAnyExcelDialogActive()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                if (foregroundWindow != IntPtr.Zero && IsExcelDialog(foregroundWindow))
                    return true;

                // Check all top-level windows
                bool foundDialog = false;
                EnumWindows((hWnd, lParam) =>
                {
                    if (IsExcelDialog(hWnd))
                    {
                        foundDialog = true;
                        return false; // Stop enumeration
                    }
                    return true; // Continue enumeration
                }, IntPtr.Zero);

                return foundDialog;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in IsAnyExcelDialogActive: {ex.Message}");
                return false;
            }
        }

        public static void UpdateFormTopMost(System.Windows.Forms.Form form, bool isDetached)
        {
            if (form == null || form.IsDisposed)
            {
                return;
            }

            try
            {
                // If any Excel dialog is active, always set TopMost = false
                if (IsAnyExcelDialogActive())
                {
                    if (form.InvokeRequired)
                    {
                        form.BeginInvoke((Action)(() => form.TopMost = false));
                    }
                    else
                    {
                        form.TopMost = false;
                    }
                    return;
                }

                // Otherwise, set TopMost according to the detached setting
                if (isDetached)
                {
                    if (form.InvokeRequired)
                    {
                        form.BeginInvoke((Action)(() => form.TopMost = true));
                    }
                    else
                    {
                        form.TopMost = true;
                    }
                }
                else
                {
                    if (form.InvokeRequired)
                    {
                        form.BeginInvoke((Action)(() => form.TopMost = false));
                    }
                    else
                    {
                        form.TopMost = false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateFormTopMost: {ex.Message}");
            }
        }
    }
}