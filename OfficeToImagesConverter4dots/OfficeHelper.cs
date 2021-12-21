using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeToImagesConverter4dots
{
    class OfficeHelper
    {
        public static object WordApp = null;
        public static Type WordApplicationType;
        public static object ExcelApp = null;
        public static object PPApp = null;
        public static Type ExcelApplicationType;
        public static Type PPApplicationType;

        public static void CreateWordApplication()
        {
            if (WordApp != null)
            {
                try
                {
                    WordApp.GetType().InvokeMember("Visible", BindingFlags.IgnoreReturn | BindingFlags.Public |
                    BindingFlags.Static | BindingFlags.SetProperty, null, WordApp, new object[] { false });
                }
                catch
                {
                    WordApplicationType = System.Type.GetTypeFromProgID("Word.Application");
                    WordApp = Activator.CreateInstance(WordApplicationType);
                }
            }
            else
            {
                WordApplicationType = System.Type.GetTypeFromProgID("Word.Application");
                WordApp = Activator.CreateInstance(WordApplicationType);
            }
        }

        public static void CreatePowerPointApplication()
        {
            if (PPApp != null)
            {
                try
                {
                    PPApp.GetType().InvokeMember("Visible", BindingFlags.IgnoreReturn | BindingFlags.Public |
                    BindingFlags.Static | BindingFlags.SetProperty, null, PPApp, new object[] { false });
                }
                catch
                {
                    PPApplicationType = System.Type.GetTypeFromProgID("PowerPoint.Application");
                    PPApp = Activator.CreateInstance(PPApplicationType);
                }
            }
            else
            {
                PPApplicationType = System.Type.GetTypeFromProgID("PowerPoint.Application");
                PPApp = Activator.CreateInstance(PPApplicationType);
            }
        }

        public static void QuitPowerPointApplication()
        {
            if (PPApp != null)
            {
                try
                {
                    PPApp.GetType().InvokeMember("Quit", BindingFlags.IgnoreReturn | BindingFlags.Instance |
                    BindingFlags.InvokeMethod, null, PPApp, null);
                }
                catch { }

                PPApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void QuitWordApplication()
        {
            if (WordApp != null)
            {
                try
                {
                    WordApp.GetType().InvokeMember("Quit", BindingFlags.IgnoreReturn | BindingFlags.Instance |
                    BindingFlags.InvokeMethod, null, WordApp, null);
                }
                catch { }

                WordApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void CreateExcelApplication()
        {
            if (ExcelApp != null)
            {
                try
                {
                    ExcelApp.GetType().InvokeMember("Visible", BindingFlags.IgnoreReturn | BindingFlags.Public |
                    BindingFlags.Static | BindingFlags.SetProperty, null, ExcelApp, new object[] { false }, System.Globalization.CultureInfo.InvariantCulture);
                }
                catch
                {
                    ExcelApplicationType = System.Type.GetTypeFromProgID("Excel.Application");
                    ExcelApp = Activator.CreateInstance(ExcelApplicationType);
                }
            }
            else
            {
                ExcelApplicationType = System.Type.GetTypeFromProgID("Excel.Application");
                ExcelApp = Activator.CreateInstance(ExcelApplicationType);
            }
        }
        public static void QuitExcelApplication()
        {
            if (ExcelApp != null)
            {
                try
                {
                    ExcelApp.GetType().InvokeMember("Quit", BindingFlags.IgnoreReturn | BindingFlags.Instance |
                    BindingFlags.InvokeMethod, null, ExcelApp, null, System.Globalization.CultureInfo.InvariantCulture);
                }
                catch { }

                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


        }

        public static void QuitOfficeApplications()
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            proc.StartInfo.FileName = "QuitOfficeApplications.exe";
            proc.StartInfo.CreateNoWindow = true;

            proc.Start();
            proc.WaitForExit();
        }



    }
}
