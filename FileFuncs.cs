using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace VariousFileFunctions
{
    class FileFunctions
    {
        // Generic Files Functions
        #region Generic Files Functions
        // ------------------------------------------------------------

        public static bool File_VerifyExistence(string filename)
        {
            if (File.Exists(filename))
                return true;
            else
                return false;
        }

        public static bool File_CutAndPaste(string filename, string destinyfile)
        {
            if (File.Exists(filename))
            {
                File.Copy(filename, destinyfile);
                if (File.Exists(destinyfile))
                {
                    try
                    {
                        File.Delete(filename);
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
                        
        }

        public static bool File_CopyAndPaste(string filename, string destinyfile)
        {
            if (File.Exists(filename))
            {
                File.Copy(filename, destinyfile);
                if (File.Exists(destinyfile))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool File_Delete(string filename)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
                return true;
            }
            else
            {
                return false;
            }
        }

        // ------------------------------------------------------------
        #endregion

        // Excel Files Functions
        #region Excel File Functions
        // ------------------------------------------------------------

        public static Workbook File_OpenExcelFile(string filename)
        {
            if (File.Exists(filename))
            {
                var processes = Process.GetProcessesByName("EXCEL");
                foreach (var p in processes)
                    p.Kill();

                _Application excel = new _Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;
                Workbook b = excel.Workbooks.Open(filename);
                return b;
            }
            else
            {
                return null;
            }

        }

        public static bool File_ExcelToCSV(string filename, string destinyfile)
        {
            var processes = Process.GetProcessesByName("EXCEL");
            foreach (var p in processes)
                p.Kill();

            if (File.Exists(destinyfile))
            {
                File.Delete(destinyfile); return true;
            }
            if (File.Exists(filename))
            {
                try
                {
                    _Application excel = new _Excel.Application();
                    excel.Visible = false;
                    excel.DisplayAlerts = false;
                    excel.ScreenUpdating = false;
                    Workbook b = excel.Workbooks.Open(filename);
                    b.SaveAs(destinyfile, XlFileFormat.xlCSVWindows, Local: true);
                    b.Close();
                    var processes1 = Process.GetProcessesByName("EXCEL");
                    foreach (var p in processes1)
                        p.Kill();
                    return true;
                }
                catch
                {
                    var processes1 = Process.GetProcessesByName("EXCEL");
                    foreach (var p in processes1)
                        p.Kill();
                    return false;
                }
            }
            else
            {
                return false;
            }           
        }

        public static bool File_TxtToCSV(string filename, string destinyfile)
        {
            string[] lines = System.IO.File.ReadAllLines(filename);
            StringBuilder builder = new StringBuilder();

            foreach (string line in lines)
            {
                var temp = line.Split('\r');
                builder.AppendLine(string.Join(";", temp[0], temp[1]));
                //builder.AppendLine(string.Format("{0}; {1}", temp[0], temp[1]));
            }
            File.WriteAllText(destinyfile, builder.ToString());
            return true;
        }

        // ------------------------------------------------------------
        #endregion

        // Directory Functions
        #region Directory Functions
        // ------------------------------------------------------------

        public static bool Dir_ClearDirectory(string directory)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(directory);
                foreach (FileInfo file in dir.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo di in dir.GetDirectories())
                {
                    dir.Delete(true);
                }
                return true;
            }
            catch
            {
                return false;
            }           
        }

        public static bool Dir_DeleteDirectory(string directory)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(directory);
                dir.Delete(true);
                return true;
            }
            catch
            {
                return false;
            }
            
        }

        // ------------------------------------------------------------
        #endregion
    }
}
