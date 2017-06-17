using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using RGiesecke.DllExport;

namespace MetatraderToExcel
{
    public class Class1
    {
        static Excel.Application excelApp;
        static Excel.Workbook excelWorkbook;
        static Excel.Worksheet excelWorksheet;
        static object misValue = System.Reflection.Missing.Value;
        static string LogFile = null;

        [DllExport]
        public static void Initialize([MarshalAs(UnmanagedType.LPWStr)]string wkb, [MarshalAs(UnmanagedType.LPWStr)] string whereToLog)
        {
            try
            {
                LogFile = whereToLog;
                excelApp = new Excel.Application();
                excelWorkbook = excelApp.Workbooks.Open(wkb, 3, false, misValue, misValue, misValue, true, misValue, misValue, misValue, true, misValue, false, false, misValue);
                excelApp.Visible = true;
            }
            catch (Exception e)
            {
                WriteMe("Error on Initialize: " + e);
                throw;
            }
            
        }
        #region PutDouble
        [DllExport]
        public static void PutDouble_Cell(double dbl, [MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Range[cell].Value = dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutDouble_Cell: " + e);
                throw;
            }
            
        }

        [DllExport]
        public static void PutDouble_intidx(double dbl, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, columnindex] = dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutDouble_intidx: " + e);
                throw;
            }
            
        }
        
        [DllExport]
        public static void PutDouble_intCell(double dbl, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, column] = dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutDouble_intCell: " + e);
                throw;
            }
        }
        #endregion
        #region PutInt
        [DllExport]
        public static void PutInt_Cell(int integer, [MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Range[cell].Value = integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutInt_Cell: " + e);
                throw;
            }
            
        }

        [DllExport]
        public static void PutInt_intidx(int integer, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, columnindex] = integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutIntx_intidx: " + e);
                throw;
            }
            
        }

        [DllExport]
        public static void PutInt_intCell(int integer, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, column] = integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutInt_intCell: " + e);
                throw;
            }
        }
        #endregion
        #region PutStr
        [DllExport]
        public static void PutStr_Cell([MarshalAs(UnmanagedType.LPWStr)]string Str, [MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Range[cell].Value = Str;

            }
            catch (Exception e)
            {
                WriteMe("Error on PutStr_Cell: " + e);
                throw;
            }
        }    

        [DllExport]
        public static void PutStr_intidx([MarshalAs(UnmanagedType.LPWStr)]string Str, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, columnindex] = Str;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutStr_intidx: " + e);
                throw;
            }
            
        }

        [DllExport]
        public static void PutStr_intCell([MarshalAs(UnmanagedType.LPWStr)]string Str, [MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                excelWorksheet = excelApp.Sheets[shts];
                excelWorksheet.Cells[rowindex, column] = Str;
            }
            catch (Exception e)
            {
                WriteMe("Error on PutStr_intCell: " + e);
                throw;
            }
            
        }

        #endregion
        #region GetDouble
        [DllExport]
        public static double GetDouble_Cell([MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                double dbl = 0.0;
                excelWorksheet = excelApp.Sheets[shts]; 
                dbl = (double)excelWorksheet.Range[cell].Value;
                return dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetDouble_Cell: " + e);
                throw;
            }
        }

        [DllExport]
        public static double GetDouble_intidx([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                double dbl = 0.0;
                excelWorksheet = excelApp.Sheets[shts];
                dbl = (double)(excelWorksheet.Cells[rowindex, columnindex] as Excel.Range).Value;
                return dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetDouble_intidx: " + e);
                throw;
            }
        }

        [DllExport]
        public static double GetDouble_intCell([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                double dbl = 0.0;
                excelWorksheet = excelApp.Sheets[shts];
                dbl = (double)(excelWorksheet.Cells[rowindex, column] as Excel.Range).Value;
                return dbl;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetDouble_intCell: " + e);
                throw;
            }
        }
        #endregion
        #region GetInt
        [DllExport]
        public static int GetInt_Cell([MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                int integer = 0;
                excelWorksheet = excelApp.Sheets[shts];
                integer = (int)excelWorksheet.Range[cell].Value;
                return integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetInt_Cell: " + e);
                throw;
            }
        }

        [DllExport]
        public static int GetInt_intidx([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                int integer = 0;
                excelWorksheet = excelApp.Sheets[shts];
                integer = (int)(excelWorksheet.Cells[rowindex, columnindex] as Excel.Range).Value;
                return integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetInt_intidx: + e");
                throw;
            }
        }

        [DllExport]
        public static int GetInt_intCell([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                int integer = 0;
                excelWorksheet = excelApp.Sheets[shts];
                integer = (int)(excelWorksheet.Cells[rowindex, column] as Excel.Range).Value;
                return integer;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetInt_intCell: " + e);
                throw;
            }
        }
        #endregion
        #region GetStr
        [DllExport]
        [return: MarshalAs(UnmanagedType.LPTStr)]
        public static string GetStr_Cell([MarshalAs(UnmanagedType.LPWStr)]string shts, [MarshalAs(UnmanagedType.LPWStr)]string cell)
        {
            try
            {
                string Str = null;
                excelWorksheet = excelApp.Sheets[shts];
                Str = (string)excelWorksheet.Range[cell].Value;
                return Str;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetStr_Cell: " + e);
                throw;
            }
        }

        [DllExport]
        [return: MarshalAs(UnmanagedType.LPTStr)]
        public static string GetStr_intidx([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, int columnindex)
        {
            try
            {
                string Str = null;
                excelWorksheet = excelApp.Sheets[shts];
                Str = (string)(excelWorksheet.Cells[rowindex, columnindex] as Excel.Range).Value;
                return Str;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetStr_intidx: " + e);
                throw;
            }
        }

        [DllExport]
        [return: MarshalAs(UnmanagedType.LPTStr)]
        public static string GetStr_intCell([MarshalAs(UnmanagedType.LPWStr)]string shts, int rowindex, [MarshalAs(UnmanagedType.LPWStr)]string column)
        {
            try
            {
                string Str = null;
                excelWorksheet = excelApp.Sheets[shts];
                Str = (string)(excelWorksheet.Cells[rowindex, column] as Excel.Range).Value;
                return Str;
            }
            catch (Exception e)
            {
                WriteMe("Error on GetStr_intCell: " + e);
                throw;
            }
        }

        #endregion
        public static void WriteMe(string line)
        {
            using (StreamWriter writer = new StreamWriter(LogFile, true))
            {
                writer.WriteLine(line);
            }
        }
    }
}

/*
        [DllExport]
        [return: MarshalAs(UnmanagedType.LPTStr)]
        public static string GetString()
        {
            return (myData);
        }
*/
