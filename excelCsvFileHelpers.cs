using System;
using System.Xml;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using TridentGoalSeek;
using ExcelDataReader;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;


public static DataTable ReadExcelSheet(string path, string sheetname)
        {
            DataTable dt = new DataTable();
            try
            {
                using (OleDbConnection conn = new OleDbConnection())
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        comm.CommandText = "Select * from [" + sheetname + "]";
                        comm.Connection = conn;

                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Instance.Log("An error occurred while reading excel sheet : " + e.Message, Logger.LogType.Fatal);
            }
            return dt;
        }
        
        
        
        public static void WriteExcelSheet(DataTable dt, List<string> underlierList, string path, string sheetName)
        {
            var excelApp = new Excel.Application();

            try
            {
                if (dt == null || dt.Columns.Count == 0)
                {
                    Logger.Instance.Log("No table to export to Excel");
                }

                // load excel, and create a new workbook
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;
                workSheet.Name = sheetName;

                // column headings
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }

                excelApp.Workbooks[1].SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, ConflictResolution: Excel.XlSaveConflictResolution.xlOtherSessionChanges);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception e)
            {
                excelApp.ActiveWorkbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                Logger.Instance.Log("Error while creating Excel file : " + e.Message, Logger.LogType.Fatal);
            }

            try
            {

                List<string> queries = new List<string>();
                string query = "";
                string columnList = "";


                columnList += "(";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    columnList += "[" + dt.Columns[j].ColumnName + "]";
                    if (j < dt.Columns.Count - 1) { columnList += ","; }
                }
                columnList += ")";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var underlying = dt.Rows[i][7].ToString();
                    if (underlierList.Contains(underlying))
                    {
                        query = "INSERT INTO [" + sheetName + "$] " + columnList + " VALUES ";
                        query += "(";
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (j==1 || j==9)
                                query += "'" + DateTime.Parse(dt.Rows[i][j].ToString()).ToString("yyyy-MM-dd") + "'";
                            else
                                query += "'" + dt.Rows[i][j].ToString() + "'";
                            if (j < dt.Columns.Count - 1) { query += ","; }
                        }
                        query += ");" + System.Environment.NewLine;
                        queries.Add(query);
                    }
                }

                WriteExcelSheet(queries.ToArray(), path);
            }
            catch (Exception e)
            {
                Logger.Instance.Log("Error while writing data to template file : " + e.Message, Logger.LogType.Fatal);
            }
        }

        public static void ExportToExcel(DataTable dt, List<string> underlierList, string excelFilePath)
        {
            try
            {
                if (dt == null || dt.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table");

                // load excel, and create a new workbook
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column headings
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }

                var k = 0;
                // rows
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    var underlying = dt.Rows[i][7].ToString();

                    if (underlierList.Contains(underlying))
                    {
                        k++;
                        // to do: format datetime values before printing
                        for (var j = 0; j < dt.Columns.Count; j++)
                        {
                            workSheet.Cells[k + 1, j + 1] = dt.Rows[i][j];
                        }
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        excelApp.Workbooks[1].SaveAs(excelFilePath, Excel.XlFileFormat.xlWorkbookNormal, ConflictResolution: Excel.XlSaveConflictResolution.xlOtherSessionChanges);
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: " + ex.Message);
            }
        }

        public static void WriteExcelSheet(string[] oleDBQueries, string path)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection())
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;

                    int rowsAffected = 0;
                    conn.Open();

                    for (int i = 0; i < oleDBQueries.Length; i++)
                    {
                        cmd.CommandText = oleDBQueries[i];
                        rowsAffected = cmd.ExecuteNonQuery();
                    }

                    conn.Close();
                }
            }
            catch (Exception e)
            {
                Logger.Instance.Log("Error while writing template data into Excel file : " + e.Message, Logger.LogType.Fatal);
            }
        }
        public static void WriteExcelSheetExEmptyCommand(string[] oleDBQueries, string path)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection())
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;

                    int rowsAffected = 0;
                    conn.Open();

                    for (int i = 0; i < oleDBQueries.Length; i++)
                    {
                        cmd.CommandText = oleDBQueries[i];
                        if(cmd.CommandText.Length > 0)
                        {
                            rowsAffected = cmd.ExecuteNonQuery();
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception e)
            {
                Logger.Instance.Log("Error while writing template data into Excel file : " + e.Message, Logger.LogType.Fatal);
            }
        }
        public static void WriteExcelSheet(DataTable dt, string path, string sheetName)
        {
            var excelApp = new Excel.Application();

            try
            {
                if (dt == null || dt.Columns.Count == 0)
                {
                    Logger.Instance.Log("No table to export to Excel");
                }

                // load excel, and create a new workbook
                excelApp.Workbooks.Add(); 

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;
                workSheet.Name = sheetName;

                // column headings
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }

                excelApp.Workbooks[1].SaveAs(path, Excel.XlFileFormat.xlExcel8, ConflictResolution: Excel.XlSaveConflictResolution.xlOtherSessionChanges);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception e)
            {
                excelApp.ActiveWorkbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                Logger.Instance.Log("Error while creating Excel file : " + e.Message, Logger.LogType.Fatal);
            }

            try
            {

                List<string> queries = new List<string>();
                string query = "";
                string columnList = "";


                columnList += "(";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    columnList += "[" + dt.Columns[j].ColumnName + "]";
                    if (j < dt.Columns.Count - 1) { columnList += ","; }
                }
                columnList += ")";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    query = "INSERT INTO [" + sheetName + "$] " + columnList + " VALUES ";
                    query += "(";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        query += "'" + dt.Rows[i][j].ToString() + "'";
                        if (j < dt.Columns.Count - 1) { query += ","; }
                    }
                    query += ");" + System.Environment.NewLine;
                    queries.Add(query);
                }

                WriteExcelSheet(queries.ToArray(), path);
            }
            catch (Exception e)
            {
                Logger.Instance.Log("Error while writing data to template file : " + e.Message, Logger.LogType.Fatal);
            }
        }

        public static bool ReadExcelFile(string path, string extension)
        {
            try
            {
                IExcelDataReader reader = null;
                using (Stream s = new FileStream(path, FileMode.Open))
                {
                    if (extension == ".xls")
                    {
                        // Reading from a binary Excel file ('97-2003 format; *.xls)
                        reader = ExcelReaderFactory.CreateBinaryReader(s);
                    }
                    else
                    {
                        // Reading from an OpenXml Excel file (2007 format; *.xlsx)
                        reader = ExcelReaderFactory.CreateOpenXmlReader(s);
                    }

                    //if (!reader.IsValid) { throw new Exception(reader.ExceptionMessage); }


                }
            }
            catch
            {
                return false;
            }

            return true;

        }
        
        
        
        
        #region CSV

        public static DataTable CsvToDataTable(string fileName, char separator)
        {
            return FileHelpers.CommonEngine.CsvToDataTable(fileName, separator);
        }

        public static void DataTableToCsv(DataTable table, string fileName)
        {
            FileHelpers.CommonEngine.DataTableToCsv(table, fileName);
        }

        public static void ListToCsv(List<string> list, string columns, string fileName)
        {
            var table = new DataTable();

            var dataRow = table.NewRow();

            var col = columns.Split(';');

            for (var j = 0; j < col.Length; j++)
            {
                var newColumn = new DataColumn {DataType = col[j].GetType(), ColumnName = col[j]};
                table.Columns.Add(newColumn);

                dataRow[j] = table.Columns[j].ColumnName;
            }

            table.Rows.InsertAt(dataRow, 0);
            
            foreach (var line in list)
            {
                var i = 0;
                var row = table.NewRow();
                var values = line.Split(';');
                foreach (var val in values)
                {
                    row[table.Columns[i].ColumnName] = val;
                    ++i;
                }
                table.Rows.Add(row);
            }

            FileHelpers.CommonEngine.DataTableToCsv(table, fileName);
        }

        public static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            var properties = TypeDescriptor.GetProperties(typeof(T));
            var table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (var item in data)
            {
                var row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        #endregion
        
        #region Date Helpers

        public static double GetDoubleFromDate(DateTime date)
        {
            return (date - new DateTime(1899, 12, 30)).Days;
        }
        public static DateTime GetDateTimeFromDouble(double date)
        {
            return new DateTime(1899, 12, 30).AddDays(date);
        }
        public static DateTime AddWorkingDays(DateTime date, int d)
        {
            DateTime dt_temp = date;
            for (int i = 0; i < Math.Abs(d); i++)
            {
                dt_temp = dt_temp.AddDays(Math.Sign(d));
                while (dt_temp.DayOfWeek == DayOfWeek.Saturday || dt_temp.DayOfWeek == DayOfWeek.Sunday)
                {
                    dt_temp = dt_temp.AddDays(Math.Sign(d));
                }
            }

            return dt_temp;
        }
        public static int NbWorkingDays(DateTime startDate, DateTime endDate)
        {
            //Includes both StartDate and EndDate
            int nbDays = 0;
            DateTime date = startDate;

            while (date <= endDate)
            {
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday) { nbDays += 1; }
                date = date.AddDays(1);
            }
            return nbDays;
        }
        public static DateTime GetEOMDay(int year, int month)
        {
            DateTime res = new DateTime(year, month, 1).AddMonths(1).AddDays(-1);
            if (res.DayOfWeek == DayOfWeek.Saturday || res.DayOfWeek == DayOfWeek.Sunday)
            {
                res = AddWorkingDays(res, -1);
            }
            return res;
        }
        public static DateTime GetPreviousEOMDay(DateTime date)
        {
            DateTime dt_temp = new DateTime(date.Year, date.Month, 1).AddDays(-1);
            return GetEOMDay(dt_temp.Year, dt_temp.Month);
        }
        public static bool isEOMDay(DateTime date)
        {
            return (date.Month != AddWorkingDays(date, 1).Month);
        }
        public static DateTime GetUTCPlusOneDateAndTimeForZone(DateTime date, TimeZone zone)
        {
            DateTime UTCPlusOnedate = date;
            switch (zone)
            {
                case TimeZone.TK:
                    if (date < GetMonthLastSunday(3, date.Year) || date > GetMonthLastSunday(10, date.Year))
                        UTCPlusOnedate = date.AddHours(-8);
                    else
                        UTCPlusOnedate = date.AddHours(-7);
                    break;
                case TimeZone.HK:
                    if (date < GetMonthLastSunday(3, date.Year) || date > GetMonthLastSunday(10, date.Year))
                        UTCPlusOnedate = date.AddHours(-7);
                    else
                        UTCPlusOnedate = date.AddHours(-6);
                    break;
                case TimeZone.UK:
                    UTCPlusOnedate = date.AddHours(1);
                    break;
                case TimeZone.NY:
                    if (date < GetMonthLastSunday(10, date.Year).AddDays(7) && date > GetMonthLastSunday(10, date.Year))
                        UTCPlusOnedate = date.AddHours(5);
                    else
                        UTCPlusOnedate = date.AddHours(6);
                    break;
                default:
                    break;
            }

            return UTCPlusOnedate;
            
        }
        public static DateTime GetMonthLastSunday(int month, int year)
        {
            DateTime d = DateTime.Parse("31/" + month + "/" + year);
            while (d.DayOfWeek != DayOfWeek.Sunday)
                d = d.AddDays(-1);

            return d;
        }

        #endregion
        
        
        #region Files

        public static void CopyAndArchive(string origin, string destination)
        {
            if (origin.ToUpper() == destination.ToUpper()) { return; }

            int lastPointPosition = destination.LastIndexOf(".");
            string extension = destination.Substring(lastPointPosition);
            string formatArchiveFile = "yyyy_MM_dd_HH_mm_ss";
            string sNow = DateTime.Now.ToString(formatArchiveFile);

            //Archive
            if (File.Exists(destination))
            {
                File.Move(destination, destination.Replace(extension, "_" + sNow + extension));
            }

            //Copy
            File.Copy(origin, destination);
        }

        #endregion
