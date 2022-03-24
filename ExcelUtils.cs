using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace SendMailThue
{
    public class ExcelUtils
    {

        public static List<string> GetValuesFromFile(string file, List<string> cells)
        {
            List<string> fieldValues = new List<string>();
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(file);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;
                var values = (xlRange.Value as Object[,]);

                foreach (string cell in cells)
                {
                    string[] coor = cell.Split("-");
                    string cellValue = "";
                    try
                    {
                        cellValue = Convert.ToString(values[int.Parse(coor[0]), int.Parse(coor[1])]);

                    }
                    catch (Exception ex) { }
                    fieldValues.Add(cellValue);
                }
            }
            catch (Exception ex) { }
            finally
            {
                closeExcel(xlApp, xlWorkbook, xlWorksheet, null);
            }


            return fieldValues;
        }

        public static void GetGroupValuesByGroup(string file, List<string> fields, int spaceRowBetweenGroup, string outputDir, Action<List<List<string>>> callback)
        {
            List<List<string>> groups = new List<List<string>>();

            FileUtils.CreateDir(outputDir, true);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(file);
                xlWorksheet = xlWorkbook.ActiveSheet;
                xlRange = xlWorksheet.UsedRange;
                int rowCount = xlWorksheet.UsedRange.Rows.Count;
                int columnCount = xlWorksheet.UsedRange.Columns.Count;

                var values = (xlRange.Value as Object[,]);
                List<int[]> rangs = GetListOfRangeHasData(values, rowCount, columnCount, 30);

                foreach (int[] rang in rangs)
                {
                    // 250-300
                    List<string> valueinGroup = new List<string>();
                    foreach (string cell in fields)
                    {
                        string[] coor = cell.Split("_");
                        int rowIndex = int.Parse(coor[0]);
                        if (rowIndex < 0)
                        {
                            rowIndex = rang[1] + rowIndex;
                        }
                        else
                        {
                            rowIndex += rang[0] - 1;
                        }
                        int columnIndex = int.Parse(coor[1]);
                        string cellValue = "";
                        try
                        {
                            cellValue = Convert.ToString(values[rowIndex, columnIndex]);
                        }
                        catch (Exception ex) { }
                        valueinGroup.Add(cellValue);
                    }
                    groups.Add(valueinGroup);
                }
                callback(groups);

                var t = new Thread(() => {
                    try
                    {
                        foreach (int[] rang in rangs)
                        {
                            Excel.Range startCell = xlWorksheet.Cells[rang[0], 1];//not sure the second parameter needed?
                            Excel.Range endCell = xlWorksheet.Cells[rang[1], columnCount];//not sure the second parameter needed?

                            Excel.Workbook workbook = xlApp.Workbooks.Add();
                            Excel.Worksheet worksheet = workbook.Sheets.Add();
                            Excel.Range rng = xlWorksheet.Range[startCell, endCell];
                            Excel.Range rngDest = (Excel.Range)worksheet.Cells[1, 1];
                            rng.Copy(rngDest);
                            string outPath = outputDir + @"\out" + rang[0] + "_" + rang[1] + "";
                            workbook.SaveAs(outPath, Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            workbook.Close();
                            CloseWorkBook(workbook);
                        }
                    }
                    catch (Exception ex) { }
                    finally
                    {
                        closeExcel(xlApp, xlWorkbook, xlWorksheet, null);
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
            catch (Exception ex) { }
            finally
            {
               
            }
        }


        private static List<int[]> GetListOfRangeHasData(Object[,] values, int rowCount, int columnCount, int spaceRowBetweenGroup) 
        {
            
            List<int[]> ranges = new List<int[]>();
            
            int[] indexs = new int[2];
            indexs[0] = -1;
            int emptyRowCount = 0;
            for (int i = 1; i < rowCount; i++)
            {
                int emptyCells = 0;
                for (int j = 1; j <= columnCount; j++)
                {
                    string cellValue = "";
                    cellValue = Convert.ToString(values[i, j]);
                    if (cellValue == "")
                    {
                        emptyCells++;
                    }
                }
                if (emptyCells == columnCount)
                {
                    emptyRowCount++;
                }
                else
                {
                    emptyRowCount = 0;
                    if (indexs[0] == -1)
                    {
                        indexs[0] = i;
                    }
                }

                if (emptyRowCount == spaceRowBetweenGroup)
                {
                    indexs[1] = i - spaceRowBetweenGroup;
                    ranges.Add(indexs);

                    indexs = new int[2];
                    indexs[0] = -1;
                }
            }
            indexs[1] = rowCount;
            ranges.Add(indexs);

            return ranges;
        }

        public static List<List<string>> GetGroupValuesByGroupRow(string file, List<string> fields, int rowInGroup)
        {
            List<List<string>> groups = new List<List<string>>();
      
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(file);
                xlWorksheet = xlWorkbook.ActiveSheet;
                xlRange = xlWorksheet.UsedRange;
                int rowCount = xlWorksheet.UsedRange.Rows.Count;
                int columnCount = xlWorksheet.UsedRange.Columns.Count;

                int maxloops = rowCount / rowInGroup;
                var values = (xlRange.Value as Object[,]);

                for (int i = 0; i < maxloops; i++)
                {
                    List<string> valueinGroup = new List<string>();
                    foreach (string cell in fields)
                    {
                        string[] coor = cell.Split("-");
                        int rowIndex = int.Parse(coor[0]) + (i * rowInGroup);
                        int columnIndex = int.Parse(coor[1]);
                        string cellValue = "";
                        try
                        {
                            cellValue = Convert.ToString(values[rowIndex, columnIndex]);

                        }
                        catch (Exception ex) { }
                        valueinGroup.Add(cellValue);
                    }
                    groups.Add(valueinGroup);
                }
            }
            catch (Exception ex) { }
            finally
            {
                closeExcel(xlApp, xlWorkbook, xlWorksheet, null);
            }

            return groups;
        }

        public static List<String> SplitToMultiFiles(string filePath, int maxrows)
        {
            List<String> files = new List<string>();
            string startPath = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            string excelOutFilesFolderName = startPath + @"\excels";
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                bool exists = System.IO.Directory.Exists(excelOutFilesFolderName);

                if (exists)
                {
                    System.IO.Directory.Delete(excelOutFilesFolderName, true);
                }
                System.IO.Directory.CreateDirectory(excelOutFilesFolderName);


                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(filePath);
                xlWorksheet = xlWorkbook.ActiveSheet;

                int iRowCount = xlWorksheet.UsedRange.Rows.Count;
                int countColumns = xlWorksheet.UsedRange.Columns.Count;

                int maxloops = iRowCount / maxrows;
                int beginrow = 1; //skipping the header row.

                for (int i = 1; i <= maxloops; i++)
                {
                    Excel.Range startCell = xlWorksheet.Cells[beginrow, 1];//not sure the second parameter needed?
                    Excel.Range endCell = xlWorksheet.Cells[beginrow + maxrows - 1, countColumns];//not sure the second parameter needed?

                    Excel.Workbook workbook = xlApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.Sheets.Add();
                    Excel.Range rng = xlWorksheet.Range[startCell, endCell];
                    Excel.Range rngDest = (Excel.Range)worksheet.Cells[1, 1];
                    rng.Copy(rngDest);
                    string outPath = excelOutFilesFolderName + @"\out" + i + "";
                    workbook.SaveAs(outPath, Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    var misValue = Type.Missing;
                    workbook.Close();
                    files.Add(outPath);
                    beginrow = beginrow + maxrows;
                }
            }
            catch (Exception ex) { }
            finally
            {
                closeExcel(xlApp, xlWorkbook, xlWorksheet, null);
            }

            return files;
        }

        public static List<List<string>> GetDataFromFile(string file, int[] vs, bool skipHeader)
        {
            List<List<string>> table = new List<List<string>>();
        
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(file);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;
                int iRowCount = xlWorksheet.UsedRange.Rows.Count;
                int countColumns = xlWorksheet.UsedRange.Columns.Count;
                var values = (xlRange.Value as Object[,]);
                
                int a = skipHeader ? 2 : 1;
                List<string> vvv = new List<string>();
                for (int i = a; i <= iRowCount; i++)
                {

                    vvv = new List<string>();
                    for (int j = 1; j <= countColumns; j++)
                    {
                        string cellValue = "";
                        try
                        {
                            cellValue = Convert.ToString(values[i, j]);
                        }
                        catch (Exception ex) { }
                        vvv.Add(cellValue);
                    }

                    table.Add(vvv);
                }
            }
            catch (Exception ex) { }
            finally
            {
                closeExcel(xlApp, xlWorkbook, xlWorksheet, null);
            }
            return table;
        }

        private static void closeExcel(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel._Worksheet xlWorksheet, Excel.Range xlRange)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            if (xlRange != null)
            {
                Marshal.ReleaseComObject(xlRange);
            }
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }

            //close and release
            CloseWorkBook(xlWorkbook);
          
            //quit and release
            if (xlApp != null)
            {
                xlApp.Quit();
            }

            Marshal.ReleaseComObject(xlApp);
        }


        private static void CloseWorkBook(Excel.Workbook workbook) {
            try
            {
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
            }
            catch (Exception ex) { }
        }
    }


}
