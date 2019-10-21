using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;


namespace Task2.Helpers
{
    public static class ExcelHelper
    {
        public static ArrayList ExcelAccessor(string filePath)
        {
            ArrayList values = new ArrayList();
            Application xlsApp = new Application();
            Workbook workbook = xlsApp.Workbooks.Open(Path.GetFullPath(filePath), ReadOnly: true);
            Worksheet worksheet = (Worksheet)workbook.Sheets.get_Item(1);
            Range cellsRange = worksheet.UsedRange;

            values.Add((object[,])cellsRange.get_Value(XlRangeValueDataType.xlRangeValueDefault));
            values.Add(worksheet);
            values.Add(workbook);
            values.Add(xlsApp);

            return values;
        }
        public static void CloseExcel(Workbook workbook, Application xlsApp)
        {
            workbook.Close(true);
            Marshal.ReleaseComObject(workbook);
            xlsApp.Quit();
            Marshal.FinalReleaseComObject(xlsApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
