using System;
using System.Collections.Generic;
using System.Linq;
using Task2.Helpers;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;

namespace Task2.Helpers
{
    public static class DataReader
    {
        static System.Data.DataTable dataTable = new System.Data.DataTable();

        public static DataTable GetExcelTableData()
         {
            string excelFinalPath = ConfigurationManager.AppSettings["UserData"];
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = application.Workbooks.Open(excelFinalPath);
     
            for (int i = 1; i <= workBook.Sheets.Count; i++)
            {
                Worksheet worksheet = workBook.Worksheets[i];

                Range excelCell = worksheet.UsedRange;
                Object[,] sheetValues = (Object[,])excelCell.Value;
                int noOfRows = sheetValues.GetLength(0);
                int noOfColumns = sheetValues.GetLength(1);

                //add column names to datatable
                
                for (int j = 1; j <= noOfColumns; j++)
                {
                    dataTable.Columns.Add(new DataColumn(((Range)worksheet.Cells[1, j]).Value));
                }

                //as first column has header, start at second row
                for (int x = 2; x <= noOfRows; x++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int l = 1; l <= noOfColumns; l++)
                    {
                        dataRow[l - 1] = ((Range)worksheet.Cells[x, l]).Value;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            workBook.Close(false, excelFinalPath, null);
            Marshal.ReleaseComObject(workBook);
            return dataTable;
        }

    }
}
