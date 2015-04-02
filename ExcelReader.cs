using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XapTesterStatus
{
    class ExcelReader
    {
        public void readRows(string fileName, int sheetNum, int startRow, int endRow, List<string> rows) 
        {
            bool IgnoreReadOnlyRecommended = true; //MSDN: True to have Microsoft Excel not display the read-only recommended message
            bool ReadOnly = true; //MSDN: True to open the workbook in read-only mode.

            Application app = new Application();
            Workbooks books = app.Workbooks;
            Workbook wb = books.Open(fileName,
                                Type.Missing, ReadOnly, Type.Missing, Type.Missing,
                                Type.Missing, IgnoreReadOnlyRecommended, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);

            Worksheet statsSheet = (Worksheet)wb.Sheets[sheetNum];
            Range excelRange = statsSheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(
                XlRangeValueDataType.xlRangeValueDefault);

            int num_rows = valueArray.GetLength(0);
            int num_cols = valueArray.GetLength(1);
            for (int idx = startRow; idx <= endRow; idx++)
            {
                string row = "";
                for (int col = 1; col <= num_cols; col++)
                {
                    string cell= valueArray[idx, col].ToString();
                    if ( (col != num_cols) && (col != (num_cols-1)) )
                    {
                        double num = 0;
                        if (double.TryParse(cell, out num)) {
                            cell = num.ToString("0.0");
                        }
                        row += cell + ";";
                    }
                    else 
                    {
                        //if the last 2 columns is numeric, make it percentage
                        double percent = 0;
                        if(double.TryParse(cell, out percent)){
                            percent = 100 * percent;
                            if (percent < 0) { percent = 0; } //fix for divide by zero
                            cell = percent.ToString("0.0") + "%";
                        }
                        if (col == num_cols)
                        {
                            row += cell;
                        }
                        else 
                        {
                            row += cell + ";";
                        }
                    }
                }
                rows.Add(row.Trim());
            }
            wb.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            books.Close();
            app.Quit();
            ReleaseCOMObject(excelRange);
            ReleaseCOMObject(statsSheet);
            ReleaseCOMObject(wb);
            ReleaseCOMObject(books);
            ReleaseCOMObject(app);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ReleaseCOMObject(object obj)
        {
            if (obj != null)
                Marshal.FinalReleaseComObject(obj);
            obj = null;
        }
    }
}
