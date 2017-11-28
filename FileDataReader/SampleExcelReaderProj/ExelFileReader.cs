using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SampleExcelReaderProj
{
    class ExelFileReader
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(@"C:\Users\aprakash\Desktop\CLTS_Cities_Data.xlsx"));

            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);

            Excel.Range xlRange = xlWorksheet.UsedRange; 

            object[,] valueArray = (object[,])xlRange.get_Value(
                        Excel.XlRangeValueDataType.xlRangeValueDefault);

            for (int row = 1; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 1; col <= xlWorksheet.UsedRange.Columns.Count; ++col)
                {
                    if (valueArray[row, col] != null)
                        Console.Write(valueArray[row, col].ToString());
                        Console.Write("     ");
                }
                Console.WriteLine();
               
            }

            xlWorkbook.Close(false);

            xlApp.Quit();

            Console.ReadLine();
        }
    }
}
