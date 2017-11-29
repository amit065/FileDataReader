using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SampleListToExcel
{
    public class Car
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public string  Model { get; set; }

        public static void Main(string[] args)
        {
            ExportToExcel(GetList());
        }
        public static List<Car> GetList()
        {
            List<Car> cars = new List<Car>()
            {
                 new Car {Name = "Toyota", Color = "Red", Model = "1"},
                 new Car {Name = "Honda", Color = "Blue", Model = "2"},
                  new Car {Name = "Mazda", Color = "Green", Model = "3"}
            };

            return cars;
        }

        public static void ExportToExcel(List<Car> cars)
        {
            Excel.Application excel = new Excel.Application();

            excel.Workbooks.Add();

            Excel._Worksheet workSheet = excel.ActiveSheet;

            try
            {

                workSheet.Cells[1, "A"] = "Name";
                workSheet.Cells[1, "B"] = "Color";
                workSheet.Cells[1, "C"] = "Model";


                int row = 2; 
                foreach (Car car in cars)
                {
                    workSheet.Cells[row, "A"] = car.Name;
                    workSheet.Cells[row, "B"] = car.Color;
                    workSheet.Cells[row, "C"] = car.Model;

                    row++;
                }

                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelData.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);

              
            }
            catch (Exception exception)
            {
                
            }
            finally
            {
                // Quit Excel application
                excel.Quit();

                // Release COM objects (very important!)
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                excel = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
        }

    }


}
