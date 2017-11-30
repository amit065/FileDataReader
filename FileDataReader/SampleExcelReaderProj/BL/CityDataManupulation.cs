using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SampleExcelReaderProj.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace SampleExcelReaderProj.BL
{
    public static class CityDataManupulation
    {
        public static void ManipulateCityFromExcel(string filePath)
        {

            List<City> cities = ReadCityFromExcelFile(filePath);

            ExportToExcel(cities);

            GenerateInsertScript(cities);
        }

        private static List<City> ReadCityFromExcelFile(string filePath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int row = xlRange.Rows.Count;
            int columnl = xlRange.Columns.Count;

            List<City> cities = new List<City>();

            for (int i = 1; i <= row; i++)
            {
                cities.Add(new City
                {
                    CityName = Convert.ToString((xlRange.Cells[i, 1] as Excel.Range).Value2),
                    StateCode = Convert.ToString((xlRange.Cells[i, 2] as Excel.Range).Value2),
                    CountryCode = Convert.ToString((xlRange.Cells[i, 3] as Excel.Range).Value2),
                    Latitude = Convert.ToString((xlRange.Cells[i, 4] as Excel.Range).Value2),
                    Longitude = Convert.ToString((xlRange.Cells[i, 5] as Excel.Range).Value2),
                    IsEnabled = Convert.ToString((xlRange.Cells[i, 6] as Excel.Range).Value2),
                    IataCityCode = Convert.ToString((xlRange.Cells[i, 7] as Excel.Range).Value2),
                    FullTextColumn = (i == 1 ? "FullTextSearch" : GetFullTextSearch((string)(xlRange.Cells[i, 1] as Excel.Range).Value2, Convert.ToString((xlRange.Cells[i, 7] as Excel.Range).Value2)))
                });
            }

            return cities;
        }

        private static string GetFullTextSearch(string cityName, string IataCityCode)
        {
            string textSearch = IataCityCode != null ? IataCityCode : string.Empty;

            var CityNameloweCase = cityName.ToLower();

            foreach (var part in CityNameloweCase.Split(' '))
            {
                if (part.Length > 3)
                {

                    for (int i = 3; i <= part.Length; i++)
                    {
                        textSearch += " " + part.Substring(0, i);
                    }
                }
                else
                {
                    textSearch += " " + part;
                }
            }

            return textSearch;
        }

        public static void ExportToExcel(List<City> cities)
        {
            Excel.Application excel = new Excel.Application();

            excel.Workbooks.Add();

            Excel._Worksheet workSheet = excel.ActiveSheet;

            workSheet.Cells[1, "A"] = "CityName";
            workSheet.Cells[1, "B"] = "StateCode";
            workSheet.Cells[1, "C"] = "CountryCode";
            workSheet.Cells[1, "D"] = "Latitude";
            workSheet.Cells[1, "E"] = "Longitude";
            workSheet.Cells[1, "F"] = "IsEnabled";
            workSheet.Cells[1, "G"] = "IataCityCode";
            workSheet.Cells[1, "H"] = "FullTextColumn";

            int row = 1;
            foreach (City city in cities)
            {
                workSheet.Cells[row, "A"] = city.CityName;
                workSheet.Cells[row, "B"] = city.StateCode;
                workSheet.Cells[row, "C"] = city.CountryCode;
                workSheet.Cells[row, "D"] = city.Latitude;
                workSheet.Cells[row, "E"] = city.Longitude;
                workSheet.Cells[row, "F"] = city.IsEnabled;
                workSheet.Cells[row, "G"] = city.IataCityCode;
                workSheet.Cells[row, "H"] = city.FullTextColumn;

                row++;
            }

            workSheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

            string fileName = string.Format(@"{0}\ExcelCityData.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            workSheet.SaveAs(fileName);

            excel.Quit();

            if (excel != null)
                Marshal.ReleaseComObject(excel);

            if (workSheet != null)
                Marshal.ReleaseComObject(workSheet);


            excel = null;
            workSheet = null;

            // Force garbage collector cleaning
            GC.Collect();

        }

        public static void GenerateInsertScript(List<City> cities)
        {
            using (StreamWriter writer = new StreamWriter(@"C:\Users\aprakash\Desktop\CityScript.txt", false))
            {

                foreach (City city in cities)
                {
                    writer.WriteLine("Insert into Cities (CityName, StateCode, CountryCode, Latitude, Longitude, IsEnabled, IataCityCode, FullTextColumn) values('" + city.CityName + "' , '" + city.StateCode + "' , '" + city.CountryCode + "' , '" + city.Latitude + "' , '" + city.Longitude + "' , '" + city.IsEnabled + "' , '" + (city.IataCityCode != null ? city.IataCityCode : city.IataCityCode = "NULL") + "' , '" + city.FullTextColumn + "');");

                }

            }  
          
        }

    }
}

