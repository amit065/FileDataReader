using SampleExcelReaderProj.BL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleExcelReaderProj.Host
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\aprakash\Desktop\CLTS_Cities_Data.xlsx";
            CityDataManupulation.ManipulateCityFromExcel(fileName);
        }
    }
}
