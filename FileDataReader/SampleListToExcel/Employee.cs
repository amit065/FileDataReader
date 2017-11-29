using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SampleListToExcel
{
    public class Employee
    {
        public int Empid { get; set; }
        public string Empname { get; set; }
        public string City { get; set; }

       public  static void Main(string[] args)
        {
            ListToExcel(GetEmpList());
        }
        public static List<Employee> GetEmpList()
        {
            List<Employee> list = new List<Employee>();
            Employee e = new Employee();
            e.Empid = 2001;
            e.Empname = "DEVESH";
            e.City = "NOIDA";
            list.Add(e);
            e = new Employee();
            e.Empid = 2002;
            e.Empname = "NIKHIL";
            e.City = "DELHI";
            list.Add(e);
            e = new Employee();
            e.Empid = 2003;
            e.Empname = "AVINASH";
            e.City = "NAGPUR";
            list.Add(e);
            e = new Employee();
            e.Empid = 2004;
            e.Empname = "SHRUTI";
            e.City = "NOIDA";
            list.Add(e);
            e = new Employee();
            e.Empid = 2004;
            e.Empname = "ROLI";
            e.City = "KANPUR";
            list.Add(e);
            return list;
        }

        public static void ListToExcel(List<Employee> listOfEmployee)
        {

            Excel.Application xlApp = new Excel.Application();
          
            xlApp.Visible = true;
  
            string workbookPath = @"C:\Users\aprakash\Desktop\Testing.xlsx";
            var workbook = xlApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);

            var sheet = (Excel.Worksheet)workbook.Sheets[1];

            var range = sheet.get_Range("A1", "A1");
            range.Value2 = "test"; 


            string cellName;
            int counter = 1;
            foreach (var item in listOfEmployee)
            {
                cellName = "A" + counter.ToString();
                var rang = sheet.get_Range(cellName, cellName);
                rang.Value2 = item.ToString();
                ++counter;
            }

        }
    }


}
