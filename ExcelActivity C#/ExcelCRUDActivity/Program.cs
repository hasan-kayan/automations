using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelActivities
{
    class Program
    {
        public static void ReadExcelFile(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.ActiveSheet; // Use the active sheet

            Excel.Range range = worksheet.UsedRange;

            // Iterate through each cell in the range
            foreach (Excel.Range cell in range)
            {
                string cellValue = cell.Value2?.ToString(); // Get the value of the cell

                // Do something with the cell value
                Console.WriteLine("Cell [{0},{1}]: {2}", cell.Row, cell.Column, cellValue);
            }

            // Close Excel objects and release resources
            workbook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            worksheet = null;
            workbook = null;
            excelApp = null;
            GC.Collect();
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Please Enter Excel Path You Want to Work on:");
            string ExcelPath = Console.ReadLine();
            ReadExcelFile(ExcelPath);
            
        }
    }
}
