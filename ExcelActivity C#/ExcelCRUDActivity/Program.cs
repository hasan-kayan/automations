using System;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelActivities
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel Activities");
            Console.WriteLine("Select the process:");
            Console.WriteLine("1. Read Cell");
            Console.WriteLine("2. Write Cell");
            Console.WriteLine("3. Read Range");
            Console.WriteLine("4. Write Range");

            string input = Console.ReadLine();

            if (!string.IsNullOrEmpty(input))
            {
                int process = int.Parse(input);

                // Rest of your code based on the value of 'process'
            }
            else
            {
                Console.WriteLine("Invalid input. Please provide a valid process.");
            }


            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            switch (process)
            {
                case 1:
                    ReadCell(workbook);
                    break;
                case 2:
                    WriteCell(workbook);
                    break;
                case 3:
                    ReadRange(workbook);
                    break;
                case 4:
                    WriteRange(workbook);
                    break;
                default:
                    Console.WriteLine("Invalid process selection.");
                    break;
            }

            workbook.Save();
            workbook.Close();
            excelApp.Quit();

            ReleaseObject(workbook);
            ReleaseObject(excelApp);

            Console.WriteLine("Excel activities completed.");
            Console.ReadLine();
        }

        static void ReadCell(Excel.Workbook workbook)
        {
            Console.WriteLine("Enter the cell address (e.g., A1):");
            string cellAddress = Console.ReadLine();

            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.get_Range(cellAddress);

            if (range != null)
            {
                string cellValue = range.Value?.ToString();
                Console.WriteLine($"Cell {cellAddress} value: {cellValue}");
            }
            else
            {
                Console.WriteLine($"Cell {cellAddress} does not exist.");
            }
        }

        static void WriteCell(Excel.Workbook workbook)
        {
            Console.WriteLine("Enter the cell address (e.g., A1):");
            string cellAddress = Console.ReadLine();

            Console.WriteLine("Enter the data to write:");
            string data = Console.ReadLine();

            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.get_Range(cellAddress);

            if (range != null)
            {
                range.Value = data;
                Console.WriteLine($"Data '{data}' written to cell {cellAddress} successfully.");
            }
            else
            {
                Console.WriteLine($"Cell {cellAddress} does not exist.");
            }
        }

        static void ReadRange(Excel.Workbook workbook)
        {
            Console.WriteLine("Enter the range address (e.g., A1:B5):");
            string rangeAddress = Console.ReadLine();

            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.get_Range(rangeAddress);

            if (range != null)
            {
                object[,] values = range.Value2 as object[,];

                if (values != null)
                {
                    int rows = values.GetLength(0);
                    int columns = values.GetLength(1);

                    for (int row = 1; row <= rows; row++)
                    {
                        for (int column = 1; column <= columns; column++)
                        {
                            string cellValue = values[row, column]?.ToString();
                            Console.Write($"{cellValue}\t");
                        }

                        Console.WriteLine();
                    }
                }
            }
            else
            {
                Console.WriteLine($"Range {rangeAddress} does not exist.");
            }
        }

        static void WriteRange(Excel.Workbook workbook)
        {
            Console.WriteLine("Enter the range address (e.g., A1:B5):");
            string rangeAddress = Console.ReadLine();

            Console.WriteLine("Enter the data table (comma-separated values):");
            string data = Console.ReadLine();

            DataTable dataTable = new DataTable();

            foreach (string columnName in data.Split(','))
            {
                dataTable.Columns.Add(columnName);
            }

            Console.WriteLine("Enter the rows of data (comma-separated values). Type 'done' when finished:");

            while (true)
            {
                string rowData = Console.ReadLine();

                if (rowData.Equals("done", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                DataRow row = dataTable.NewRow();
                row.ItemArray = rowData.Split(',');

                dataTable.Rows.Add(row);
            }

            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.get_Range(rangeAddress);

            if (range != null)
            {
                range.Value = dataTable;
                Console.WriteLine($"Data table written to range {rangeAddress} successfully.");
            }
            else
            {
                Console.WriteLine($"Range {rangeAddress} does not exist.");
            }
        }

        static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception occurred while releasing object: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
