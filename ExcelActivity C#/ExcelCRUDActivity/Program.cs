using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelActivities
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Select the process:");
                Console.WriteLine("1. Read Cell");
                Console.WriteLine("2. Write Cell");
                Console.WriteLine("3. Read Range");
                Console.WriteLine("4. Write Range");
                Console.Write("Enter your choice: ");
                int process = int.Parse(Console.ReadLine());

                Console.Write("Enter the Excel file path: ");
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

                ReleaseComObject(workbook);
                ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        static void ReadCell(Excel.Workbook workbook)
        {
            Console.Write("Enter the cell address (e.g., A1): ");
            string cellAddress = Console.ReadLine();

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Excel.Range cell = worksheet.Range[cellAddress];
            string cellValue = cell.Value?.ToString();

            Console.WriteLine($"The value of the cell {cellAddress} is: {cellValue}");
        }

        static void WriteCell(Excel.Workbook workbook)
        {
            Console.Write("Enter the cell address (e.g., A1): ");
            string cellAddress = Console.ReadLine();

            Console.Write("Enter the data to write: ");
            string data = Console.ReadLine();

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Excel.Range cell = worksheet.Range[cellAddress];
            cell.Value = data;

            Console.WriteLine($"Data '{data}' written to cell {cellAddress} successfully.");
        }

        static void ReadRange(Excel.Workbook workbook)
        {
            Console.Write("Enter the range (e.g., A1:B5): ");
            string rangeAddress = Console.ReadLine();

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Excel.Range range = worksheet.Range[rangeAddress];
            object[,] values = range.Value as object[,];

            int rowCount = values?.GetLength(0) ?? 0;
            int colCount = values?.GetLength(1) ?? 0;

            Console.WriteLine("Range values:");
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    string cellValue = values[row, col]?.ToString();
                    Console.Write($"{cellValue}\t");
                }
                Console.WriteLine();
            }
        }

        static void WriteRange(Excel.Workbook workbook)
        {
            Console.Write("Enter the range (e.g., A1:B5): ");
            string rangeAddress = Console.ReadLine();

            Console.WriteLine("Enter the data table:");
            DataTable dataTable = new DataTable();

            Console.Write("Enter the number of columns: ");
            int colCount = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter column names:");
            for (int col = 0; col < colCount; col++)
            {
                Console.Write($"Column {col + 1}: ");
                string columnName = Console.ReadLine();
                dataTable.Columns.Add(columnName);
            }

            Console.Write("Enter the number of rows: ");
            int rowCount = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the data:");

            for (int row = 0; row < rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 0; col < colCount; col++)
                {
                    Console.Write($"Enter the value for cell [{row + 1},{col + 1}]: ");
                    string cellValue = Console.ReadLine();
                    dataRow[col] = cellValue;
                }
                dataTable.Rows.Add(dataRow);
            }

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            Excel.Range range = worksheet.Range[rangeAddress];
            range.Value = dataTable;

            Console.WriteLine("Data written to the range successfully.");
        }

        static void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception occurred while releasing COM object: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
