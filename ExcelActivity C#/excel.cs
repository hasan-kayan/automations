using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the file path:");  
        string filePath = Console.ReadLine();

        if (File.Exists(filePath))
        {
            Console.WriteLine("Excel file found. Opening the file...");

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            Console.WriteLine("Select a process:");
            Console.WriteLine("1. Read cell");
            Console.WriteLine("2. Write cell");
            Console.WriteLine("3. Read range");
            Console.WriteLine("4. Write range");

            int choice = Convert.ToInt32(Console.ReadLine());

            switch (choice)
            {
                case 1:
                    Console.WriteLine("Enter the cell address (e.g., A1):");
                    string cellAddress = Console.ReadLine();
                    Excel.Range cell = worksheet.Range[cellAddress];
                    string cellValue = cell.Value.ToString();
                    Console.WriteLine("Cell value: " + cellValue);
                    break;
                case 2:
                    Console.WriteLine("Enter the cell address (e.g., A1):");
                    cellAddress = Console.ReadLine();
                    Console.WriteLine("Enter the value to write:");
                    string cellValueToWrite = Console.ReadLine();
                    cell = worksheet.Range[cellAddress];
                    cell.Value = cellValueToWrite;
                    Console.WriteLine("Value written successfully.");
                    break;
                case 3:
                    Console.WriteLine("Enter the range address (e.g., A1:B5):");
                    string rangeAddress = Console.ReadLine();
                    Excel.Range range = worksheet.Range[rangeAddress];
                    object[,] rangeValues = range.Value;
                    int rowCount = rangeValues.GetLength(0);
                    int colCount = rangeValues.GetLength(1);
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            string value = rangeValues[row, col]?.ToString() ?? string.Empty;
                            Console.WriteLine("Cell[" + row + "," + col + "]: " + value);
                        }
                    }
                    break;
                case 4:
                    Console.WriteLine("Enter the range address (e.g., A1:B5):");
                    rangeAddress = Console.ReadLine();
                    Console.WriteLine("Enter the values to write (separated by commas):");
                    string valuesToWrite = Console.ReadLine();
                    string[] cellValues = valuesToWrite.Split(',');
                    range = worksheet.Range[rangeAddress];
                    range.Value = cellValues;
                    Console.WriteLine("Values written successfully.");
                    break;
                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
            }

            workbook.Save();
            workbook.Close();
            excelApp.Quit();
        }
        else
        {
            Console.WriteLine("File not found. Exiting the program.");
        }

        Console.WriteLine("Press any key to exit.");
        Console.ReadKey();
    }
}
