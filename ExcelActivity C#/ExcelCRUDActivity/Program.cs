using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        // Get Excel file path from the user
        Console.WriteLine("Enter the Excel file path:");
        string filePath = Console.ReadLine();

        // Create an Excel application instance
        Excel.Application excelApp = new Excel.Application();

        try
        {
            // Open the Excel workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            // Get the first worksheet
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Get the used range of the worksheet
            Excel.Range usedRange = worksheet.UsedRange;

            // Convert the used range to a DataTable
            DataTable dataTable = ConvertRangeToDataTable(usedRange);

            // Display the DataTable (optional)
            Console.WriteLine("DataTable contents:");
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item.ToString() + "\t");
                }
                Console.WriteLine();
            }

            // Write the DataTable back to the worksheet starting from cell A1
            Excel.Range targetRange = worksheet.Range["A1"].Resize[dataTable.Rows.Count, dataTable.Columns.Count];
            targetRange.Value = dataTable;

            // Save and close the workbook
            workbook.Save();
            workbook.Close();

            Console.WriteLine("DataTable written to Excel successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Clean up the Excel application
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }

        Console.ReadLine();
    }

    static DataTable ConvertRangeToDataTable(Excel.Range range)
    {
        DataTable dataTable = new DataTable();

        // Add columns to the DataTable
        foreach (Excel.Range column in range.Columns)
        {
            dataTable.Columns.Add(column.Value.ToString());
        }

        // Add rows to the DataTable
        for (int row = 2; row <= range.Rows.Count; row++)
        {
            DataRow dataRow = dataTable.NewRow();
            for (int col = 1; col <= range.Columns.Count; col++)
            {
                dataRow[col - 1] = range.Cells[row, col].Value;
            }
            dataTable.Rows.Add(dataRow);
        }

        return dataTable;
    }
}
