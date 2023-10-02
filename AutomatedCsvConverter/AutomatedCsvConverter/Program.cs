using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Console.WriteLine("Enter the full or relative path to the Excel file: ");
        string inputFilePath = Console.ReadLine();

        // Ensure the provided path is valid
        if (!File.Exists(inputFilePath))
        {
            Console.WriteLine("Invalid file path. Exiting...");
            return;
        }

        Console.WriteLine("Enter the desired CSV file name: ");
        string outputFileName = Console.ReadLine();

        ConvertExcelToCsvWithColumns(inputFilePath, outputFileName);

        Console.WriteLine("Conversion successful!");
    }
    static void ConvertExcelToCsvWithColumns(string inputFilePath, string outputFileName)
    {
        try
        {
            using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assumes only one worksheet

                // Create a StreamWriter to write to a CSV file
                using (StreamWriter sw = new StreamWriter(outputFileName))
                {
                    // Write header (column names)
                    sw.WriteLine(string.Join(",", worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column].Select(cell => cell.Text)));

                    // Copy all records under the columns to the CSV file
                    for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        var rowData = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column].Select(cell => cell.Text);
                        sw.WriteLine(string.Join(",", rowData));
                    }
                }

                Console.WriteLine($"CSV file with columns created: {outputFileName}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

        // Convert Excel column letter to column index (A -> 1, B -> 2, ..., Z -> 26, AA -> 27, ...)
    //static int ExcelColumnToIndex(string columnName)
    //{
    //    int index = 0;
    //    foreach (char c in columnName)
    //    {
    //        index = index * 26 + (c - 'A' + 1);
    //    }
    //    return index;
    //}
}