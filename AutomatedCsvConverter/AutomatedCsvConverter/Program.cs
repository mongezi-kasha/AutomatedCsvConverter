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

        Console.WriteLine("Enter the password for the Excel file: ");
        string excelPassword = Console.ReadLine();

        ConvertPasswordProtectedExcelToCsvWithColumns(inputFilePath, outputFileName, excelPassword);

        Console.WriteLine("Conversion successful!");
    }

    static void ConvertPasswordProtectedExcelToCsvWithColumns(string inputFilePath, string outputFileName, string excelPassword)
    {
        try
        {
            using (var package = new ExcelPackage(new FileInfo(inputFilePath), excelPassword))
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

}