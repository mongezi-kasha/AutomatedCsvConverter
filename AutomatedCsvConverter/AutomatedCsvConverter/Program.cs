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

        Console.WriteLine("Enter the number of columns: ");
        int numColumns = int.Parse(Console.ReadLine());

        string[] selectedColumns = new string[numColumns];
        for (int i = 0; i < numColumns; i++)
        {
            Console.WriteLine($"Enter the name of column {i + 1}: ");
            selectedColumns[i] = Console.ReadLine();
        }

        string column1 = ""; string column2= "";
        ConvertExcelToCsv(inputFilePath, outputFileName, column1, column2);

        Console.WriteLine("Conversion successful!");
    }

    static void ConvertExcelToCsv(string inputFilePath, string outputFileName, string column1, string column2)
    {
        try
        {
            using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assumes only one worksheet

                // Create a StreamWriter to write to a CSV file
                using (StreamWriter sw = new StreamWriter(outputFileName))
                {
                    // Write header
                    sw.WriteLine($"{column1},{column2},{string.Join(",", worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column].Select(cell => cell.Text))}");

                    // Copy all records under the specified columns to the CSV file
                    int columnIndex1 = ExcelColumnToIndex(column1);
                    int columnIndex2 = ExcelColumnToIndex(column2);

                    // Check if the columns exist in the worksheet
                    if (columnIndex1 >= worksheet.Dimension.Start.Column && columnIndex1 <= worksheet.Dimension.End.Column &&
                        columnIndex2 >= worksheet.Dimension.Start.Column && columnIndex2 <= worksheet.Dimension.End.Column)
                    {
                        // Iterate through used rows in the specified columns
                        var column1Data = worksheet.Cells[worksheet.Dimension.Start.Row + 1, columnIndex1, worksheet.Dimension.End.Row, columnIndex1].Select(cell => cell.Text);
                        var column2Data = worksheet.Cells[worksheet.Dimension.Start.Row + 1, columnIndex2, worksheet.Dimension.End.Row, columnIndex2].Select(cell => cell.Text);

                        for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                        {
                            var rowData = new List<string>
                        {
                            column1Data.ElementAt(rowNum - 2),
                            column2Data.ElementAt(rowNum - 2),
                            string.Join(",", worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column].Select(cell => cell.Text))
                        };

                            sw.WriteLine(string.Join(",", rowData));
                        }
                    }
                    else
                    {
                        Console.WriteLine($"One or both of the specified columns do not exist in the worksheet.");
                    }
                }

                Console.WriteLine($"CSV file saved to: {outputFileName}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Convert Excel column letter to column index (A -> 1, B -> 2, ..., Z -> 26, AA -> 27, ...)
    static int ExcelColumnToIndex(string columnName)
    {
        int index = 0;
        foreach (char c in columnName)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }
}