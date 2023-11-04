using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace ExpenseReportConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\\code\\ExpenseReportConverter\\Development References\\YMB Rentals Management Company LLC_Transaction Detail by  Account-5.xlsx"; // Replace with your Excel file path
            List<string> excelData = ReadDataFromExcel(excelFilePath);

            //foreach (var cellValue in excelData)
            //{
            //    Console.WriteLine(cellValue);
            //}

            string newExcelFilePath = @"C:\code\ExpenseReportConverter\TestOutputFiles\output_file.xlsx"; // Specify the output file path
            CreateExcelFileFromData(excelData, newExcelFilePath);


            // APPEND LOGIC
            //string excelFilePath = "path_to_existing_excel_file.xlsx";
            List<string> newData = new List<string> { "New Data 1", "New Data 2" };

            AppendDataToExcel(excelFilePath, newData);
        }

        public static List<string> ReadDataFromExcel(string filePath)
        {
            List<string> data = new List<string>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Get the first worksheet

                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var cellValue = worksheet.Cells[row, 1].Text;
                        data.Add(cellValue);
                    }
                }
            }

            return data;
        }

        public static void CreateExcelFileFromData(List<string> data, string outputFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Define the number of rows you want to format
                int rowsToFormat = Math.Min(2, data.Count);

                // Write the data to the Excel sheet
                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 1, 1].Value = data[i];

                    if (i < rowsToFormat)
                    {
                        // Apply formatting to the first 'rowsToFormat' rows
                        worksheet.Cells[i + 1, 1].Style.Font.Bold = true;
                        worksheet.Cells[i + 1, 1].Style.Font.Size = 14;
                        worksheet.Cells[i + 1, 1].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                    }


                    // Remove the "$" sign and parse the string as a double
                    if (double.TryParse(data[i].Replace("$", ""), out double numericValue))
                    {
                        worksheet.Cells[i + 1, 1].Value = numericValue;

                        // Apply formatting to the cell based on the numeric value
                        var cell = worksheet.Cells[i + 1, 1];
                        if (numericValue >= 0)
                        {
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                        }
                        else
                        {
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSalmon);
                        }
                    }
                    else
                    {
                        // Handle the case where the value couldn't be parsed as a double
                        worksheet.Cells[i + 1, 1].Value = data[i];
                    }

                    //TODO: when searching for the total money amount, scan the column headers for 'Amount' then go down
                    // also can find the totals easily by finding the intersection of 'Amount' column and 'Total for [expense]' row

                }

                int columnToFormat = 2; // Column B
                using (var range = worksheet.Cells[1, columnToFormat, data.Count, columnToFormat])
                {
                    range.Style.Numberformat.Format = "$#,##0.00";
                }



                // Save the Excel package to a file
                package.SaveAs(new FileInfo(outputFilePath));
            }
        }



        public static void AppendDataToExcel(string excelFilePath, List<string> newData)
        {
            // Load the existing Excel file
            FileInfo existingFile = new FileInfo(excelFilePath);

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    // Find the last used row in the worksheet
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    // Determine the starting row for appending new data
                    int newRow = lastUsedRow + 1;

                    // Append the new data to the worksheet
                    foreach (var item in newData)
                    {
                        worksheet.Cells[newRow, 1].Value = item;
                        newRow++;
                    }

                    // Save the updated Excel file
                    package.Save();
                }
            }
        }


    }
}