using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using static System.Net.WebRequestMethods;
using System.Runtime.ConstrainedExecution;
using System.Net;

namespace ExpenseReportConverter
{
    internal class Program
    {
        public static int InputsHeaderRow { get; set; }
        public static int OutputMasterK1HeaderRow { get; set; }
        static void Main(string[] args)
        {
            // prompt user for name of master k-1 doc
            string masterK1DocName = @"Master K-1.xlsx";
            InputsHeaderRow = 5;
            OutputMasterK1HeaderRow = 3;
            // ASSUMPTION: name of column to search for a match is named 'Street' in K1 sheet

            OutputLine("Master K-1 output File: " + masterK1DocName);

            // Get the application's base directory
            string directoryPath = @"C:\code\ExpenseReportConverter\Development References"; //AppDomain.CurrentDomain.BaseDirectory;
            OutputLine("Looking for files in directory: " + directoryPath);

            //TODO: prompt for accept defaults - Inputs header row of 5 and ouput header row of 3


            //set license information for the excel library nuget package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //set up exception handling
            System.AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;


            DirectoryInfo directoryInfo = new DirectoryInfo(directoryPath);

            if (directoryInfo.Exists)
            {
                FileInfo[] allFiles = directoryInfo.GetFiles();
                List<FileInfo> inputSpreadsheetFiles = new List<FileInfo>();
                FileInfo masterK1FileInfo = null;

                foreach (FileInfo file in allFiles)
                {
                    if (!file.Name.Equals(masterK1DocName) && file.Extension.Equals(".xlsx"))
                    {
                        inputSpreadsheetFiles.Add(file);
                    }
                    if (file.Name.Equals(masterK1DocName))
                    {
                        masterK1FileInfo = file;
                    }
                }
                if (masterK1FileInfo == null)
                {
                    throw new Exception("The master K-1 file specified was not found.");
                }

                // Check if main output K1 file is open, if it is prompt user to close it.
                SaveFile(masterK1FileInfo);

                OutputLine(inputSpreadsheetFiles.Count + " file(s) were found for processing.");
                OutputLine("    _____");
                OutputLine("   /\\\\   \\\\");
                OutputLine("  /  \\\\   \\\\            <=======================================>");
                OutputLine(" /    \\\\___\\\\            WELCOME TO THE EXPENSE REPORT CONVERTER");
                OutputLine("/___________\\\\          <=======================================>");
                OutputLine("|   ______   |");
                OutputLine("|  |      |  |                   - Foster");
                OutputLine("|__|______|__|");
                OutputLine();

                foreach (FileInfo file in inputSpreadsheetFiles)
                {
                    OutputLine("Starting to process: " + file.Name + "...");
                    // parse address and expenses from input excel sheet
                    InputExcelReport ier = ParseInputExcel(file);

                    // Output ier data to master K-1 sheet.
                    //OutputLine("Apending input data for " + ier.Address + " to " + masterK1DocName + "...");
                    string apendExcelFilePath = masterK1FileInfo.FullName;
                    AppendDataToExcel(apendExcelFilePath, ier);

                    // Output ier data to pdf sheet.

                }
            }
            else
            {
                OutputLine("The specified directory does not exist.");
            }

            //TODO: specify where output log goes.
            //OutputLine("The output log can be found at " + )

            //prompt user to exit the application
            OutputLine("Press Enter to exit application.");
            Console.ReadLine();
            Environment.Exit(1);

        }

        public static InputExcelReport ParseInputExcel(FileInfo file)
        {
            InputExcelReport ier = new InputExcelReport();
            ier.FileInfo = file;
            ier.Address = file.Name.Contains('&') ? file.Name.Substring(0, file.Name.IndexOf('&')) : file.Name.Substring(0, file.Name.IndexOf(file.Extension));
            ier.Expenses = ReadDataFromExcel(ier.FileInfo.FullName);
            return ier;
        }

        public static Dictionary<string, double> ReadDataFromExcel(string filePath)
        {
            Dictionary<string, double> expenses = new Dictionary<string, double>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Get the first worksheet

                if (worksheet != null)
                {
                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;

                    int amountColumnIndex = 0;

                    for(int column = 1; column <= columnCount; column++)
                    {
                        if (worksheet.Cells[InputsHeaderRow, column].Text.Equals("Amount"))
                        {
                            amountColumnIndex = column;
                        }
                    }

                    if(amountColumnIndex == 0)
                    {
                        throw new Exception("ERROR: couldn't find Amount column in inputted file");
                    }

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var cellValue = worksheet.Cells[row, 1].Text;
                        if (cellValue.Contains("Total for"))
                        {
                            string expenseTitle = cellValue.Substring(cellValue.IndexOf("Total for ") + 10);
                            var amountValue = worksheet.Cells[row, amountColumnIndex].Text;
                            // Remove the "$" sign and parse the string as a double
                            if (double.TryParse(amountValue.Replace("$", ""), out double numericValue))
                            {
                                expenses.Add(expenseTitle, numericValue);
                            }
                            else
                            {
                                // Handle the case where the value couldn't be parsed as a double
                                OutputLine("Couldn't parse data: " + amountValue);
                            }
                        }
                    }                    
                }
            }

            return expenses;
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



        public static void AppendDataToExcel(string excelFilePath, InputExcelReport report)
        {
            // Load the existing Excel file
            FileInfo existingFile = new FileInfo(excelFilePath);

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                char successIcon = '\u2713';
                char failIcon = 'x';

                if (worksheet != null)
                {
                    // IDENTIFY ALL COLUMNS 
                    Dictionary<string, int> columnTitles = new Dictionary<string, int>();
                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;

                    for (int column = 1; column <= columnCount; column++)
                    {
                        string cellValue = worksheet.Cells[OutputMasterK1HeaderRow, column].Text.Trim();
                        if (!cellValue.Trim().Equals(""))
                        {
                            if (!cellValue.Equals("Administration"))
                            {
                                columnTitles.Add(cellValue, column);
                            }
                        }
                    }

                    // IDENTIFY ROW OF ADDRESS TO APPEND TO
                    int reportAddressRow = 0;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var cellValue = worksheet.Cells[row, columnTitles["Street"]].Text;
                        if (cellValue.Contains(report.Address))
                        {
                            reportAddressRow = row;
                        }
                    }

                    // IF ADDRESS NOT FOUND, DISPLAY MSG TO USER
                    if (reportAddressRow == 0)
                    {
                        OutputLine(failIcon + " [" + report.Address + "] was not found in K1 :(");
                    } else
                    {
                        List<string> successfulOutputs = new List<string>();
                        List<string> erroredOutputs = new List<string>();

                        // ELSE APPEND EXPENSE DATA TO K1
                        foreach (var expense in report.Expenses)
                        {
                            try
                            {
                                int expenseColumn = columnTitles[expense.Key];
                                worksheet.Cells[reportAddressRow, expenseColumn].Value = expense.Value;
                                successfulOutputs.Add("\t" + successIcon + " " + expense.Key + "[" + expense.Value + "]");

                            } catch(KeyNotFoundException)
                            {
                                erroredOutputs.Add("\t" + failIcon + " " + expense.Key + "[" + expense.Value + "]" + " - Column was not found in K1 so it wasn't added.");
                            }                            
                        }

                        //output successful writes first, then errored writes
                        foreach (string successMessage in successfulOutputs)
                        {
                            OutputLine(successMessage);
                        }
                        foreach (string errorMessage in erroredOutputs)
                        {
                            OutputLine(errorMessage);
                        }

                        SaveFile(existingFile);
                        OutputLine("Success!");
                        OutputLine();
                    }
                }
            }
        }

        public static void SaveFile(FileInfo existingFile)
        {
            using (var package = new ExcelPackage(existingFile))
            {
                // Save the updated Excel file
                try
                {
                    package.Save();
                }
                catch (InvalidOperationException)
                {
                    throw new Exception("Please make sure it's not open and run the program again.");
                }
            }
        }

        public static void OuputErrorMessage(string errorMessage)
        {
            OutputLine("   _________");
            OutputLine("  /         \\");
            OutputLine(" |   Error   |");
            OutputLine("  \\_________/");
            OutputLine();
            OutputLine($"   {errorMessage}");
        }

        public static void OutputLine(string message = "")
        {
            Console.WriteLine(message);
            //TODO: also log this to an ouput file in a new directory ./Logs
        }

        public static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {

            OuputErrorMessage("A system error happened - " + ((Exception)e.ExceptionObject).Message);
            OutputLine();
            OutputLine("==============================");
            OutputLine("Full Stack Trace:");
            OutputLine(e.ExceptionObject.ToString());
            OutputLine("==============================");
            OutputLine("Press Enter to exit application.");
            Console.ReadLine();
            Environment.Exit(1);
        }
    }

    public class InputExcelReport
    {
        public FileInfo FileInfo { get; set; }
        public Dictionary<string, double> Expenses { get; set; } //Key is Title, Value will be the value
        public string Address { get; set; }
    }
}