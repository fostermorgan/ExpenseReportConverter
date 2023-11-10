using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using static System.Net.WebRequestMethods;
using System.Runtime.ConstrainedExecution;
using System.Net;
using System.Diagnostics.Metrics;
using System.Xml.Linq;
using System.Reflection.PortableExecutable;
using iText.Kernel.Pdf;
using iText.Forms;
using iText.Forms.Fields;
using static iText.IO.Codec.TiffWriter;

namespace ExpenseReportConverter
{
    internal class Program
    {
        public static int InputsHeaderRow { get; set; }
        public static int OutputMasterK1HeaderRow { get; set; }

        public static List<string> SuccessfullAddressWritesToK1 { get; set; } = new List<string>();
        public static List<string> ErroredAddressWritesToK1 { get; set; } = new List<string>();
        public static char successIcon = '\u2713';
        public static char failIcon = 'x';

        public static string? directoryPath = "";
        public static string dateTimeOnRun = "it shouldnt be named this.txt";

        static void Main(string[] args)
        {
            WriteToPdf();
            dateTimeOnRun = $"File_{DateTime.Now:yyyyMMddHHmmss}.txt";
            //SET UP APPLICATION
            //set license information for the excel library nuget package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //set up exception handling
            System.AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
            
            
            // prompt user for directory of where all the files are
            directoryPath = string.Empty;
            while (string.IsNullOrEmpty(directoryPath))
            {
                OutputLine("Please enter your Base Directory (where your files are stored):");
                directoryPath = Console.ReadLine();

                if (string.IsNullOrEmpty(directoryPath))
                {
                    OutputLine("Base Directory cannot be empty. Please try again.");
                }

                if (!Directory.Exists(directoryPath))
                {
                    OutputLine("The directory does not exist. Please provide a valid directory.");
                    directoryPath = string.Empty;
                }
            }
            
            string? masterK1DocName = string.Empty;
            // prompt user for name of master k - 1 doc
            while (string.IsNullOrEmpty(masterK1DocName))
            {
                //note: could add message of must be xlsx file?
                OutputLine("Please enter the name of your master k1 document ('Master K-1.xlsx' if you hit Enter):");
                masterK1DocName = Console.ReadLine();
                if (string.IsNullOrEmpty(masterK1DocName))
                {
                    masterK1DocName = @"Master K-1";
                }
            }
            if (!masterK1DocName.Contains(".xlsx"))
            {
                masterK1DocName += ".xlsx";
            }

            OutputLine();
            DisplayWelcomeMessage();

            InputsHeaderRow = 5;
            OutputMasterK1HeaderRow = 3;

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
                    throw new Exception("The master K-1 file '" + masterK1DocName + "' was not found.");
                }

                // Check if main output K1 file is open, if it is prompt user to close it.
                SaveFile(masterK1FileInfo);

                // display information user inputted
                OutputLine("Looking for files in directory: " + directoryPath);
                OutputLine("Master K-1 output File: " + masterK1DocName);
                OutputLine("Other assumptions: \n" +
                    "\t1. Header row for each input excel sheet (5). \n" +
                    "\t2. Header row for K1 doc (3)\n" +
                    "\t3. Name of column to search for a match is named 'Street' in K1 sheet\n"); //note could prompt user or have a env.txt on build with defaults.

                OutputLine("=======================");
                OutputLine("\tSTARTING...");
                OutputLine("=======================");
                OutputLine(inputSpreadsheetFiles.Count + " file(s) were found for processing.");

                foreach (FileInfo file in inputSpreadsheetFiles)
                {
                    OutputLine("Starting to process: " + file.Name + "...");
                    // parse address and expenses from input excel sheet
                    InputExcelReport ier = ParseInputExcel(file);

                    // Output ier data to master K-1 sheet.
                    string apendExcelFilePath = masterK1FileInfo.FullName;
                    AppendDataToExcel(apendExcelFilePath, ier);
                    OutputLine();

                    // Output ier data to pdf sheet.

                }
            }
            else
            {
                OutputLine("The specified directory does not exist.");
            }

            OutputLine("=======================");
            OutputLine("\tSummary");
            OutputLine("=======================");

            //print summary of application run.
            OutputLine(successIcon + " Successful writes to K1:");
            OutputLine(string.Join("\n", SuccessfullAddressWritesToK1));
            OutputLine();
            OutputLine(failIcon + " Addresses found, but not written to K1:");
            OutputLine(string.Join("\n", ErroredAddressWritesToK1));
            OutputLine();

            //TODO: specify where output log goes.
            //OutputLine("The output log can be found at " + )

            //prompt user to exit the application
            OutputLine("Press Enter to exit application.");
            Console.ReadLine();
            Environment.Exit(1);

        }

        public static void WriteToPdf()
        {
            //string pdfFilePath = "Path_to_Your_PDF_File.pdf"; // Replace with your PDF file path

            //string fieldName = "YourFieldName"; // Replace with the actual field name in your PDF

            //string fieldValue = "Value to Fill"; // The value you want to set in the field

            //try
            //{
            //    PdfDocument pdfDocument = new PdfDocument(new PdfReader(pdfFilePath), new PdfWriter(pdfFilePath));

            //    PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            //    form.GetField(fieldName).SetValue(fieldValue);

            //    pdfDocument.Close();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("An error occurred: " + ex.Message);
            //}


            string pdfFilePath = @"C:\code\ExpenseReportConverter\Development References\2022 CLA Tax Organizer Zach Zank FINAL -part-2.pdf"; // Replace with your PDF file path
            //try
            //{
            //    using (PdfReader pdfReader = new PdfReader(pdfFilePath))
            //    {
            //        PdfDocument pdfDocument = new PdfDocument(pdfReader, new PdfWriter(pdfFilePath + ".tmp"));
            //        PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);

            //        if (form != null)
            //        {
            //            ICollection<PdfFormField> fields = form.GetAllFormFields().Values;

            //            foreach (PdfFormField field in fields)
            //            {
            //                string fieldName = field.GetFieldName().GetValue();
            //                field.SetValue(fieldName);
            //            }
            //        }
            //        else
            //        {
            //            Console.WriteLine("The document does not contain any form fields.");
            //        }

            //        pdfDocument.Close();
            //    }

            //    // Delete the original file and rename the modified one
            //    System.IO.File.Delete(pdfFilePath);
            //    System.IO.File.Move(pdfFilePath + ".tmp", pdfFilePath);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("An error occurred: " + ex.Message);
            //}


            try
            {
                PdfDocument pdfDocument = new PdfDocument(new PdfWriter(pdfFilePath));

                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
                if (form != null)
                {
                    ICollection<PdfFormField> fields = form.GetAllFormFields().Values;

                    foreach (PdfFormField field in fields)
                    {
                        string fieldName = field.GetFieldName().GetValue();
                        field.SetValue(fieldName);
                    }
                }
                else
                {
                    Console.WriteLine("The document does not contain any form fields.");
                }

                pdfDocument.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }


            //try
            //{
            //    PdfDocument pdfDocument = new PdfDocument(new PdfReader(pdfFilePath));

            //    PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, false);
            //    if (form != null)
            //    {
            //        ICollection<PdfFormField> fields = form.GetAllFormFields().Values;

            //        Console.WriteLine("Form Field Names:");
            //        foreach (PdfFormField field in fields)
            //        {

            //            string fieldName = field.GetFieldName().GetValue();
            //            Console.WriteLine(field.GetFieldName());
            //            form.GetField(fieldName).SetValue(fieldName);
            //        }
            //    }
            //    else
            //    {
            //        Console.WriteLine("The document does not contain any form fields.");
            //    }

            //    pdfDocument.Close();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("An error occurred: " + ex.Message);
            //}
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

        public static void AppendDataToExcel(string excelFilePath, InputExcelReport report)
        {
            // Load the existing Excel file
            FileInfo existingFile = new FileInfo(excelFilePath);

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    // IDENTIFY ALL COLUMNS 
                    Dictionary<string, int> k1ColumnTitles = new Dictionary<string, int>();
                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;

                    for (int column = 1; column <= columnCount; column++)
                    {
                        string cellValue = worksheet.Cells[OutputMasterK1HeaderRow, column].Text.Trim();
                        if (!cellValue.Trim().Equals(""))
                        {
                            k1ColumnTitles.Add(cellValue, column);     
                        }
                    }

                    // IDENTIFY ROW OF ADDRESS TO APPEND TO
                    int reportAddressRow = 0;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var cellValue = worksheet.Cells[row, k1ColumnTitles["Street"]].Text;
                        if (cellValue.Contains(report.Address))
                        {
                            reportAddressRow = row;
                        }
                    }

                    // IF ADDRESS NOT FOUND, DISPLAY MSG TO USER
                    if (reportAddressRow == 0)
                    {
                        OutputLine(failIcon + " [" + report.Address + "] was not found in K1 so nothing was written.");
                        ErroredAddressWritesToK1.Add(report.Address);
                    } else
                    {
                        List<string> successfulOutputs = new List<string>();
                        List<string> erroredOutputs = new List<string>();

                        // ELSE APPEND EXPENSE DATA TO K1
                        foreach (var expense in report.Expenses)
                        {
                            try
                            {
                                int expenseColumn = k1ColumnTitles[expense.Key];
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
                        SuccessfullAddressWritesToK1.Add(report.Address);
                        OutputLine("Success!");
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
                    throw new Exception("Please make sure the K1 file not open and run the program again.");
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
            // WRITE MESSAGE TO CONSOLE
            Console.WriteLine(message);

            // WRITE MESSAGE TO LOG FILE
            if(directoryPath != null)
            {
                string directoryName = Path.Combine(directoryPath, "Logs");

                if (!Directory.Exists(directoryName))
                {
                    Directory.CreateDirectory(directoryName);
                }
                string filePath = Path.Combine(directoryName, dateTimeOnRun);
                //if (!System.IO.File.Exists(filePath))
                //{
                //    // If the file doesn't exist, create a new file with a name based on the current timestamp
                //    filePath = Path.Combine(directoryName, $"File_{DateTime.Now:yyyyMMddHHmmss}.txt");
                //}

                // Append the line to the file
                using (StreamWriter sw = System.IO.File.AppendText(filePath))
                {
                    sw.WriteLine(message);
                }
            }
        }

        public static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {

            OuputErrorMessage("A system error happened - " + ((Exception)e.ExceptionObject).Message);
            OutputLine();
            OutputLine("==============================");
            OutputLine("Full Stack Trace:");
            OutputLine(e.ExceptionObject.ToString());
            OutputLine("==============================");
            OutputLine("Press enter to exit application.");
            Console.ReadLine();
            Environment.Exit(1);
        }

        public static void DisplayWelcomeMessage()
        {
            OutputLine("    _____");
            OutputLine("   /\\\\   \\\\");
            OutputLine("  /  \\\\   \\\\            <=======================================>");
            OutputLine(" /    \\\\___\\\\            WELCOME TO THE EXPENSE REPORT CONVERTER");
            OutputLine("/___________\\\\          <=======================================>");
            OutputLine("|   ______   |");
            OutputLine("|  |      |  |                   - Foster");
            OutputLine("|__|______|__|");
            OutputLine();
        }
    }

    public class InputExcelReport
    {
        public FileInfo FileInfo { get; set; }
        public Dictionary<string, double> Expenses { get; set; } //Key is Title, Value will be the value
        public string Address { get; set; }
    }
}