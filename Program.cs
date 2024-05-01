using System;
using System.Collections.Generic;
using System.IO;
using iText.Kernel.Pdf;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using iText.Kernel.Pdf.Canvas.Parser;

namespace PdfEmailExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input PDF file
            string pdfFilePath = "input.pdf";

            // Path to the output Excel file
            string excelFilePath = "output1.xlsx";

            // Extract text from PDF and find email addresses
            List<string> emails = ExtractEmailsFromPdf(pdfFilePath);

            // Write emails to Excel
            WriteEmailsToExcel(emails, excelFilePath);

            Console.WriteLine("Email extraction completed. Press any key to exit.");
            Console.ReadKey();
        }

        static List<string> ExtractEmailsFromPdf(string pdfFilePath)
        {
            List<string> emails = new List<string>();

            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                PdfDocument pdfDoc = new PdfDocument(reader);

                for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
                {
                    string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNum));

                    // Use regular expression to find email addresses
                    MatchCollection matches = Regex.Matches(text, @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b");
                    foreach (Match match in matches)
                    {
                        emails.Add(match.Value);
                    }
                }
            }

            return emails;
        }

        static void WriteEmailsToExcel(List<string> emails, string excelFilePath)
        {
            // Set license context to NonCommercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo newFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Emails");

                for (int i = 0; i < emails.Count; i++)
                {
                    worksheet.Cells[i + 1, 1].Value = emails[i];
                }

                package.Save();
            }
        }
    }
}
