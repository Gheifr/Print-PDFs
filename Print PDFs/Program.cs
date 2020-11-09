using System;
using System.IO;
using System.Windows.Forms;
using System.Configuration;
using Spire.Pdf;

namespace Print_PDFs
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string folderPath;

            var appSet = ConfigurationSettings.AppSettings;

            int startPage = string.IsNullOrEmpty(appSet["StartPage"]) ? 0 : int.Parse(appSet["StartPage"]);
            int endPage = string.IsNullOrEmpty(appSet["EndPage"]) ? 0 : int.Parse(appSet["EndPage"]);
            bool onlyLastPage = string.IsNullOrEmpty(appSet["LastPageOnly"]) ? true : bool.Parse(appSet["LastPageOnly"]);
            bool wholeDocument = string.IsNullOrEmpty(appSet["Whole"]) ? false : bool.Parse(appSet["Whole"]);
            int pages;

            Console.WriteLine("Select the folder with files.");

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult dialogResult = fbd.ShowDialog();
                folderPath = fbd.SelectedPath;
            }

            foreach (var file in Directory.GetFiles(folderPath))
            {
                FileInfo fi = new FileInfo(file);

                if (fi.Extension == ".pdf")
                {
                    PdfDocument pdfDocument = new PdfDocument(file);

                    pages = pdfDocument.Pages.Count;

                    if (onlyLastPage)
                    {
                        pdfDocument.PrintSettings.SelectPageRange(pages, pages);
                        pdfDocument.Print();
                        Console.WriteLine($"Page {pages} of document {file} was sent to printer.");
                    }
                    else
                    {
                        if (wholeDocument)
                        {
                            pdfDocument.PrintSettings.SelectPageRange(1, pages);
                            pdfDocument.Print();
                            Console.WriteLine($"Document {file} was sent to printer.");
                            return;
                        }

                        if (startPage == 0)
                        {
                            Console.WriteLine("Start page parameter is not set or incorrect");
                            return;
                        }

                        if (endPage == 0)
                        {
                            Console.WriteLine("End page parameter is not set or incorrect");
                            return;
                        }
                        else
                        {
                            if (endPage > pages)
                            {
                                endPage = pages;
                            }
                            pdfDocument.PrintSettings.SelectPageRange(startPage, endPage);
                            pdfDocument.Print();

                            Console.WriteLine($"Document {file} was sent to printer from page {startPage} to {endPage}.");
                        }
                    }
                }
            }
            Console.WriteLine("Finished!");
            Console.ReadKey();
        }
    }
}
