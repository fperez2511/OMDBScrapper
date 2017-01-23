using OpenQA.Selenium;
using OpenQA.Selenium.PhantomJS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;


namespace DataCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            string str;//Declaring and intialising Variables
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"Path to the Input File", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);// The Excel Sheet to read from usually the file with IMDB ID's / Names
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;//No of Rows	
            cl = range.Columns.Count; // No of Columns

            List<String> id = new List<string>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    id.Add(str);
                }
            }

            xlWorkBook.Close(true, null, null); // Closing the Excel Sheet
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            for (int k = 0; k < id.Count; k++)
            {
                //IWebDriver driver = new ChromeDriver(@"Path to the Chrome Driver Location"); // Using Chrome Driver
                IWebDriver driver = new PhantomJSDriver(); // Using PhantomJS Driver
                //Navigate to required page
                driver.Navigate().GoToUrl("http://www.omdbapi.com/?i=" + id[k] + "&plot=full&r=json"); // Navigate to the required page

                //Append the entire page source into a string

                String source = driver.PageSource;
                string path = @"Path to the Output File";
                // This text is always added, making the file longer over time
                // if it is not deleted.
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(StripTagsRegex(source));// Writing to the output Text file
                }
                Console.WriteLine(StripTagsRegex(source));// See whats being Scraped 
                //Close the browser
                driver.Close();
                driver.Quit(); // Close and quit the browser
            }
        }
        public static string StripTagsRegex(string source) // A function to remove the HTML tags in the text
        {
            return Regex.Replace(source, "<.*?>", string.Empty);
        }
    }
}
