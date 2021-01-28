
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace FragranceInfo
{

    public class FragranticaCrawler
    {
        const int MAX_THREAD = 2;
        static int NUM_OF_THREADS = 0;
        static int STARTING_ROW = 1;
        static int STARTING_COLUMN = 1;
        static ExcelWorksheet Worksheet;
        static ExcelPackage Package;
        static string URL = "URL";
        static string Name = "Name";
        static bool ALWAYS_PROCESS = false; // crawl the website again even if the fragrance's info is already populated
        static bool HEADLESS = true;
        public static void OpenExcelFile()
        {
            string path = @"C:\Users\khoan\Desktop\Fragance.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            Package = new ExcelPackage(fileInfo);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Worksheet = Package.Workbook.Worksheets.FirstOrDefault();
        }
        public static void ProcessURLs()
        {

            OpenExcelFile();

            IList<Tuple<string, int>> webs = GetURLs();
            int total = webs.Count;
            IList<Task<Fragrance>> tasks = new List<Task<Fragrance>>();
            int count = 0;
            while (tasks.Count < total)
            {
                if (NUM_OF_THREADS < MAX_THREAD)
                {
                    string web = webs[0].Item1;
                    int row = webs[0].Item2;
                    webs.RemoveAt(0);
                    Interlocked.Increment(ref NUM_OF_THREADS);
                    tasks.Add(Task.Run(() => GetFranganceInfo(web, row)));
                    count++;
                }
            }

            Task.WaitAll(tasks.ToArray());

            IList<Fragrance> fragrances = tasks.Select(x => x.Result).ToList();
            WriteExcelFile(fragrances);
        }

        public static Fragrance GetFranganceInfo(string url, int row)
        {
            Console.WriteLine($"Processing {url}");
            int attempts = 4;

            ChromeOptions chromeOptions = new ChromeOptions();
            if (HEADLESS)
            {
                chromeOptions.AddArguments("headless");
            }

            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            IWebDriver driver = new ChromeDriver(chromeDriverService, chromeOptions);

            driver.Navigate().GoToUrl(url);

            Fragrance fragance = new Fragrance(driver, url);
            fragance.Row = row;

            // Sometimes the website is not fully loaded, wait if it fails to grab any attribute. 10 times max. 
            while (attempts > 0)
            {
                try
                {
                    Thread.Sleep(500);
                    Console.WriteLine($"{attempts} attempts left for {url}");
                    fragance.GetAllAttributes();
                    break;
                }
                catch (Exception e)
                {
                    Thread.Sleep(1000);
                    attempts -= 1;
                }
            }
            Interlocked.Decrement(ref NUM_OF_THREADS);
            driver.Close();
            Console.WriteLine($"Done - {fragance.Name}. Is Processed? {fragance.IsProcessed}");
            return fragance;
        }

        public static int FindColumnIndex(string name)
        {
            int lastNonEmptyColumn = GetLastNonEmptyColumnIndex();
            int column = STARTING_COLUMN;

            while (column <= lastNonEmptyColumn)
            {
                if (Worksheet.Cells[STARTING_ROW, column].Value?.ToString() == name)
                {
                    return column;
                }

                column++;
            }

            return 0;
        }

        public static IList<Tuple<string, int>> GetURLs()
        {
            var urlColumn = FindColumnIndex(URL);
            var nameColumn = FindColumnIndex(Name);
            int row = STARTING_ROW + 1;
            IList<Tuple<string, int>> urls = new List<Tuple<string, int>>();
            while (true)
            {
                ExcelRange urlCell = Worksheet.Cells[row, urlColumn];
                var url = urlCell.Value;

                ExcelRange nameCell = Worksheet.Cells[row, nameColumn];
                var name = nameCell.Value;

                if (url == null)
                {
                    break;
                }

                if (ALWAYS_PROCESS || name == null)
                {
                    urls.Add(new Tuple<string, int>(url.ToString(), row));
                }
                row++;
            }
            return urls;
        }

        public static void WriteExcelFile(IList<Fragrance> fragrances)
        {
            foreach (Fragrance fragrance in fragrances.Where(x => x.IsProcessed))
            {
                if (!fragrance.IsProcessed)
                {
                    continue;
                }

                IList<FieldInfo> fields = fragrance.GetType().GetFields().ToList();

                foreach (FieldInfo field in fields)
                {
                    DescriptionAttribute descriptionAttribute = (DescriptionAttribute)field.GetCustomAttribute(typeof(DescriptionAttribute), false);
                    ReadOnlyAttribute readonlyAttribute = (ReadOnlyAttribute)field.GetCustomAttribute(typeof(ReadOnlyAttribute), false);

                    if (descriptionAttribute == null)
                    {
                        continue;
                    }

                    int column = FindColumnIndex(descriptionAttribute.Description);

                    if (column == 0)
                    {
                        column = InsertEmptyColumn();
                        Worksheet.Cells[STARTING_ROW, column].Value = descriptionAttribute.Description;
                    }
                    else if (readonlyAttribute != null && readonlyAttribute.IsReadOnly && Worksheet.Cells[fragrance.Row, column].Value != null) // If it's readonly and the cell already has a value, skip
                    {
                        continue;
                    }

                    var value = field.GetValue(fragrance);

                    if (field.FieldType == typeof(decimal))
                    {
                        value = Math.Round((decimal)value, 2);
                    }

                    Worksheet.Cells[fragrance.Row, column].Value = value;
                    Worksheet.Cells[Worksheet.Dimension.Address].AutoFitColumns();
                    Package.Save();
                }
            }           
        }

        private static int InsertEmptyColumn()
        {
            int emptyColumn = FindColumnIndex(null);

            if (emptyColumn > 0)
            {
                return emptyColumn;
            }

            emptyColumn = GetLastNonEmptyColumnIndex() + 1;
            Worksheet.InsertColumn(emptyColumn, 1);

            return emptyColumn;
        }

        private static int GetLastNonEmptyColumnIndex()
        {
            return Worksheet.Cells["1:1"].Last().End.Column;
        }
    }
}
