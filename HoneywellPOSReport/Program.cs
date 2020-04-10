using ClosedXML.Excel;
using CsvHelper;
using LiteDB;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace HoneywellPOSReport
{
    public class Program
    {
        static string currentDirectory;
        static string destinationDirectory;
        static string sourceDirectory;
        static string archiveDirectory;
        static int distributerRefNumber;

        static string fileMonth;

        //static string seedDataDirectory;

        public static void Main(string[] args)
        {
            currentDirectory = Directory.GetCurrentDirectory();

            var builder = new ConfigurationBuilder()
                .SetBasePath(currentDirectory)
                .AddJsonFile("appsettings.json");

            var config = builder.Build();

            destinationDirectory = config.GetSection("destinationDirectory").Get<string>();
            sourceDirectory = config.GetSection("sourceDirectory").Get<string>();
            archiveDirectory = config.GetSection("archiveDirectory").Get<string>();
            distributerRefNumber = config.GetSection("distributerRefNumber").Get<int>();

            // use to insert seed data 
            //seedDataDirectory = config.GetSection("seedDataDirectory").Get<string>();
            //string input2 = $"{currentDirectory}\\{seedDataDirectory}\\";

            //DirectoryInfo dirInfo2 = new DirectoryInfo(input2);
            //FileInfo[] files2 = dirInfo2.GetFiles("*.csv*");

            //foreach (FileInfo f in files2)
            //{
            //    using (var reader = new StreamReader(f.FullName))
            //    {

            //        using CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            //        csv.Configuration.RegisterClassMap<GTHMap2>();

            //        List<CsvColumns2> csvFile = csv.GetRecords<CsvColumns2>().ToList();

            //        foreach (var item in csvFile)
            //        {
            //            if (!string.IsNullOrEmpty(item.CustomerName)) {
            //                Utilities.InsertSicValue(item.CustomerName.Trim(), item.SIC.Trim());
            //            }
            //        }
            //    }
            //}

            //return;

            string input = $"{currentDirectory}\\{sourceDirectory}\\";

            DirectoryInfo dirInfo = new DirectoryInfo(input);
            FileInfo[] files = dirInfo.GetFiles("*.csv*");

            if (files.Length > 0)
            {
                foreach (var file in files)
                {

                    if (!Object.Equals(file, null))
                    {
                        string splMth = file.Name.Split(".")[0];
                        fileMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(Convert.ToInt32(splMth.Substring(splMth.Length - 2).Trim())).ToUpper();

                        using (var reader = new StreamReader(file.FullName))
                        {
                            using CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture);
                            csv.Configuration.IgnoreBlankLines = true;
                            csv.Configuration.IgnoreQuotes = true;
                            csv.Configuration.BadDataFound = x =>
                            {
                                x.Record = null;
                            };

                            csv.Configuration.RegisterClassMap<GTHMap>();

                            List<CsvColumns> csvFile = csv.GetRecords<CsvColumns>().Where(c => (c.State.ToUpper() == "CA" || c.State.ToUpper() == "NV") && c.ShipQty > 0).ToList();
                            csvFile.ForEach(c => c.Description = Utilities.CleanUpDescription(c.Description));
                            WriteExcelFile(csvFile);
                        }

                        string dest = $"{currentDirectory}\\{archiveDirectory}\\";

                        if (!Directory.Exists(dest))
                        {
                            Directory.CreateDirectory(dest);
                        }

                        string movedFile = $"{dest}{file.Name}";

                        if (File.Exists(movedFile))
                        {
                            File.Delete(movedFile);
                        }

                        File.Move(file.FullName, movedFile);

                        Console.WriteLine($"\r\nDONE PROCESSING [{file.Name}]\r\n");

                    }
                    else
                    {
                        Console.WriteLine($"Source CSV file does not exist in the [{sourceDirectory}] folder");
                    }

                }
            }
            else {

                Console.WriteLine($"The [{sourceDirectory}] directory does not contain any .CSV data files");
            }

            
            
            Console.ReadKey();
        }

        private static void WriteExcelFile(List<CsvColumns> csvFiles)
        {
            List<ProductDetails> products = new List<ProductDetails>();

            ConsoleTable ct = new ConsoleTable
            {
                TextAlignment = ConsoleTable.AlignText.ALIGN_LEFT
            };

            ct.SetHeaders(new string[] {"", "PART NAME", "CUSTOMER NAME" });

            int rowCount = 1;
            csvFiles.ForEach(csvItem =>
            {
                products.Add(

                    new ProductDetails
                    {
                        City = csvItem.City,
                        Country = "USA",
                        CustomerName = csvItem.CustomerName.Trim(),
                        DateSold = Convert.ToDateTime(csvItem.ShipRecDate.Value.ToShortDateString()),
                        DistributerRefNumber = distributerRefNumber,
                        PartName = csvItem.Description,
                        QTY = csvItem.ShipQty,
                        Sic = Utilities.AddSicValue(csvItem.CustomerName.Trim()),
                        State = csvItem.State,
                        Zip = csvItem.ZipCode,

                    });

                ct.AddRow(new List<string> { rowCount.ToString(), csvItem.Description, csvItem.CustomerName });

                rowCount++;
            });

            ct.PrintTable();

            Console.WriteLine($"\r\n{csvFiles.Count} rows mapped");

            products = products.OrderByDescending(c => c.DateSold).ToList();

            string productsJSON = JsonConvert.SerializeObject(products);
            DataTable table = (DataTable)JsonConvert.DeserializeObject(productsJSON, (typeof(DataTable)));
            string FileName = $"HI POS {fileMonth} {DateTime.Now.Year}.xlsx";
            string Output = $"{currentDirectory}\\{destinationDirectory}\\{FileName}";
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(table, Constants.WorksheetName);

            ws.Rows(1, 1).Height = 30;
            ws.SheetView.FreezeRows(1);

            for (int i = 1; i < ws.Columns().ToArray().Length + 1; i++) 
            {
                ws.Cell(1, i).Style.Fill.BackgroundColor = XLColor.PastelYellow;
                ws.Cell(1, i).Style.Font.FontColor = XLColor.SmokyBlack;
                ws.Cell(1, i).Style.Font.Bold = true;
            }

            wb.SaveAs(Output);
        }
    }
}