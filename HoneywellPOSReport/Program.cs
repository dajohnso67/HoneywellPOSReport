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

            string input = $"{currentDirectory}\\{sourceDirectory}\\";

            DirectoryInfo dirInfo = new DirectoryInfo(input);
            FileInfo[] files = dirInfo.GetFiles("*.csv*");
            FileInfo file = (files.Length > 0) ? files[0] : null;

            if (!Object.Equals(file, null))
            {
                using (var reader = new StreamReader(file.FullName)) 
                {
                    using CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture);
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
            }
            else 
            {
                Console.WriteLine($"Source CSV file does not exist in the {sourceDirectory} folder");
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
                        CustomerName = csvItem.CustomerName,
                        DateSold = Convert.ToDateTime(csvItem.ShipRecDate.Value.ToShortDateString()),
                        DistributerRefNumber = 291375,
                        PartName = csvItem.Description,
                        QTY = csvItem.ShipQty,
                        Sic = Utilities.AddSicValue(csvItem.CustomerName),
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
            string FileName = $"HI POS {Utilities.GetAbbreviatedFromFullName(DateTime.Now.ToString("MMMM"))} {DateTime.Now.Year}.xlsx";
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