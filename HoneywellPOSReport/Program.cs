using ClosedXML.Excel;
using CsvHelper;
using LiteDB;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace HoneywellPOSReport
{
    class Program
    {
        static void Main(string[] args)
        {
            List<CsvColumns> csvFile = null;

            string input = $"{Environment.CurrentDirectory}\\Source_csv\\";

            DirectoryInfo dirInfo = new DirectoryInfo(input);
            FileInfo file = dirInfo.GetFiles("*.*")[0];

            if (file.Extension.Contains("csv", StringComparison.OrdinalIgnoreCase))
            {
                using (var reader = new StreamReader(file.FullName))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Configuration.RegisterClassMap<GTHMap>();
                    csvFile = csv.GetRecords<CsvColumns>().Where(c => (c.State.ToUpper() == "CA" || c.State.ToUpper() == "NV") && c.ShipQty > 0).ToList();

                    csvFile.ForEach(c => c.Description = Utilities.CleanUpDescription(c.Description));
                }

                WriteExcelFile(csvFile);
            }
            else 
            {
                Console.WriteLine("File type must be CSV");
            }

            Console.ReadKey();
        }

        static void WriteExcelFile(List<CsvColumns> csvFiles)
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
                        DateSold = csvItem.ShipRecDate.Value.ToShortDateString(),
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

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(products), (typeof(DataTable)));
            string FileName = $"HI POS {Utilities.GetAbbreviatedFromFullName(DateTime.Now.ToString("MMMM"))} {DateTime.Now.Year}.xlsx";
            string Output = $"{Environment.CurrentDirectory}\\Destination_xlsx\\{FileName}";
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(table, Constants.WorksheetName);

            ws.Rows(1, 1).Height = 30;
            ws.SheetView.FreezeRows(1);
    
            for (int i = 1; i < ws.Columns().ToArray().Length + 1; i++) 
            {
                ws.Cell(1, i).Style.Fill.BackgroundColor = XLColor.PastelYellow;
                ws.Cell(1, i).Style.Font.FontColor = XLColor.RichBlack;
                ws.Cell(1, i).Style.Font.Bold = true;
            }

            wb.SaveAs(Output);
        }
    }
}