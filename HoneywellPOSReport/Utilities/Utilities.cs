using LiteDB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace HoneywellPOSReport
{
    public static class Utilities
    {
        /// <summary>
        /// CleanUpDescription
        /// </summary>
        /// <param name="desc"></param>
        /// <returns></returns>
        public static string CleanUpDescription(string desc)
        {
            List<string> nums = desc.Replace("HONEYWELL-IND ", string.Empty).Split("-").ToList();
            desc = string.Empty;
            nums.ForEach(c => { desc += $"{c.TrimStart()}-"; });
            return desc.Split(" ")[0].TrimEnd('-');
        }

        /// <summary>
        /// GetAbbreviatedFromFullName
        /// </summary>
        /// <param name="fullname"></param>
        /// <returns></returns>
        public static string GetAbbreviatedFromFullName(string fullname)
        {
            return DateTime.ParseExact(fullname, "MMMM", CultureInfo.CurrentCulture)
               .ToString("MMM")
               .ToUpper();
        }

        /// <summary>
        /// AddSicValue
        /// </summary>
        /// <param name="customerName"></param>
        /// <returns></returns>
        public static string AddSicValue(string customerName)
        {
            string dataFile = $"{Environment.CurrentDirectory}\\Data\\{Constants.DatabaseName}";

            using (var db = new LiteDatabase(dataFile))
            {
                var col = db.GetCollection<CustomerSic>(Constants.Tables.CustomerSicCodes);

                var results = col.Query()
                    .Where(x => x.CustomerName.Trim().ToUpper() == customerName.Trim().ToUpper())
                    .Select(x => new { x.CustomerName, x.SIC })
                    .SingleOrDefault();

                if (!Equals(results, null))
                {
                    return results.SIC.Trim();
                }
                else
                {
                    Console.Write($"Please enter SIC value for [{customerName}] >> ");
                    string sicValue = Console.ReadLine();

                    col.Insert(

                        new CustomerSic
                        {
                            CustomerName = customerName.Trim(),
                            SIC = sicValue.Trim()
                        }
                    );
                    return sicValue;
                }

            }
        }

        public static void InsertSicValue(string customerName, string sic)
        {
            string dataFile = $"{Environment.CurrentDirectory}\\Data\\{Constants.DatabaseName}";

            using (var db = new LiteDatabase(dataFile))
            {
                var col = db.GetCollection<CustomerSic>(Constants.Tables.CustomerSicCodes);

                //var results = col.Query().ToList();

                //var query = col.Query().ToList()
                //    .GroupBy(x => x.CustomerName)
                //  .Where(g => g.Count() > 1)
                //  .Select(y => y.First())
                //  .ToList();

                var results = col.Query()
                    .Where(x => x.CustomerName.Trim().ToUpper() == customerName.Trim().ToUpper())
                    .Select(x => new { x.CustomerName, x.SIC }).ToList();

                if (results.Count() > 1)
                {
                    Console.WriteLine(customerName);
                }

                if (results.Count() == 0)
                {
                    col.Insert(
                       new CustomerSic
                       {
                           CustomerName = customerName.Trim(),
                           SIC = sic.Trim()
                       }
                   );
                }
            }
        }
    }
}
