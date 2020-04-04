using LiteDB;
using System;
using System.Globalization;

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
            desc = desc.Replace("HONEYWELL-IND ", string.Empty);
            string[] nums = desc.Split("-");

            desc = string.Empty;

            foreach (var item in nums)
            {

                desc += $"{item.TrimStart()}-";
            }

            return desc.Split(" ")[0].TrimEnd('-');
        }

        /// <summary>
        /// GetAbbreviatedFromFullName
        /// </summary>
        /// <param name="fullname"></param>
        /// <returns></returns>
        public static string GetAbbreviatedFromFullName(string fullname)
        {
            string[] names = DateTimeFormatInfo.CurrentInfo.MonthNames;
            foreach (var item in names)
            {
                if (item == fullname)
                {
                    return DateTime.ParseExact(item, "MMMM", CultureInfo.CurrentCulture).ToString("MMM").ToUpper();
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// AddSicValue
        /// </summary>
        /// <param name="customerName"></param>
        /// <returns></returns>
        public static string AddSicValue(string customerName)
        {

            //using (StreamReader r = new StreamReader(@"C:\HoneywellPOSReport\HoneywellPOSReport\HoneywellPOSReport\CustomerSIC.json"))
            //{
            //string json = r.ReadToEnd();
            //List<CustomerSic> items = JsonConvert.DeserializeObject<List<CustomerSic>>(json);

            string dataFile = $"{Environment.CurrentDirectory}\\Data\\{Constants.DatabaseName}";

            using (var db = new LiteDatabase(dataFile))
            {
                var col = db.GetCollection<CustomerSic>(Constants.Tables.CustomerSicCodes);

                var results = col.Query()
                    .Where(x => x.CustomerName.ToUpper() == customerName.ToUpper())
                    .Select(x => new { x.CustomerName, x.SIC })
                    .SingleOrDefault();

                if (!Equals(results, null))
                {
                    return results.SIC;
                }
                else
                {
                    Console.Write($"Please enter SIC value for [{customerName}] >> ");
                    string sicValue = Console.ReadLine();

                    col.Insert(

                        new CustomerSic
                        {
                            CustomerName = customerName,
                            SIC = sicValue
                        }
                    );
                    return sicValue;
                }

            }

            //}

        }


    }
}
