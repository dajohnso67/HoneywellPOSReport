using CsvHelper.Configuration;
using System;

namespace HoneywellPOSReport
{
    public class CsvColumns
    {
        public string Description { get; set; }
        public int? ShipQty { get; set; }
        public string CustomerName { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public DateTime? ShipRecDate { get; set; }
        public string PriceLine { get; set; }
    }

    public class GTHMap : ClassMap<CsvColumns>
    {
        public GTHMap()
        {
            Map(m => m.Description).Name("Description");
            Map(m => m.ShipQty).Name("Ship Qty");
            Map(m => m.CustomerName).Name("Customer Name");
            Map(m => m.City).Name("City");
            Map(m => m.State).Name("State");
            Map(m => m.ZipCode).Name("Zip Code");
            Map(m => m.ShipRecDate).Name("Ship/Rec. Date");
            Map(m => m.PriceLine).Name("Price Line");
        }
    }
}
