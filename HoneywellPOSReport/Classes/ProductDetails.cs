using Newtonsoft.Json;

namespace HoneywellPOSReport
{
    public class ProductDetails {

        [JsonProperty(PropertyName = "PART NAME")]
        public string       PartName { get; set; }

        public int?         QTY { get; set; }

        [JsonProperty(PropertyName = "Distributor Ref No")]
        public int?         DistributerRefNumber { get; set; }

        [JsonProperty(PropertyName = "CUSTOMER NAME")]
        public string       CustomerName { get; set; }

        public string       City { get; set; }

        public string       State { get; set; }

        public string       Country { get; set; }

        public string       Zip { get; set; }

        [JsonProperty(PropertyName = "DATE SOLD")]
        public string       DateSold { get; set; }

        [JsonProperty(PropertyName = "SIC")]
        public string       Sic { get; set; }
    }
}
