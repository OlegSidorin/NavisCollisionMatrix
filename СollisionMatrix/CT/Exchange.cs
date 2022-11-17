using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Exchange
    {
        [JsonProperty("@xmlns:xsi")]
        public string xmlnsxsi { get; set; }

        [JsonProperty("@units")]
        public string units { get; set; }

        [JsonProperty("@filename")]
        public string filename { get; set; }

        [JsonProperty("@filepath")]
        public string filepath { get; set; }
        public Batchtest batchtest { get; set; }
    }


}
