using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Clashtest
    {
        [JsonProperty("@name")]
        public string name { get; set; }

        [JsonProperty("@test_type")]
        public string test_type { get; set; }

        [JsonProperty("@status")]
        public string status { get; set; }

        [JsonProperty("@tolerance")]
        public string tolerance { get; set; }

        [JsonProperty("@merge_composites")]
        public string merge_composites { get; set; }
        public Linkage linkage { get; set; }
        public Left left { get; set; }
        public Right right { get; set; }
        public object rules { get; set; }
        public Summary summary { get; set; }
        public Clashresults clashresults { get; set; }
    }


}
