using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Clashtest
    {
        [JsonProperty("@name")]
        public string Name { get; set; }

        [JsonProperty("@test_type")]
        public string TestType { get; set; }

        [JsonProperty("@status")]
        public string Status { get; set; }

        [JsonProperty("@tolerance")]
        public string Tolerance { get; set; }

        [JsonProperty("@merge_composites")]
        public string MergeComposites { get; set; }
        public Linkage linkage { get; set; }
        public Left left { get; set; }
        public Right right { get; set; }
        public object rules { get; set; }
    }


}
