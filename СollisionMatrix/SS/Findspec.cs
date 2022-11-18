using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Findspec
    {
        [JsonProperty("@mode")]
        public string Mode { get; set; }

        [JsonProperty("@disjoint")]
        public string Disjoint { get; set; }
        public Conditions conditions { get; set; }
        public string locator { get; set; }
    }


}
