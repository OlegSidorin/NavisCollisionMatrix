using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Batchtest
    {
        [JsonProperty("@name")]
        public string name { get; set; }

        [JsonProperty("@internal_name")]
        public string internal_name { get; set; }

        [JsonProperty("@units")]
        public string units { get; set; }
        public Clashtests clashtests { get; set; }
        public SS.Selectionsets selectionsets { get; set; }
    }

}
