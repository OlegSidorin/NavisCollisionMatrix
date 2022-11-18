using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Batchtest
    {
        [JsonProperty("@name")]
        public string Name { get; set; }

        [JsonProperty("@internal_name")]
        public string InternalName { get; set; }

        [JsonProperty("@units")]
        public string Units { get; set; }
        public Clashtests clashtests { get; set; }
        public Selectionsets selectionsets { get; set; }
    }


}
