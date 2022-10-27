using Newtonsoft.Json;


namespace СollisionMatrix.JS
{
    public class Selectionset
    {
        [JsonProperty("@name")]
        public string Name { get; set; }

        [JsonProperty("@guid")]
        public string Guid { get; set; }
        public Findspec findspec { get; set; }
    }


}
