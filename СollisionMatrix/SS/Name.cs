using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Name
    {
        [JsonProperty("@internal")]
        public string Internal { get; set; }

        [JsonProperty("#text")]
        public string Text { get; set; }
    }


}
