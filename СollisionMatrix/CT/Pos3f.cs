using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Pos3f
    {
        [JsonProperty("@x")]
        public string x { get; set; }

        [JsonProperty("@y")]
        public string y { get; set; }

        [JsonProperty("@z")]
        public string z { get; set; }
    }


}
