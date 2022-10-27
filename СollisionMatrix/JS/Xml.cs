using Newtonsoft.Json;

namespace СollisionMatrix.JS
{
    public class Xml
    {
        [JsonProperty("@version")]
        public string Version { get; set; }

        [JsonProperty("@encoding")]
        public string Encoding { get; set; }
    }


}
