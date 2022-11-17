using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Xml
    {
        [JsonProperty("@version")]
        public string version { get; set; }

        [JsonProperty("@encoding")]
        public string encoding { get; set; }
    }


}
