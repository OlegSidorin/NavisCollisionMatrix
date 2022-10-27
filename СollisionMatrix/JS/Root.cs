using Newtonsoft.Json;

namespace СollisionMatrix.JS
{
    public class Root
    {
        [JsonProperty("?xml")]
        public Xml Xml { get; set; }
        public Exchange exchange { get; set; }
    }
}
