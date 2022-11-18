using Newtonsoft.Json;

namespace СollisionMatrix.SS
{
    public class Root
    {
        [JsonProperty("?xml")]
        public Xml Xml { get; set; }
        public Exchange exchange { get; set; }
    }
}
