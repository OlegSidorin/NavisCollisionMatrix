using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Root
    {
        [JsonProperty("?xml")]
        public Xml xml { get; set; }
        public Exchange exchange { get; set; }
    }


}
