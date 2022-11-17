using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Clashselection
    {
        [JsonProperty("@selfintersect")]
        public string selfintersect { get; set; }

        [JsonProperty("@primtypes")]
        public string primtypes { get; set; }
        public string locator { get; set; }
    }


}
