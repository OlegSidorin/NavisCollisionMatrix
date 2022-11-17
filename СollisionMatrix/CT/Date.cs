using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Date
    {
        [JsonProperty("@year")]
        public string year { get; set; }

        [JsonProperty("@month")]
        public string month { get; set; }

        [JsonProperty("@day")]
        public string day { get; set; }

        [JsonProperty("@hour")]
        public string hour { get; set; }

        [JsonProperty("@minute")]
        public string minute { get; set; }

        [JsonProperty("@second")]
        public string second { get; set; }
    }


}
