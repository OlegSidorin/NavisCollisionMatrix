using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Summary
    {
        [JsonProperty("@total")]
        public string total { get; set; }

        [JsonProperty("@new")]
        public string @new { get; set; }

        [JsonProperty("@active")]
        public string active { get; set; }

        [JsonProperty("@reviewed")]
        public string reviewed { get; set; }

        [JsonProperty("@approved")]
        public string approved { get; set; }

        [JsonProperty("@resolved")]
        public string resolved { get; set; }
        public string testtype { get; set; }
        public string teststatus { get; set; }
    }


}
