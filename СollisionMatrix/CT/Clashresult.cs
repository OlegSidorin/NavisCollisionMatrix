using Newtonsoft.Json;

namespace СollisionMatrix.CT
{
    public class Clashresult
    {
        [JsonProperty("@name")]
        public string name { get; set; }

        [JsonProperty("@guid")]
        public string guid { get; set; }

        [JsonProperty("@status")]
        public string status { get; set; }

        [JsonProperty("@distance")]
        public string distance { get; set; }
        public string description { get; set; }
        public string resultstatus { get; set; }
        public Clashpoint clashpoint { get; set; }
        public string gridlocation { get; set; }
        public Createddate createddate { get; set; }
        public Clashobjects clashobjects { get; set; }
        public Approveddate approveddate { get; set; }
        public string approvedby { get; set; }
    }


}
