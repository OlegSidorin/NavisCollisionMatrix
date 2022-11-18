using Newtonsoft.Json;

namespace СollisionMatrix.SS
{
    public class Exchange
    {
        [JsonProperty("@xmlns:xsi")]
        public string XmlnsXsi { get; set; }

        [JsonProperty("@xsi:noNamespaceSchemaLocation")]
        public string XsiNoNamespaceSchemaLocation { get; set; }

        [JsonProperty("@units")]
        public string Units { get; set; }

        [JsonProperty("@filename")]
        public string Filename { get; set; }

        [JsonProperty("@filepath")]
        public string Filepath { get; set; }
        public Batchtest batchtest { get; set; }
    }
}
