using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Condition
    {
        [JsonProperty("@test")]
        public string Test { get; set; }

        [JsonProperty("@flags")]
        public string Flags { get; set; }
        public Category category { get; set; }
        public Property property { get; set; }
        public Value value { get; set; }
    }


}
