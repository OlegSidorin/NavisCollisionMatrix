﻿using Newtonsoft.Json;

namespace СollisionMatrix.SS
{
    public class Data
    {
        [JsonProperty("@type")]
        public string Type { get; set; }

        [JsonProperty("#text")]
        public string Text { get; set; }
    }


}
