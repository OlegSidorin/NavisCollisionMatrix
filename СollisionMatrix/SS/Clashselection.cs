﻿using Newtonsoft.Json;


namespace СollisionMatrix.SS
{
    public class Clashselection
    {
        [JsonProperty("@selfintersect")]
        public string Selfintersect { get; set; }

        [JsonProperty("@primtypes")]
        public string Primtypes { get; set; }
        public string locator { get; set; }
    }


}
