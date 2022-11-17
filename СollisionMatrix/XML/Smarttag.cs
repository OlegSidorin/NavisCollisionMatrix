﻿using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "smarttag")]
    public class Smarttag
    {

        [XmlElement(ElementName = "name")]
        public string Name { get; set; }

        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
    }


}
