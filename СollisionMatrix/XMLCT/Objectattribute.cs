﻿using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "objectattribute")]
    public class Objectattribute
    {

        [XmlElement(ElementName = "name")]
        public string Name { get; set; }

        [XmlElement(ElementName = "value")]
        public int Value { get; set; }
    }


}
