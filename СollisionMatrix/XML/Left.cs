﻿using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "left")]
    public class Left
    {

        [XmlElement(ElementName = "clashselection")]
        public Clashselection Clashselection { get; set; }
    }


}
