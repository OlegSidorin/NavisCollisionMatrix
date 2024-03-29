﻿using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "batchtest")]
    public class Batchtest
    {

        [XmlElement(ElementName = "clashtests")]
        public Clashtests Clashtests { get; set; }

        [XmlElement(ElementName = "selectionsets")]
        public object Selectionsets { get; set; }

        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [XmlAttribute(AttributeName = "internal_name")]
        public string InternalName { get; set; }

        [XmlAttribute(AttributeName = "units")]
        public string Units { get; set; }

        [XmlText]
        public string Text { get; set; }
    }


}
