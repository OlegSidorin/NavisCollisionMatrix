using System;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "pos3f")]
    public class Pos3f
    {

        [XmlAttribute(AttributeName = "x")]
        public double X { get; set; }

        [XmlAttribute(AttributeName = "y")]
        public double Y { get; set; }

        [XmlAttribute(AttributeName = "z")]
        public double Z { get; set; }
    }


}
