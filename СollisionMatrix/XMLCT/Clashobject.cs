using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashobject")]
    public class Clashobject
    {

        [XmlElement(ElementName = "objectattribute")]
        public Objectattribute Objectattribute { get; set; }

        [XmlElement(ElementName = "layer")]
        public string Layer { get; set; }

        [XmlElement(ElementName = "pathlink")]
        public Pathlink Pathlink { get; set; }

        [XmlElement(ElementName = "smarttags")]
        public Smarttags Smarttags { get; set; }
    }


}
