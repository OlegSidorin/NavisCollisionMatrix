using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashselection")]
    public class Clashselection
    {

        [XmlElement(ElementName = "locator")]
        public string Locator { get; set; }

        [XmlAttribute(AttributeName = "selfintersect")]
        public int Selfintersect { get; set; }

        [XmlAttribute(AttributeName = "primtypes")]
        public int Primtypes { get; set; }

        [XmlText]
        public string Text { get; set; }
    }


}
