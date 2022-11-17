using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "linkage")]
    public class Linkage
    {

        [XmlAttribute(AttributeName = "mode")]
        public string Mode { get; set; }
    }


}
