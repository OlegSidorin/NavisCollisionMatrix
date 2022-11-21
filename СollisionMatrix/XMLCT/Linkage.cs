using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "linkage")]
    public class Linkage
    {

        [XmlAttribute(AttributeName = "mode")]
        public string Mode { get; set; }
    }


}
