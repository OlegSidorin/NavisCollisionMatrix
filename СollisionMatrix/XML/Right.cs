using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "right")]
    public class Right
    {

        [XmlElement(ElementName = "clashselection")]
        public Clashselection Clashselection { get; set; }
    }


}
