using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "right")]
    public class Right
    {

        [XmlElement(ElementName = "clashselection")]
        public Clashselection Clashselection { get; set; }
    }


}
