using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "left")]
    public class Left
    {

        [XmlElement(ElementName = "clashselection")]
        public Clashselection Clashselection { get; set; }
    }


}
