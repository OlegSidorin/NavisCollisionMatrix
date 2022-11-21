using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashpoint")]
    public class Clashpoint
    {

        [XmlElement(ElementName = "pos3f")]
        public Pos3f Pos3f { get; set; }
    }


}
