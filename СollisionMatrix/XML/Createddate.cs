using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "createddate")]
    public class Createddate
    {

        [XmlElement(ElementName = "date")]
        public Date Date { get; set; }
    }


}
