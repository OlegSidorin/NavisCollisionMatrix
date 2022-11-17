using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "approveddate")]
    public class Approveddate
    {

        [XmlElement(ElementName = "date")]
        public Date Date { get; set; }
    }


}
