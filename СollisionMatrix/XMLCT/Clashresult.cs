using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashresult")]
    public class Clashresult
    {

        [XmlElement(ElementName = "description")]
        public string Description { get; set; }

        [XmlElement(ElementName = "resultstatus")]
        public string Resultstatus { get; set; }

        [XmlElement(ElementName = "clashpoint")]
        public Clashpoint Clashpoint { get; set; }

        [XmlElement(ElementName = "gridlocation")]
        public string Gridlocation { get; set; }

        [XmlElement(ElementName = "createddate")]
        public Createddate Createddate { get; set; }

        [XmlElement(ElementName = "clashobjects")]
        public Clashobjects Clashobjects { get; set; }

        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [XmlAttribute(AttributeName = "guid")]
        public string Guid { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string Status { get; set; }

        [XmlAttribute(AttributeName = "distance")]
        public double Distance { get; set; }

        [XmlText]
        public string Text { get; set; }
    }


}
