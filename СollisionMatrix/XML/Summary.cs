using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "summary")]
    public class Summary
    {

        [XmlElement(ElementName = "testtype")]
        public string Testtype { get; set; }

        [XmlElement(ElementName = "teststatus")]
        public string Teststatus { get; set; }

        [XmlAttribute(AttributeName = "total")]
        public int Total { get; set; }

        [XmlAttribute(AttributeName = "new")]
        public int New { get; set; }

        [XmlAttribute(AttributeName = "active")]
        public int Active { get; set; }

        [XmlAttribute(AttributeName = "reviewed")]
        public int Reviewed { get; set; }

        [XmlAttribute(AttributeName = "approved")]
        public int Approved { get; set; }

        [XmlAttribute(AttributeName = "resolved")]
        public int Resolved { get; set; }

        [XmlText]
        public string Text { get; set; }
    }


}
