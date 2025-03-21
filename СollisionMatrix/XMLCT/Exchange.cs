using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "exchange")]
    public class Exchange
    {

        [XmlElement(ElementName = "batchtest")]
        public Batchtest Batchtest { get; set; }

        [XmlAttribute(AttributeName = "units")]
        public string Units { get; set; }

        [XmlAttribute(AttributeName = "filename")]
        public string Filename { get; set; }

        [XmlAttribute(AttributeName = "filepath")]
        public string Filepath { get; set; }

        [XmlAttribute(AttributeName = "xmlns_xsi")]
        public string XmlnsXsi { get; set; }

        [XmlAttribute(AttributeName = "xsi_noNamespaceSchemaLocation")]
        public string XsiNoNamespaceSchemaLocation { get; set; }

        [XmlText]
        public string Text { get; set; }
    }


}
