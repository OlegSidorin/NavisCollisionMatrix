using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashtest")]
    public class Clashtest
    {

        [XmlElement(ElementName = "clashresults")]
        public Clashresults Clashresults { get; set; }

        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [XmlAttribute(AttributeName = "test_type")]
        public string TestType { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string Status { get; set; }

        [XmlAttribute(AttributeName = "tolerance")]
        public double Tolerance { get; set; }

        [XmlAttribute(AttributeName = "merge_composites")]
        public int MergeComposites { get; set; }

        [XmlText]
        public string Text { get; set; }

        [XmlElement(ElementName = "linkage")]
        public Linkage Linkage { get; set; }

        [XmlElement(ElementName = "left")]
        public Left Left { get; set; }

        [XmlElement(ElementName = "right")]
        public Right Right { get; set; }

        [XmlElement(ElementName = "rules")]
        public object Rules { get; set; }

        [XmlElement(ElementName = "summary")]
        public Summary Summary { get; set; }
    }


}
