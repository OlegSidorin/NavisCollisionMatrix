using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "date")]
    public class Date
    {

        [XmlAttribute(AttributeName = "year")]
        public int Year { get; set; }

        [XmlAttribute(AttributeName = "month")]
        public int Month { get; set; }

        [XmlAttribute(AttributeName = "day")]
        public int Day { get; set; }

        [XmlAttribute(AttributeName = "hour")]
        public int Hour { get; set; }

        [XmlAttribute(AttributeName = "minute")]
        public int Minute { get; set; }

        [XmlAttribute(AttributeName = "second")]
        public int Second { get; set; }
    }


}
