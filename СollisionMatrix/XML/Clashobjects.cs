using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XML
{
    [XmlRoot(ElementName = "clashobjects")]
    public class Clashobjects
    {

        [XmlElement(ElementName = "clashobject")]
        public List<Clashobject> Clashobject { get; set; }
    }


}
