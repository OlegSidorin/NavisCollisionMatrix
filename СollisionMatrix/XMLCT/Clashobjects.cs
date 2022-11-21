using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashobjects")]
    public class Clashobjects
    {

        [XmlElement(ElementName = "clashobject")]
        public List<Clashobject> Clashobject { get; set; }
    }


}
