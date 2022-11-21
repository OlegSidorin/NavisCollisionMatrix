using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "pathlink")]
    public class Pathlink
    {

        [XmlElement(ElementName = "node")]
        public List<string> Node { get; set; }
    }


}
