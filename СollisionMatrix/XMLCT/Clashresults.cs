using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashresults")]
    public class Clashresults
    {

        [XmlElement(ElementName = "clashresult")]
        public List<Clashresult> Clashresult { get; set; }
    }


}
