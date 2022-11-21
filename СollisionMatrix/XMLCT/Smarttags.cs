using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "smarttags")]
    public class Smarttags
    {

        [XmlElement(ElementName = "smarttag")]
        public List<Smarttag> Smarttag { get; set; }
    }


}
