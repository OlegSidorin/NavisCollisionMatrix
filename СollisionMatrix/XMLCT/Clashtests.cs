﻿using System.Collections.Generic;
using System.Xml.Serialization;

namespace СollisionMatrix.XMLCT
{
    [XmlRoot(ElementName = "clashtests")]
    public class Clashtests
    {

        [XmlElement(ElementName = "clashtest")]
        public List<Clashtest> Clashtest { get; set; }
    }


}
