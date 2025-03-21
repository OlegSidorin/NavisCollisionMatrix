namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestClashtestClashresultClashobject
    {

        private exchangeBatchtestClashtestClashresultClashobjectObjectattribute objectattributeField;

        private string layerField;

        private string[] pathlinkField;

        private exchangeBatchtestClashtestClashresultClashobjectSmarttag[] smarttagsField;

        /// <remarks/>
        public exchangeBatchtestClashtestClashresultClashobjectObjectattribute objectattribute
        {
            get
            {
                return this.objectattributeField;
            }
            set
            {
                this.objectattributeField = value;
            }
        }

        /// <remarks/>
        public string layer
        {
            get
            {
                return this.layerField;
            }
            set
            {
                this.layerField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("node", IsNullable = false)]
        public string[] pathlink
        {
            get
            {
                return this.pathlinkField;
            }
            set
            {
                this.pathlinkField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("smarttag", IsNullable = false)]
        public exchangeBatchtestClashtestClashresultClashobjectSmarttag[] smarttags
        {
            get
            {
                return this.smarttagsField;
            }
            set
            {
                this.smarttagsField = value;
            }
        }
    }


}
