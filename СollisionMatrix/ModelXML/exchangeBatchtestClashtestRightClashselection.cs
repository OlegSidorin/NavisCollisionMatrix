namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestClashtestRightClashselection
    {

        private string locatorField;

        private byte selfintersectField;

        private byte primtypesField;

        /// <remarks/>
        public string locator
        {
            get
            {
                return this.locatorField;
            }
            set
            {
                this.locatorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte selfintersect
        {
            get
            {
                return this.selfintersectField;
            }
            set
            {
                this.selfintersectField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte primtypes
        {
            get
            {
                return this.primtypesField;
            }
            set
            {
                this.primtypesField = value;
            }
        }
    }


}
