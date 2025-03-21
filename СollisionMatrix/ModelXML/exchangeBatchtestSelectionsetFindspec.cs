namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestSelectionsetFindspec
    {

        private exchangeBatchtestSelectionsetFindspecCondition[] conditionsField;

        private string locatorField;

        private string modeField;

        private byte disjointField;

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("condition", IsNullable = false)]
        public exchangeBatchtestSelectionsetFindspecCondition[] conditions
        {
            get
            {
                return this.conditionsField;
            }
            set
            {
                this.conditionsField = value;
            }
        }

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
        public string mode
        {
            get
            {
                return this.modeField;
            }
            set
            {
                this.modeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte disjoint
        {
            get
            {
                return this.disjointField;
            }
            set
            {
                this.disjointField = value;
            }
        }
    }


}
