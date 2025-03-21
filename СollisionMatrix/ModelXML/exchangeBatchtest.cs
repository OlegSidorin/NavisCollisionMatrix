namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtest
    {

        private exchangeBatchtestClashtest[] clashtestsField;

        private exchangeBatchtestSelectionset[] selectionsetsField;

        private string nameField;

        private string internal_nameField;

        private string unitsField;

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("clashtest", IsNullable = false)]
        public exchangeBatchtestClashtest[] clashtests
        {
            get
            {
                return this.clashtestsField;
            }
            set
            {
                this.clashtestsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("selectionset", IsNullable = false)]
        public exchangeBatchtestSelectionset[] selectionsets
        {
            get
            {
                return this.selectionsetsField;
            }
            set
            {
                this.selectionsetsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string internal_name
        {
            get
            {
                return this.internal_nameField;
            }
            set
            {
                this.internal_nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string units
        {
            get
            {
                return this.unitsField;
            }
            set
            {
                this.unitsField = value;
            }
        }
    }


}
