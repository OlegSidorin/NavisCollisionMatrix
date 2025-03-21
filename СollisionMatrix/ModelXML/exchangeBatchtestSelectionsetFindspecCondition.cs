namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestSelectionsetFindspecCondition
    {

        private exchangeBatchtestSelectionsetFindspecConditionCategory categoryField;

        private exchangeBatchtestSelectionsetFindspecConditionProperty propertyField;

        private exchangeBatchtestSelectionsetFindspecConditionValue valueField;

        private string testField;

        private byte flagsField;

        /// <remarks/>
        public exchangeBatchtestSelectionsetFindspecConditionCategory category
        {
            get
            {
                return this.categoryField;
            }
            set
            {
                this.categoryField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestSelectionsetFindspecConditionProperty property
        {
            get
            {
                return this.propertyField;
            }
            set
            {
                this.propertyField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestSelectionsetFindspecConditionValue value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string test
        {
            get
            {
                return this.testField;
            }
            set
            {
                this.testField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte flags
        {
            get
            {
                return this.flagsField;
            }
            set
            {
                this.flagsField = value;
            }
        }
    }


}
