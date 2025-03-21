namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestClashtest
    {

        private exchangeBatchtestClashtestLinkage linkageField;

        private exchangeBatchtestClashtestLeft leftField;

        private exchangeBatchtestClashtestRight rightField;

        private object rulesField;

        private exchangeBatchtestClashtestSummary summaryField;

        private exchangeBatchtestClashtestClashresult[] clashresultsField;

        private string nameField;

        private string test_typeField;

        private string statusField;

        private decimal toleranceField;

        private byte merge_compositesField;

        /// <remarks/>
        public exchangeBatchtestClashtestLinkage linkage
        {
            get
            {
                return this.linkageField;
            }
            set
            {
                this.linkageField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestClashtestLeft left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestClashtestRight right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        /// <remarks/>
        public object rules
        {
            get
            {
                return this.rulesField;
            }
            set
            {
                this.rulesField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestClashtestSummary summary
        {
            get
            {
                return this.summaryField;
            }
            set
            {
                this.summaryField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("clashresult", IsNullable = false)]
        public exchangeBatchtestClashtestClashresult[] clashresults
        {
            get
            {
                return this.clashresultsField;
            }
            set
            {
                this.clashresultsField = value;
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
        public string test_type
        {
            get
            {
                return this.test_typeField;
            }
            set
            {
                this.test_typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string status
        {
            get
            {
                return this.statusField;
            }
            set
            {
                this.statusField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal tolerance
        {
            get
            {
                return this.toleranceField;
            }
            set
            {
                this.toleranceField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte merge_composites
        {
            get
            {
                return this.merge_compositesField;
            }
            set
            {
                this.merge_compositesField = value;
            }
        }
    }


}
