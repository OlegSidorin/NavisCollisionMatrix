namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestClashtestSummary
    {

        private string testtypeField;

        private string teststatusField;

        private ushort totalField;

        private ushort newField;

        private byte activeField;

        private byte reviewedField;

        private byte approvedField;

        private byte resolvedField;

        /// <remarks/>
        public string testtype
        {
            get
            {
                return this.testtypeField;
            }
            set
            {
                this.testtypeField = value;
            }
        }

        /// <remarks/>
        public string teststatus
        {
            get
            {
                return this.teststatusField;
            }
            set
            {
                this.teststatusField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ushort total
        {
            get
            {
                return this.totalField;
            }
            set
            {
                this.totalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ushort @new
        {
            get
            {
                return this.newField;
            }
            set
            {
                this.newField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte active
        {
            get
            {
                return this.activeField;
            }
            set
            {
                this.activeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte reviewed
        {
            get
            {
                return this.reviewedField;
            }
            set
            {
                this.reviewedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte approved
        {
            get
            {
                return this.approvedField;
            }
            set
            {
                this.approvedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte resolved
        {
            get
            {
                return this.resolvedField;
            }
            set
            {
                this.resolvedField = value;
            }
        }
    }


}
