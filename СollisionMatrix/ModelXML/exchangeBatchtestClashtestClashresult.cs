namespace СollisionMatrix.ModelXML
{
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class exchangeBatchtestClashtestClashresult
    {

        private string descriptionField;

        private string resultstatusField;

        private exchangeBatchtestClashtestClashresultClashpoint clashpointField;

        private string gridlocationField;

        private exchangeBatchtestClashtestClashresultCreateddate createddateField;

        private exchangeBatchtestClashtestClashresultClashobject[] clashobjectsField;

        private string nameField;

        private string guidField;

        private string statusField;

        private decimal distanceField;

        /// <remarks/>
        public string description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }

        /// <remarks/>
        public string resultstatus
        {
            get
            {
                return this.resultstatusField;
            }
            set
            {
                this.resultstatusField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestClashtestClashresultClashpoint clashpoint
        {
            get
            {
                return this.clashpointField;
            }
            set
            {
                this.clashpointField = value;
            }
        }

        /// <remarks/>
        public string gridlocation
        {
            get
            {
                return this.gridlocationField;
            }
            set
            {
                this.gridlocationField = value;
            }
        }

        /// <remarks/>
        public exchangeBatchtestClashtestClashresultCreateddate createddate
        {
            get
            {
                return this.createddateField;
            }
            set
            {
                this.createddateField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("clashobject", IsNullable = false)]
        public exchangeBatchtestClashtestClashresultClashobject[] clashobjects
        {
            get
            {
                return this.clashobjectsField;
            }
            set
            {
                this.clashobjectsField = value;
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
        public string guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
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
        public decimal distance
        {
            get
            {
                return this.distanceField;
            }
            set
            {
                this.distanceField = value;
            }
        }
    }


}
