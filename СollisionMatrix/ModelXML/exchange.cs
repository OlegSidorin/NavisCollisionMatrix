using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace СollisionMatrix.ModelXML
{

    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    //[System.Xml.Serialization.XmlRootAttribute(Namespace = "http://www.w3.org/2001/XMLSchema", IsNullable = true)]
    public partial class exchange
    {

        private exchangeBatchtest batchtestField;

        private string xmlnsxsiField;

        private string xsinoNamespaceSchemaLocationField;

        private string unitsField;

        private string filenameField;

        private string filepathField;

        /// <remarks/>
        public exchangeBatchtest batchtest
        {
            get
            {
                return this.batchtestField;
            }
            set
            {
                this.batchtestField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string xsinoNamespaceSchemaLocation
        {
            get
            {
                return this.xsinoNamespaceSchemaLocationField;
            }
            set
            {
                this.xsinoNamespaceSchemaLocationField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string xmlnsxsi
        {
            get
            {
                return this.xmlnsxsiField;
            }
            set
            {
                this.xmlnsxsiField = value;
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

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string filename
        {
            get
            {
                return this.filenameField;
            }
            set
            {
                this.filenameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string filepath
        {
            get
            {
                return this.filepathField;
            }
            set
            {
                this.filepathField = value;
            }
        }
    }


}
