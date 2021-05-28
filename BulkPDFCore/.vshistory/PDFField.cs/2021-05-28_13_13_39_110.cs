using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace BulkPDF
{
    public struct PDFField
    {
        public string Name;
        public string Typ;
        public string CurrentValue;
        public string DataSourceValue;
        public bool MakeReadOnly;
        public bool UseValueFromDataSource;
        public int Row;
        /// <summary>
        /// Rellena el objeto con los datos de la configuración
        /// </summary>
        /// <param name="node"></param>
        public void Rellena(XElement node)
        {
            this.DataSourceValue = node.Element("NewValue").Value;
            this.UseValueFromDataSource = Convert.ToBoolean(node.Element("UseValueFromDataSource").Value);
            this.MakeReadOnly = Convert.ToBoolean(node.Element("MakeReadOnly").Value);
            int.TryParse(node.Element("Row").Value, out int row);
            Row = row;
        }
        public void XMLWritePDFField(XmlWriter xmlWriter)
        {
            xmlWriter.WriteStartElement("PDFFieldValue"); // <PDFFieldValue>
            xmlWriter.WriteElementString("Name", Name);
            xmlWriter.WriteElementString("NewValue", DataSourceValue);
            xmlWriter.WriteElementString("UseValueFromDataSource", UseValueFromDataSource.ToString());
            xmlWriter.WriteElementString("MakeReadOnly", MakeReadOnly.ToString());
            xmlWriter.WriteElementString("Row", Row.ToString());
            xmlWriter.WriteEndElement(); // </PDFFieldValue>
        }
    }
}
