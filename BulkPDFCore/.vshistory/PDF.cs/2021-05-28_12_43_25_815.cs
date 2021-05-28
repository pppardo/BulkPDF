using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.exceptions;
using System.Text.RegularExpressions;

namespace BulkPDF
{
    enum AcroFieldsTypes
    {
        BUTTON = 1,
        CHECK_BOX = 2,
        RADIO_BUTTON = 3,
        TEXT_FIELD = 4,
        LIST_BOX = 5,
        COMBO_BOX = 6
    }

    /// <summary>
    /// PDF Dinámico
    /// </summary>
    public class PDFDynamic : PDF
    {
        internal PDFDynamic(PdfReader pdfReader)
        {
            this.pdfReader = pdfReader;
        }

        internal override void Finalizar(PdfStamper pdfStamper)
        {
            pdfStamper.AcroFields.Xfa.FillXfaForm(pdfStamper.AcroFields.Xfa.DatasetsNode, true);
        }

        internal override void HacerSoloLectura(PdfStamper pdfStamper, FieldWriteData field)
        {
            string name = Regex.Match(field.Name, @"([A-Za-z0-9]+)(\[[0-9]+\]|)$").Groups[1].Value;
            for (int i = 0; i < pdfStamper.AcroFields.Xfa.DomDocument.SelectNodes("//*[@name='" + name + "']").Count; i++)
            {
                var attr = pdfStamper.AcroFields.Xfa.DomDocument.CreateAttribute("access");
                attr.Value = "readOnly";
                pdfStamper.AcroFields.Xfa.DomDocument.SelectNodes("//*[@name='" + name + "']")[i].Attributes.Append(attr);
            }
        }
        internal override void EscribirCampos(PdfStamper pdfStamper, FieldWriteData field, bool unicode, bool customFont)
        {
            var node = pdfStamper.AcroFields.Xfa.FindDatasetsNode(field.Name);
            var text = node.OwnerDocument.CreateTextNode(field.Value);
            node.AppendChild(text);

            pdfStamper.AcroFields.Xfa.Changed = true;
        }
        public override List<PDFField> ListFields()
        {
            XfaForm xfa = new XfaForm(pdfReader);
            var acroFields = pdfReader.AcroFields;
            return ListDynamicXFAFields(acroFields.Xfa.DatasetsNode.FirstChild);
        }

        private List<PDFField> ListDynamicXFAFields(System.Xml.XmlNode n)
        {
            List<PDFField> pdfFields = new List<PDFField>();

            foreach (System.Xml.XmlNode child in n.ChildNodes) // > 0 Childs == Group
                pdfFields.AddRange(ListDynamicXFAFields(child)); // Search field

            if (n.ChildNodes.Count == 0) // 0 Childs == Field 
            {
                var acroFields = pdfReader.AcroFields;

                var pdfField = new PDFField();

                // If a value is set the value of n.Name would be "#text"
                if ((n.Name.ToCharArray(0, 1))[0] != '#')
                {
                    pdfField.Name = acroFields.GetTranslatedFieldName(n.Name);
                }
                else
                {
                    pdfField.Name = acroFields.GetTranslatedFieldName(n.ParentNode.Name);
                }

                pdfField.CurrentValue = n.Value;

                pdfField.Typ = "";

                pdfFields.Add(pdfField);

                return pdfFields;
            }

            return pdfFields;
        }
        public override string getTipoTexto()
        {
            return "XFA Form";
        }

    }
    /// <summary>
    /// PDF Dinámico
    /// </summary>
    public class PDFStatic : PDF
    {
        internal PDFStatic(PdfReader pdfReader)
        {
            this.pdfReader = pdfReader;
        }
        internal override void Finalizar(PdfStamper pdfStamper)
        {
            foreach (var field in ListFields())
                pdfStamper.AcroFields.SetFieldProperty(field.Name, "setfflags", PdfFormField.FF_READ_ONLY, null);
        }
        internal override void HacerSoloLectura(PdfStamper pdfStamper, FieldWriteData field)
        {
            // Read only for not dynamic XFAs
            pdfStamper.AcroFields.SetFieldProperty(field.Name, "setfflags", PdfFormField.FF_READ_ONLY, null);
        }
        internal override void EscribirCampos(PdfStamper pdfStamper, FieldWriteData field, bool unicode, bool customFont)
        {
            string value = field.Value;
            AcroFields.Item item = pdfStamper.AcroFields.GetFieldItem(field.Name);
            int x = pdfStamper.AcroFields.GetFieldType(field.Name);
            // AcroFields radiobutton start by zero -> dataSourceIndex-1 
            if (pdfStamper.AcroFields.GetFieldType(field.Name) == (int)AcroFieldsTypes.RADIO_BUTTON)
            {
                int radiobuttonIndex = 0;
                if (int.TryParse(field.Value, out radiobuttonIndex))
                {
                    radiobuttonIndex--;
                    value = radiobuttonIndex.ToString();
                }
            }
            else if (x == (int)AcroFieldsTypes.CHECK_BOX)
            {
                string[] vs = pdfStamper.AcroFields.GetAppearanceStates(field.Name);

            }

            // Unicode
            if (unicode)
            {
                BaseFont bf = BaseFont.CreateFont(Path.Combine(Directory.GetCurrentDirectory(), "unifont.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                pdfStamper.AcroFields.AddSubstitutionFont(bf);
                pdfStamper.AcroFields.SetFieldProperty(field.Name, "textfont", bf, null);
                pdfStamper.AcroFields.RegenerateField(field.Name);
            }

            // Write text
            pdfStamper.AcroFields.SetField(field.Name, value, true);
        }
        public override List<PDFField> ListFields()
        {
            return ListGenericFields();
        }
        private List<PDFField> ListGenericFields()
        {
            var fields = new List<PDFField>();
            var acroFields = pdfReader.AcroFields;

            foreach (var field in acroFields.Fields)
                if (acroFields.GetFieldType(field.Key.ToString()) != (int)AcroFieldsTypes.BUTTON)
                {
                    // If readonly, continue loop
                    var n = field.Value.GetMerged(0).GetAsNumber(PdfName.FF);
                    if (n != null && ((n.IntValue & (int)PdfFormField.FF_READ_ONLY) > 0))
                    {
                        continue;
                    }
                    else
                    {
                        var pdfField = new PDFField();
                        pdfField.Name = field.Key.ToString();
                        pdfField.CurrentValue = acroFields.GetField(field.Key.ToString());
                        try
                        {
                            pdfField.Typ = Convert.ToString((AcroFieldsTypes)acroFields.GetFieldType(pdfField.Name));
                        }
                        catch (Exception)
                        {
                            pdfField.Typ = "";
                        }

                        fields.Add(pdfField);
                    }
                }

            return fields;
        }
        public override string getTipoTexto()
        {
            return "Acroform";
        }

    }

    /// <summary>
    /// Clase
    /// </summary>
    public abstract class PDF
    {
        private protected PdfReader pdfReader;
//        List<FieldWriteData> writerFieldList = new List<FieldWriteData>();


        //public bool IsXFA
        //{
        //    get { return isXFA; }
        //}
        //bool isXFA = false;
        //bool isDynamicXFA = false;

        /// <summary> Finalizar </summary>
        /// <param name="pdfStamper"></param>
        internal abstract void Finalizar(PdfStamper pdfStamper);

        internal abstract void HacerSoloLectura(PdfStamper pdfStamper, FieldWriteData field);

        internal abstract void EscribirCampos(PdfStamper pdfStamper, FieldWriteData field, bool unicode, bool customFont);

        /// <summary>Genera la lista del PDF. Depende de si es Formulario Dinámico a estático</summary>
        public abstract List<PDFField> ListFields();

        /// <summary> Obtiene el tipo de formulario como texto para poner en el label</summary><returns>Cadena con el tipo de formulario</returns>
        public abstract string getTipoTexto();

        /// <summary>Estructura para almacenar un campo del formulario con su valor</summary>
        public struct FieldWriteData
        {
            /// <summary>Nombre del campo de formulario</summary>
            public string Name;
            /// <summary>Valor del campo de formulario</summary>
            public string Value;
            /// <summary>Especifique que hay que poner el campo como solo lectura.</summary>
            public bool MakeReadOnly;
 //           private string field;
            /// <summary>Estructura para almacenar un campo del formulario con su valor</summary>
            public FieldWriteData(string name, string value, bool makeReadOnly) : this()
            {
                Name = name;
                Value = value;
                MakeReadOnly = makeReadOnly;
            }
        }

        public static PDF Open(String filePath)
        {
            PdfReader pdfReader;
            try
            {
                PdfReader.unethicalreading = true;
                pdfReader = new PdfReader(filePath);
                //pdfReader.RemoveUsageRights();
            }
            catch (InvalidPdfException e)
            {
                throw new Exception(e.ToString());
            }

            // XFA?
            XfaForm xfa = new XfaForm(pdfReader);
            if (xfa != null && xfa.XfaPresent)
            {
                //isXFA = true;

                if (xfa.Reader.AcroFields.Fields.Keys.Count == 0)
                {
                    //isDynamicXFA = true;
                    return new PDFDynamic(pdfReader);
                }
                else
                {
                    //isDynamicXFA = false;
                    return new PDFStatic(pdfReader);
                }
            }
            else
            {
                return new PDFStatic(pdfReader);
                //isXFA = false;
            }
        }

        public void Close()
        {
            if (pdfReader != null)
                pdfReader.Close();
        }

        public void SaveFilledPDF(string filePath, byte[] content)
        {
            try
            {
                using (var fs = File.Create(filePath))
                {
                    fs.Write(content, 0, (int)content.Length);
                }
            }
            catch (IOException e)
            {
                throw new Exception(e.Message);
            }
        }

        private byte[] GeneratePDFBytes(IEnumerable<FieldWriteData> writerFieldList, bool flatten, bool finalize, bool unicode, bool customFont)
        {
            PdfReader copiedPdfReader = new PdfReader(pdfReader);
            var pdfStamperMemoryStream = new MemoryStream();
            PdfStamper pdfStamper = new PdfStamper(copiedPdfReader, pdfStamperMemoryStream, '\0', true);

            // Fill
            foreach (FieldWriteData field in writerFieldList)
            {
                // Write
                EscribirCampos(pdfStamper, field, unicode, customFont );

                // Read Only
                if (field.MakeReadOnly)
                {
                    HacerSoloLectura(pdfStamper, field);
                }
            }

            // Global Finalize
            if (finalize)
            {
                Finalizar(pdfStamper);
            }
            pdfStamper.Close();
            byte[] content = flatten?FlattenPdfFormToBytes(pdfStamperMemoryStream.ToArray()): pdfStamperMemoryStream.ToArray();
            return content;
        }

        private byte[] FlattenPdfFormToBytes(byte[] bytes)
        {
            PdfReader reader = new PdfReader(bytes);
            var memStream = new MemoryStream();
            var stamper = new PdfStamper(reader, memStream) { FormFlattening = true };
            stamper.Close();
            return memStream.ToArray();
        }

 
/*        /// <summary> Añade un campo al formulario</summary>
        /// <param name="fieldname"></param>
        /// <param name="value"></param>
        /// <param name="makeReadOnly"></param>
        public void SetFieldValue(string fieldname, string value, bool makeReadOnly = false)
        {
            var field = new FieldWriteData();
            field.Name = fieldname;
            field.Value = value;
            field.MakeReadOnly = makeReadOnly;

            writerFieldList.Add(field);
        }
*/
/*        public void ResetFieldValue()
        {
            writerFieldList.Clear();
        }
*/

        /// <summary>
        /// Creaficheros nuevo
        /// </summary>
        /// <param name="filler"> Data source, solo lo usa para pasrlo al rellenador</param>
        /// <param name="pdfFields"></param>
        /// <param name="setPercent"></param>
        /// <param name="isAborted"></param>
        public void CreateFiles(PDFFiller filler, Dictionary<string, PDFField> pdfFields,
             DelSetPercent setPercent = null, DelIsAborted isAborted = null)
        {
            //ResetFieldValue();

            var nombres = new HashSet<string>();

            foreach (FillData dataSet in filler.GetDataForms(pdfFields))
            {
                //writerFieldList = dataSet.listaCampos;

                SaveFilledPDF(filler.opt.OutputDir + @"\" + dataSet.filename, GeneratePDFBytes(dataSet.listaCampos, filler.opt.Flatten, filler.opt.Finalize, filler.opt.Unicode, filler.opt.CustomFont));

                //ResetFieldValue();

                setPercent?.Invoke(dataSet.realizado);
                if (isAborted != null && isAborted())
                    break;
            }
            if (setPercent != null) setPercent(100);
            // Abort?
        }
        /// <summary>
        /// Almacena un único PDF
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="pdfFields"></param>
        /// <param name="setPercent"></param>
        /// <param name="isAborted"></param>
        public void CreateOneFile(PDFFiller filler, Dictionary<string, PDFField> pdfFields,
             DelSetPercent setPercent = null, DelIsAborted isAborted = null)
        {
            Document doc = new Document();
            string filepath = filler.opt.OutputDir + @"\" + filler.opt.Prefix + ".pdf";
            FileStream fileStream = new FileStream(filepath, FileMode.CreateNew);
            PdfSmartCopy copy = new PdfSmartCopy(doc, fileStream);
            doc.Open();

            foreach (FillData dataSet in filler.GetDataForms(pdfFields))
            {
                PdfReader pdfReader2 = new PdfReader(GeneratePDFBytes(dataSet.listaCampos, filler.opt.Flatten, filler.opt.Finalize, filler.opt.Unicode, filler.opt.CustomFont));
                for (int i = 0; i < pdfReader2.NumberOfPages; i++)
                {
                    copy.AddPage(copy.GetImportedPage(pdfReader2, i+1));
                }

                if (setPercent != null) setPercent(dataSet.realizado);
                if (isAborted != null && isAborted())
                    break;
            }
            doc.Close();
            if (setPercent != null) setPercent(100);
            // Abort?
        }
    }
}
