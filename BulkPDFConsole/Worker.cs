using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BulkPDF;
using System.Xml.Linq;
using System.IO;


namespace BulkPDFConsole
{
    class Worker
    {
        IDataSource dataSource;
        PDF pdf;
        Opciones opt;
        

        public bool Do(string configurationFilePath)
        {
            try
            {
                Console.WriteLine("--- Load configuration file ---");
                configurationFilePath = Environment.ExpandEnvironmentVariables(configurationFilePath);
                XDocument xDocument = XDocument.Parse(File.ReadAllText(configurationFilePath));

                //// Options
                var xmlOptions = xDocument.Root.Element("Options");
                // DataSource
                Console.WriteLine("Load spreadsheet");
                dataSource = new Spreadsheet();
                dataSource.Open(Environment.ExpandEnvironmentVariables(xmlOptions.Element("DataSource").Element("Parameter").Value));
                ((Spreadsheet)dataSource).SetSheet(xmlOptions.Element("Spreadsheet").Element("Table").Value);

                // PDF
                Console.WriteLine("Load PDF");
                
                pdf = PDF.Open(Environment.ExpandEnvironmentVariables(xmlOptions.Element("PDF").Element("Filepath").Value));

                //// PDFFieldValues
                Console.WriteLine("Load field configuration");
                Dictionary<string, PDFField> pdfFields = new Dictionary<string, PDFField>();
                foreach (var node in xDocument.Root.Element("PDFFieldValues").Descendants("PDFFieldValue"))
                {
                    var pdfField = new PDFField();
                    pdfField.Name = node.Element("Name").Value;
                    pdfField.Rellena(node);
                    pdfFields.Add(pdfField.Name, pdfField);
                }

                //// Carga opciones
                Console.WriteLine("Load options");
                opt = Opciones.Load(xmlOptions, dataSource);


                Console.WriteLine("--- Start processing ---");
                //  -- OLD PDFFiller.CreateFiles(pdf, dataSource, pdfFields, opt, ConcatFilename, WriteLinePercent, null); 
                PDFFiller filler = new PDFFiller(pdf, dataSource, opt, ConcatFilename);
                pdf.CreateFiles(filler, pdfFields, WriteLinePercent, null);

                Console.WriteLine("!!! Finished !!!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

        private string ConcatFilename(int pageGroup)
        {
            return Opciones.ConcatFilename(dataSource, opt, pageGroup);
        }

        private void WriteLinePercent(int percent)
        {
            Console.WriteLine(String.Format("{0:000}%", percent));
        }
    }
}