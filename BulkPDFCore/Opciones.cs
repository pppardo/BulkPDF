using BulkPDF;
using System;
using System.Xml;
using System.Xml.Linq;

namespace BulkPDF
{
    public struct Opciones
    {
        private const string XML_FILENAME = "Filename";
        private const string XML_PREFIX = "Prefix";
        private const string XML_DSVALUE = "UseValueFromDataSource";
        private const string XML_DSCOLUMN = "DataSourceColumnsFilename";
        private const string XML_SUFIX1 = "Suffix1";
        private const string XML_USEROW = "UseRowNumber";
        private const string XML_USEGRP = "UseGroup";
        private const string XML_SUFIX2 = "Suffix2";
        private const string XML_USEPAG = "UsePag";
        private const string XML_OUTDIR = "OutputDir";
        private const string XML_FUSION = "Fusion";
        private const string XML_FLATTEN = "Flatten";
        private const string XML_ROWSPP = "RowsPerPag";
        private const string XML_GROUPED = "Grouped";
        private const string XML_GROUPCOL = "GroupByColumn";
        private const string XML_FINALIZE = "Finalize";
        private const string XML_UNICODE = "Unicode";
        private const string XML_FILTERED = "Filtered";
        private const string XML_FILTERFIELD = "FilterField";
        private const string XML_FILTEROP = "FilterOp";
        private const string XML_CUSTOMFONT = "CustomFont";
        private const string XML_CUSTOMFONTPATH = "CustomFontPath";

        public string Prefix { get; set; }
        public string Sufix1 { get; set; }
        public string Sufix2 { get; set; }
        public bool UseValueFromDataSource { get; set; }
        public int DataSourceColumnsFilename { get; set; }
        public bool UseRowNumber { get; set; }
        public bool UseGroup { get; set; }
        public int GroupByColumn { get; set; }
        public bool UsePag { get; set; }
        public string OutputDir { get; set; }

        /// <summary> Indica si la salida es un solo archivo fusión de todos /// </summary>
        public bool Fusion { get; set; }
        public bool Flatten { get; set; }

        /// <summary> Filas de la tabla que queremos introducir en una página del PDF, por defecto 1 /// </summary>
        public int RowsPerPag { get; set; }
        public bool Grouped { get; set; }
        public bool Finalize { get; set; }
        public bool Unicode { get; set; }
        public bool CustomFont { get; set; }
        public string CustomFontPath { get; set; }

        public bool Filtered { get; set; }
        public int DSColFilter { get; set; }
        public string FilterOperator { get; set; }

        public static Opciones Load(XElement xmlOptions, IDataSource dataSource)
        {
            Opciones opt = new Opciones();
            // Primero hay que cargar las opcions referentes a agrupación para que se pueda activar el usar grupo
            var xmlFilename = xmlOptions.Element(XML_FILENAME);
            opt.RowsPerPag = Convert.ToInt32(xmlOptions.Element(XML_ROWSPP).Value);
            opt.Grouped = Convert.ToBoolean(xmlOptions.Element(XML_GROUPED).Value);
            opt.GroupByColumn = ((Spreadsheet)dataSource).Columns.IndexOf(xmlOptions.Element(XML_GROUPCOL).Value);
            opt.Finalize = Convert.ToBoolean(xmlOptions.Element(XML_FINALIZE).Value);
            opt.Unicode = Convert.ToBoolean(xmlOptions.Element(XML_UNICODE).Value);
            try
            {
                opt.CustomFont = Convert.ToBoolean(xmlOptions.Element(XML_CUSTOMFONT).Value);
                opt.CustomFontPath = Environment.ExpandEnvironmentVariables(xmlOptions.Element(XML_CUSTOMFONTPATH).Value);
            }
            catch
            {
                // Ignore. Ugly but don't hurt anyone.
            }

            opt.Filtered = Convert.ToBoolean(xmlOptions.Element(XML_FILTERED).Value);
            opt.DSColFilter = ((Spreadsheet)dataSource).Columns.IndexOf(xmlOptions.Element(XML_FILTERFIELD).Value);
            opt.FilterOperator = xmlOptions.Element(XML_FILTEROP).Value;
            opt.Flatten = Convert.ToBoolean(xmlOptions.Element(XML_FLATTEN).Value);

            // Filename
            opt.Prefix = xmlFilename.Element(XML_PREFIX).Value;
            opt.UseValueFromDataSource = Convert.ToBoolean(xmlFilename.Element(XML_DSVALUE).Value);
            opt.DataSourceColumnsFilename = ((Spreadsheet)dataSource).Columns.IndexOf(xmlFilename.Element(XML_DSCOLUMN).Value);
            opt.Sufix1 = xmlFilename.Element(XML_SUFIX1).Value;
            opt.UseRowNumber = Convert.ToBoolean(xmlFilename.Element(XML_USEROW).Value);
            opt.UseGroup = Convert.ToBoolean(xmlFilename.Element(XML_USEGRP).Value);
            opt.Sufix2 = xmlFilename.Element(XML_SUFIX2).Value;
            opt.UsePag = Convert.ToBoolean(xmlFilename.Element(XML_USEPAG).Value);
            opt.OutputDir = xmlFilename.Element(XML_OUTDIR).Value;
            opt.Fusion = Convert.ToBoolean(xmlFilename.Element(XML_FUSION).Value);
            
            // FIN de Filename
            if (!opt.Grouped)
            {
                opt.UseGroup = false;
            }


            return opt;
        }
        public void Save(XmlWriter xmlWriter, IDataSource dataSource)
        {
            xmlWriter.WriteStartElement(XML_FILENAME); // <Filename>
            xmlWriter.WriteElementString(XML_PREFIX, Prefix);
            xmlWriter.WriteElementString(XML_DSVALUE, UseValueFromDataSource.ToString());
            xmlWriter.WriteElementString(XML_DSCOLUMN, ((Spreadsheet)dataSource).Columns[DataSourceColumnsFilename]);
            xmlWriter.WriteElementString(XML_SUFIX1, Sufix1);
            xmlWriter.WriteElementString(XML_USEROW, UseRowNumber.ToString());
            xmlWriter.WriteElementString(XML_USEGRP, UseGroup.ToString());
            xmlWriter.WriteElementString(XML_SUFIX2, Sufix2);
            xmlWriter.WriteElementString(XML_USEPAG, UsePag.ToString());
            xmlWriter.WriteElementString(XML_OUTDIR, OutputDir);
            xmlWriter.WriteElementString(XML_FUSION, Fusion.ToString());


            xmlWriter.WriteEndElement(); // </Filename>

            xmlWriter.WriteElementString(XML_FLATTEN, Flatten.ToString());
            xmlWriter.WriteElementString(XML_ROWSPP, RowsPerPag.ToString());
            xmlWriter.WriteElementString(XML_GROUPED, Grouped.ToString());
            xmlWriter.WriteElementString(XML_GROUPCOL, ((Spreadsheet)dataSource).Columns[GroupByColumn]);
            xmlWriter.WriteElementString(XML_FINALIZE, Finalize.ToString());
            xmlWriter.WriteElementString(XML_UNICODE, Unicode.ToString());
            xmlWriter.WriteElementString(XML_FILTERED, Filtered.ToString());
            xmlWriter.WriteElementString(XML_FILTERFIELD, ((Spreadsheet)dataSource).Columns[DSColFilter]);
            xmlWriter.WriteElementString(XML_FILTEROP, FilterOperator.ToString());
            xmlWriter.WriteElementString(XML_CUSTOMFONT, CustomFont.ToString());
            xmlWriter.WriteElementString(XML_CUSTOMFONTPATH, CustomFontPath);
        }
        public static string ConcatFilename(IDataSource dataSource, Opciones opt, int pageGroup)
        {
            // file.Prefix, file.UseValueFromDataSource, file.DSColumn + 1, file.Suffix1, file.UseRowNumber, file.Suffix2, file.UseGrupo, file.GrupoIndex,
            string filename = "";
            if (!opt.Fusion)
            {
                filename += opt.Prefix;
                if (opt.UseValueFromDataSource)
                    filename += dataSource.GetField(opt.DataSourceColumnsFilename);
                filename += opt.Sufix1;
                if (opt.Grouped && opt.UseGroup)
                    filename += dataSource.GetField(opt.GroupByColumn);
                filename += opt.Sufix2;
                if (opt.UseRowNumber)
                    filename += dataSource.Row;
                if (opt.UsePag)
                    filename += pageGroup;
            }
            filename += ".pdf";

            return filename;

        }
    }
}