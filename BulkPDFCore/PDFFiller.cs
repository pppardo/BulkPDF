//Sample license text.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static BulkPDF.PDF;

namespace BulkPDF {
    public delegate string DelGetFilename(int pageGrupo);
    public delegate void DelSetPercent(int percent);
    public delegate bool DelIsAborted();
    public delegate bool DelFiltro(IDataSource dataSource, int campo);


    public class FillData {
        //public IEnumerable<FieldWriteData> listaCampos = new List<FieldWriteData>();
        public List<FieldWriteData> listaCampos = new List<FieldWriteData>();
        public string filename;
        /// <summary> Porcentaje realizado </summary>
        public int realizado;
    }

    /// <summary>
    /// Clase para rellenar formularios
    /// </summary>
    public class PDFFiller {

        PDF pdf;
        IDataSource dataSource;
        public Opciones opt;
        DelGetFilename GetFilename;

        /// <summary>
        /// Constructor de PDFFiller, Versión con datos internos 
        /// </summary>
        /// <param name="pdf"></param>
        /// <param name="dataSource"></param>
        /// <param name="opt"></param>
        /// <param name="getFilename"></param>
        public PDFFiller(PDF pdf, IDataSource dataSource, Opciones opt, DelGetFilename getFilename)
        {
            this.pdf = pdf;
            this.dataSource = dataSource;
            this.opt = opt;
            GetFilename = getFilename;
            // DataSourceFilter
            if (opt.Filtered && filtros.ContainsKey(opt.FilterOperator))
            {
                dataSource.FilterCol = opt.DSColFilter + 1;
                dataSource.filtro = filtros[opt.FilterOperator];
            }
        }
        /// <summary>
        /// Verifica que haya nombres de archivo únicos, usa el mismo generador de datos que el que los genera (GetDataForms)
        /// </summary>
        /// <returns></returns>
        public int IsFilenameUnique()
        {
            var nombres = new HashSet<string>();

            foreach (FillData file in GetDataForms(null))
            {
                if (!nombres.Add(file.filename))
                {
                    return 0;
                }
            }
            return nombres.Count;
        }

        /// <summary>
        /// Método que se encarga de generar los datos para cada pag/archivo de salida. Funciona como un iterador
        /// </summary>
        /// <param name="pdfFields"></param>
        /// <returns></returns>
        public IEnumerable<FillData> GetDataForms(Dictionary<string, PDFField> pdfFields)
        {
            FillData datos = new FillData();
            datos.listaCampos = new List<FieldWriteData>();

            dataSource.ResetDataSource(); // Nuevo Para usarse hay que llamar a NextFila()

            // string grupoActual = "";
            int pagGrupo = 1;
            int rowsProcesadas = 0;

            for (dataSource.NextFila(); !dataSource.EOD();)
            {   // Página / PDF nueva
                int PDFRow = 1;
                int nFirstRow = dataSource.Row;
                string grupoActual = dataSource.GetField(opt.GroupByColumn);
                datos.filename = GetFilename(pagGrupo);
                do
                { // gira mientras esté en la misma página/pdf
                    // DataSource
                    if (pdfFields != null)
                    {
                        // No se puede diferir los campos porque varia la dataSource y el deleayed no se entera
                        foreach (string field in pdfFields.Keys)
                        {
                            if (pdfFields[field].UseValueFromDataSource && pdfFields[field].Row == PDFRow)
                            {
                                var value = dataSource.GetField(dataSource.Columns.FindIndex(x => x == pdfFields[field].DataSourceValue));
                                datos.listaCampos.Add(new FieldWriteData(field, value, pdfFields[field].MakeReadOnly));
                            }
                        }
                    }
                    PDFRow++;
                    rowsProcesadas++;
                } while (dataSource.NextFila() && (!dataSource.EsCambioGrupo(ref opt, grupoActual)) && PDFRow <= opt.RowsPerPag);

                datos.realizado = (int)((float)rowsProcesadas / (float)dataSource.PossibleRows * (float)100);

                yield return datos;

                // Comprueba si ha que enviar mas
                PDFRow = 1;

                if (!dataSource.EOD())
                {
                    datos.listaCampos.Clear();
                    // Caso 2 filtroActual != filtro
                    if (dataSource.EsCambioGrupo(ref opt, grupoActual))
                    {
                        pagGrupo = 1;
                    }
                    // Caso 3 PDFRow > rows
                    else // Siguiente página
                    {
                        pagGrupo++;
                    }
                }
            }
            dataSource.ResetRowCounter();
        }

        private IEnumerable<FieldWriteData> GetFields(Dictionary<string, PDFField> pdfFields, int PDFRow)
        {
            var fieldData = new FieldWriteData();
            foreach (string field in pdfFields.Keys)
            {
                if (pdfFields[field].UseValueFromDataSource && pdfFields[field].Row == PDFRow)
                {
                    var value = dataSource.GetField(dataSource.Columns.FindIndex(x => x == pdfFields[field].DataSourceValue));
                    fieldData.Name = field;
                    fieldData.Value = value;
                    fieldData.MakeReadOnly = pdfFields[field].MakeReadOnly;
                    yield return fieldData;         //new FieldWriteData(field, value, pdfFields[field].MakeReadOnly);
                }
            }
        }

        //------------------------------------- ESTATICOS 
        /// <summary>
        /// Lista de filtros disponibles
        /// </summary>
        public static Dictionary<string, DelFiltro> filtros = new Dictionary<string, DelFiltro>();
        static PDFFiller()
            {
            filtros.Add("FilterBlanc", EsBlanco);
            filtros.Add("FilterNoBlanc", NoEsBlanco);
            }

        /// <summary>
        /// Delegado para filtrar el DS por un campo que esté en blanco
        /// </summary>
        /// <param name="dataSource"></param>
        /// <param name="campo"></param>
        /// <returns></returns>
        public static bool EsBlanco(IDataSource dataSource, int campo)
        {
            if (campo != 0)
                return ("".Equals(dataSource.GetField(campo)));
            else
                return true;
        }
        /// <summary>
        /// Delegado para filtrar el DS por un campo que no esté en blanco
        /// </summary>
        /// <param name="dataSource"></param>
        /// <param name="campo"></param>
        /// <returns></returns>
        public static bool NoEsBlanco(IDataSource dataSource, int campo)
        {
            if (campo != 0)
                return (!"".Equals(dataSource.GetField(campo)));
            else
                return true;
        }
    }
}
