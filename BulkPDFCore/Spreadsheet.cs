 using SpreadsheetLight;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace BulkPDF
{
    public class Spreadsheet : IDataSource
    {
        //public delegate bool DelFiltro(int campo);

        public string Parameter
        {
            get { return parameter; }
        }
        string parameter = "";
        public List<string> Columns
        {
            get { return columns; }
        }
        List<string> columns = new List<string>();
        public Hashtable RowValues
        {
            get { return rowValues; }
        }
        Hashtable rowValues = new Hashtable();
        public int PossibleRows
        {
            get { return possibleRows; }
        }

        public int Row => rowIndex;

        int possibleRows = 0;

        public int UltimaFila { get; private set; }
        public int FilterCol { get; set; } = 0;
        public BulkPDF.DelFiltro filtro { get; set; }

        SLDocument slDocument;
        int rowIndex = 2;


        public void Open(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (StreamReader streamReader = new StreamReader(fileStream))
                {
                    slDocument = new SLDocument(fileStream);
                }
            }

            parameter = filePath;
            SetSheet(slDocument.GetSheetNames()[0]);
        }

        public void Close()
        {

        }

        public bool NextRow()
        {
            rowIndex++;
            return true;
        }

        /// <summary>
        /// Inicializa las filas. No se puede usar hasta despues de llamar a NextRow()
        /// Obsoleto (usar ResetDataSource)
        /// </summary>
        public void ResetRowCounter()
        {
            rowIndex = 2;
        }
        /// <summary>
        /// Obtiene el valor del campo referenciado por columnIndex empezando en cero
        /// </summary>
        public string GetField(int columnIndex)
        {
            return slDocument.GetCellValueAsString(rowIndex, columnIndex+1);
        }

        public List<string> GetSheetNames()
        {
            return slDocument.GetSheetNames();
        }

        public bool SetSheet(string name)
        {
            slDocument.SelectWorksheet(name);
            possibleRows = CountPossibleRows();
            UltimaFila = possibleRows + 1;
            columns = ListColumns();
            ResetRowCounter();

            return true;
        }

        private List<string> ListColumns()
        {
            List<string> columns = new List<string>();

            for (int x = 1; !string.IsNullOrEmpty(slDocument.GetCellValueAsString(1, x)); x++)
                columns.Add(slDocument.GetCurrentWorksheetName() + "[.]" + slDocument.GetCellValueAsString(1, x));


            return columns;
        }

        private int CountPossibleRows()
        {
            // headerColumnNumber
            int headerColumnNumber = 0;
            for (int column = 1; !string.IsNullOrEmpty(slDocument.GetCellValueAsString(1, column)); column++)
                headerColumnNumber += 1;

            int maxRowsTotal = 0;

            if (headerColumnNumber > 0)
            {
                for (int row = 2; true; row++)
                {
                    // Read row
                    var rowValues = new List<string>();
                    for (int column = 1; column <= headerColumnNumber; column++)
                        rowValues.Add(slDocument.GetCellValueAsString(row, column));

                    // Check if the row is empty
                    int notEmptyCells = 0;
                    foreach (var value in rowValues)
                        if (!String.IsNullOrEmpty(value))
                            notEmptyCells++;
                    if (notEmptyCells == 0)
                        break;

                    maxRowsTotal++;
                }
            }


            return maxRowsTotal;
        }

        /// <summary>
        /// Inicializa las filas. No se puede usar hasta despues de llamar a NextRow()
        /// </summary>
        public void ResetDataSource()
        {
            rowIndex = 1;
        }

        /// <summary>
        /// Se posiciona en la siguiente fila usando el filtro
        /// Posiciona el cursor en la siguiente fila válida devuelve false si no hay mas utiliza un filtro
        /// </summary>
        /// <returns></returns>
        public bool NextFila()
        {
            do
            {
                rowIndex++;
            } while (rowIndex <= UltimaFila && (filtro != null && !filtro(this, this.FilterCol)));
            return !EOD();
        }
        // End of data (DataSource fuera de rango)
        /// <summary>
        /// End of data (DataSource fuer de rango) 
        /// </summary>
        /// <returns></returns>
        public bool EOD()
        {
            return rowIndex > UltimaFila;
        }

        /// <summary>
        /// Obtiene el valor del campo referenciado por columnIndex empezando en cero
        /// </summary>
        /// <returns></returns>
        public bool EsCambioGrupo(ref Opciones file, string grupoActual)
        {
            return file.Grouped && !grupoActual.Equals(this.GetField(file.GroupByColumn));
        }
    }
}
