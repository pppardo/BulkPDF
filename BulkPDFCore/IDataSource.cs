using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BulkPDF
{
    /// <summary>
    /// Interfaz Datasource
    /// </summary>
    public interface IDataSource
    {
        /// <summary>
        /// Columnas
        /// </summary>
        List<string> Columns { get; }

        /// <summary>
        /// Posibles filas
        /// </summary>
        int PossibleRows { get; }
        /// <summary>
        /// Parámetros
        /// </summary>
        string Parameter { get; }
        /// <summary>
        /// Fila actual
        /// </summary>
        int Row { get; }
        /// <summary>
        /// Columna de filtro (relativo 0)
        /// </summary>
        int FilterCol { get; set; }
        /// <summary>
        /// Filtro. Actualmente EsBlanco()
        /// </summary>
        DelFiltro filtro { get; set; }

        // METHODS
        /// <summary>
        /// Abre DataSource
        /// </summary>
        /// <param name="parameter"></param>
        void Open(string parameter);
        /// <summary>
        /// Cierra DataSource
        /// </summary>
        void Close();
        /// <summary>
        /// Siguente fila, devuelve true si hay siguiente. Si es false el DataSource está inservible
        /// Obsoleto usar(NextFila)
        /// </summary>
        /// <returns></returns>
        bool NextRow();
        /// <summary>
        /// Inicializa las filas. No se puede usar hasta despues de llamar a NextRow()
        /// Obsoleto (usar ResetDataSource)
        /// </summary>
        void ResetRowCounter();
        /// <summary>
        /// Obtiene el valor de la columna columnIndex. La primera columna es cero.
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        string GetField(int columnIndex);
        // Resetea el datasource pero lo pone en un estado preUso, hay que llamar a NextFila

        /// <summary>
        /// Inicializa las filas. No se puede usar hasta despues de llamar a NextRow()
        /// </summary>
        void ResetDataSource();

        /// <summary>
        /// Se posiciona en la siguiente fila usando el filtro
        /// Posiciona el cursor en la siguiente fila válida devuelve false si no hay mas utiliza un filtro
        /// </summary>
        /// <returns></returns>
        bool NextFila();

        /// <summary>
        /// End of data (DataSource fuer de rango) 
        /// </summary>
        /// <returns></returns>
        bool EOD();
 
        /// <summary>
        /// Obtiene el valor del campo referenciado por columnIndex empezando en cero
        /// </summary>
        /// <returns></returns>
        bool EsCambioGrupo(ref Opciones file, string grupoActual);
    }
}
