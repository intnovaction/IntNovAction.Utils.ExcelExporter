using IntNovAction.Utils.ExcelExporter.Exceptions;
using System.Collections.Generic;
using System.Linq;
using System;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// La colección de columnas dentro de una hoja
    /// </summary>
    internal class ColumnCollection
    {
        private List<ColumnConfigurator> _columnCol;
        public ColumnCollection()
        {
            _columnCol = new List<Utils.ColumnConfigurator>();
        }
        
        /// <summary>
        /// Añade una column
        /// </summary>
        /// <param name="sheetConfigurator">El configurador de la hoja</param>
        public void AddColumn(ColumnConfigurator columnConfigurator)
        {
            _columnCol.Add(columnConfigurator);
        }



    }
}