using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// La configuración de una columna
    /// </summary>
    public class ColumnConfigurator
    {
        /// <summary>
        /// Nombre a mostrar
        /// </summary>
        internal string DisplayName { get; set; }

        /// <summary>
        /// Orden en el que se muestra
        /// </summary>
        internal int Order { get; set; }

        /// <summary>
        /// Propiedad de la clase correspondiente a la columna
        /// </summary>
        internal PropertyInfo PropertyInfo { get; set; }
    }
}
