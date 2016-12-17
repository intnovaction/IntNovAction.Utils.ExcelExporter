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
    /// <typeparam name="TDataItem">El parametro del item del que se va a pintar la columna</typeparam>
    public class ColumnConfigurator<TDataItem>
    {
        /// <summary>
        /// Nombre a mostrar
        /// </summary>
        internal string _title { get; set; }

        public ColumnConfigurator<TDataItem> Title(string title)
        {
            _title = title;
            return this;
        }

        /// <summary>
        /// Orden en el que se muestra
        /// </summary>
        internal int Order { get; set; }

        /// <summary>
        /// Propiedad de la clase correspondiente a la columna
        /// </summary>
        internal PropertyInfo PropertyInfo { get; set; }

        private Func<TDataItem, object> _expr = null;

        public Func<TDataItem, object> Expression
        {
            get
            {
                return _expr;
            }
            internal set
            {
                _expr = value;
            }
        }
    }
}
