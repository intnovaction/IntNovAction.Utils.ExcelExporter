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
        /// Nombre a mostrar en la columna
        /// </summary>
        internal string _columnTitle { get; set; }

        public ColumnConfigurator<TDataItem> Title(string title)
        {
            _columnTitle = title;
            return this;
        }

        // TODO: Refactor, si solo se usa para montarlo.. no deberia estar ahi
        /// <summary>
        /// Orden que se saca de los metadatos. Luego no se usa!
        /// </summary>
        internal int _orderFromMetadata { get; set; }

        /// <summary>
        /// Propiedad de la clase correspondiente a la columna. Si es una expresión, es nula
        /// </summary>
        internal PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// En caso de que se utilice una expresión para poner un valor a la celda, la expresión
        /// </summary>
        private Func<TDataItem, object> _cellValueExpression = null;

        internal Func<TDataItem, object> Expression
        {
            get
            {
                return _cellValueExpression;
            }
            set
            {
                if (value != null && PropertyInfo != null)
                {
                    PropertyInfo = null;
                    _orderFromMetadata = int.MaxValue;
                }
                _cellValueExpression = value;
            }
        }
    }
}
