using System;
using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// La configuración de una columna
    /// </summary>
    /// <typeparam name="TDataItem">El parametro del item del que se va a pintar la columna</typeparam>
    public class ColumnConfigurator<TDataItem>
    {
        /// <summary>
        /// En caso de que se utilice una expresión para poner un valor a la celda, la expresión
        /// </summary>
        private Func<TDataItem, object> _cellValueExpression = null;

        /// <summary>
        /// La configuración de formato de la columna
        /// </summary>
        internal ColumnFormatConfigurator _columnFormat { get; set; }

        /// <summary>
        /// Nombre a mostrar en la columna
        /// </summary>
        internal string _columnTitle { get; set; }

        // TODO: Refactor, si solo se usa para montarlo.. no deberia estar ahi
        /// <summary>
        /// Orden que se saca de los metadatos. Luego no se usa!
        /// </summary>
        internal int _orderFromMetadata { get; set; }

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

        /// <summary>
        /// Propiedad de la clase correspondiente a la columna. Si es una expresión, es nula
        /// </summary>
        internal PropertyInfo PropertyInfo { get; set; }

        public ColumnConfigurator<TDataItem> Title(string title)
        {
            _columnTitle = title;
            return this;
        }

        /// <summary>
        /// Establece un formato para la columna
        /// </summary>
        /// <param name="formatConfigurator"></param>
        internal void Format(Action<ColumnFormatConfigurator> formatConfigurator)
        {
            _columnFormat = new ColumnFormatConfigurator();
            formatConfigurator.Invoke(_columnFormat);
        }
    }
}