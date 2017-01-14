using System;
using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    /// <summary>
    /// La configuración de una columna
    /// </summary>
    /// <typeparam name="TDataItem">El parametro del item del que se va a pintar la columna</typeparam>
    public class ColumnConfigurator<TDataItem>
    {
        public ColumnConfigurator()
        {
            _columnHeaderFormat = new ColumnHeaderConfigurator();
        }

        /// <summary>
        /// En caso de que se utilice una expresión para poner un valor a la celda, la expresión
        /// </summary>
        private Func<TDataItem, object> _cellValueExpression = null;

        internal ColumnHeaderConfigurator _columnHeaderFormat;

        /// <summary>
        /// La configuración de formato de la columna
        /// </summary>
        internal ColumnDataFormatConfigurator _columnFormat { get; set; }



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

        public ColumnConfigurator<TDataItem> Header(string title)
        {
            _columnHeaderFormat.Text(title);
            return this;
        }

        /// <summary>
        /// Establece un formato para la columna
        /// </summary>
        /// <param name="formatConfigurator"></param>
        public void DataFormat(Action<ColumnDataFormatConfigurator> formatConfigurator)
        {
            _columnFormat = new ColumnDataFormatConfigurator();
            formatConfigurator.Invoke(_columnFormat);
        }

        public ColumnConfigurator<TDataItem> Header(Action<ColumnHeaderConfigurator> headerConfigurator)
        {
            headerConfigurator.Invoke(_columnHeaderFormat);

            return this;
        }

    }
}