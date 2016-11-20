using System;
using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// Configurador de una hoja
    /// </summary>
    public class SheetConfigurator<TDataItem> : SheetConfiguratorBase
        where TDataItem : new()
    {
        internal IEnumerable<TDataItem> _data;

        internal List<Tuple<Func<TDataItem, bool>, FormatConfigurator>> _fontFormatters;

        internal bool _hideHeaders = false;

        public SheetConfigurator()
        {
            _fontFormatters = new List<Tuple<Func<TDataItem, bool>, FormatConfigurator>>();
            //_fontSizeFormatters = new List<Tuple<Func<TDataItem, bool>, int>>();
        }

        /// <summary>
        /// Añade un formato condicional de fila a una hoja
        /// </summary>
        /// <param name="condition">Condición para aplicar o no el formato</param>
        /// <param name="formatConfigurator">Expresión para especificar el formato de la fila</param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> AddFormatRule(Func<TDataItem, bool> condition, Action<FormatConfigurator> formatConfigurator)
        {
            var configurator = new FormatConfigurator();
            formatConfigurator.Invoke(configurator);

            _fontFormatters.Add(new Tuple<Func<TDataItem, bool>, FormatConfigurator>(condition, configurator));

            return this;
        }

        /// <summary>
        /// Indica que no se muestren los headers
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> HideHeaders()
        {
            _hideHeaders = true;
            return this;
        }

        /// <summary>
        /// Deja la fila de las cabeceras en blanco y empieza con los datos en la siguienteo
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> JumpHeaders()
        {
            _jumpHeaders = true;
            return this;
        }

        /// <summary>
        /// Establece el nombre de la hoja
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> Name(string name)
        {
            _name = name;
            return this;
        }

        /// <summary>
        /// Establece el orden en el que la hoja se añadirá al Excel
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> Order(int order)
        {
            _order = order;
            return this;
        }

        /// <summary>
        /// Establece los datos a mostrar en la hoja
        /// </summary>
        /// <param name="data">Datos a pintar en la hoja</param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetData(IEnumerable<TDataItem> data)
        {
            _data = data;
            return this;
        }

        /// <summary>
        /// Establece el titulo de la hoja
        /// </summary>
        /// <param name="title">El titulo a poner</param>
        /// <remarks>Todavia no funciona</remarks>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetTitle(string title)
        {
            _title = title;
            return this;
        }
    }
}