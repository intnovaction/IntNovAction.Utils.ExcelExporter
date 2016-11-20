using System;
using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// Contiene los datos de una hoja
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
        /// <param name="condition"></param>
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

        public SheetConfigurator<TDataItem> JumpHeaders()
        {
            _jumpHeaders = true;
            return this;
        }

        //internal List<Tuple<Func<TDataItem, bool>, int>> _fontSizeFormatters;
        public SheetConfigurator<TDataItem> Name(string name)
        {
            _name = name;
            return this;
        }

        public SheetConfigurator<TDataItem> Order(int order)
        {
            _order = order;
            return this;
        }

        public SheetConfigurator<TDataItem> SetData(IEnumerable<TDataItem> data)
        {
            _data = data;
            return this;
        }

        public SheetConfigurator<TDataItem> SetTitle(string title)
        {
            _title = title;
            return this;
        }
    }
}