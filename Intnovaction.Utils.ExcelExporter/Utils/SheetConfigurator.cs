
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// Contiene los datos de una hoja
    /// </summary>
    public class SheetConfigurator<TDataItem> : SheetConfiguratorBase
        where TDataItem: new()
    {

        public SheetConfigurator()
        {
            _fontFormatters = new List<Tuple<Func<TDataItem, bool>, FontFormat>>();
            _fontSizeFormatters = new List<Tuple<Func<TDataItem, bool>, int>>();
        }

        internal bool _hideHeaders = false;
       

        internal List<Tuple<Func<TDataItem, bool>, FontFormat>> _fontFormatters;

        internal List<Tuple<Func<TDataItem, bool>, int>> _fontSizeFormatters;

        internal IEnumerable<TDataItem> _data;

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


        /// <summary>
        /// Añade un formato condicional a una celda
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> AddFormat(Func<TDataItem, bool> condition, FontFormat format)
        {
            _fontFormatters.Add(new Tuple<Func<TDataItem, bool>, FontFormat>(condition, format));
            return this;
        }

        /// <summary>
        /// Añade un formato condicional a una celda
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> AddFormat(Func<TDataItem, bool> condition, int size)
        {
            _fontSizeFormatters.Add(new Tuple<Func<TDataItem, bool>, int>(condition, size));
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


        public SheetConfigurator<TDataItem> SetTitle(string title)
        {
            _title = title;
            return this;
        }



    }
}
