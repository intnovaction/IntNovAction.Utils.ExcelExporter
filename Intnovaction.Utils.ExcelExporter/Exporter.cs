
using IntNovAction.Utils.ExcelExporter.Utils;
using IntNovAction.Utils.FormatExporters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter
{
    /// <summary>
    /// Genera un excel con los elementos de tipo ListItem como filas
    /// </summary>
    /// <typeparam name="ListItem">El elemento que vamos a crear</typeparam>
    public class Exporter<ListItem>
        where ListItem : new()
    {

        internal SheetCollection<ListItem> _data;
        internal List<Tuple<Func<ListItem, bool>, FontFormat>> _fontFormatters;
        internal List<Tuple<Func<ListItem, bool>, int>> _fontSizeFormatters;

        internal bool _hideHeaders = false;
        private string _title;
        private Type itemType;
        internal bool _jumpHeaders;

        public Exporter()
        {
            itemType = typeof(ListItem);
            _data = new SheetCollection<ListItem>();

            _fontFormatters = new List<Tuple<Func<ListItem, bool>, FontFormat>>();
            _fontSizeFormatters = new List<Tuple<Func<ListItem, bool>, int>>();
        }
        /// <summary>
        /// Añade un formato condicional a una celda
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public Exporter<ListItem> AddFormat(Func<ListItem, bool> condition, FontFormat format)
        {
            _fontFormatters.Add(new Tuple<Func<ListItem, bool>, FontFormat>(condition, format));
            return this;
        }

        /// <summary>
        /// Añade un formato condicional a una celda
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public Exporter<ListItem> AddFormat(Func<ListItem, bool> condition, int size)
        {
            _fontSizeFormatters.Add(new Tuple<Func<ListItem, bool>, int>(condition, size));
            return this;
        }

        public byte[] Export()
        {
            var excelExporter = GetFormatter();

            var elems = excelExporter.Export(this);

            return elems;
        }

        /// <summary>
        /// Indica que no se muestren los headers
        /// </summary>
        /// <returns></returns>
        public Exporter<ListItem> HideHeaders()
        {
            _hideHeaders = true;
            return this;
        }

        public Exporter<ListItem> JumpHeaders()
        {
            _jumpHeaders = true;
            return this;
        }


        public Exporter<ListItem> SetData(IEnumerable<ListItem> data)
        {
            _data.Add(data);
            return this;
        }

        public Exporter<ListItem> SetData(string name, IEnumerable<ListItem> data)
        {
            _data.Add(name, data);
            return this;
        }
        public Exporter<ListItem> SetTitle(string title)
        {

            _title = title;
            return this;
        }
        BaseExporter<ListItem> GetFormatter()
        {
            return new Excel<ListItem>();
        }


    }


}
