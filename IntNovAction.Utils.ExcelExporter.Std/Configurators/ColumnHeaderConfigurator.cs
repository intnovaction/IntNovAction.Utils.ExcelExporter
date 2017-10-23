using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    public class ColumnHeaderConfigurator
    {
        internal string _text { private set; get; }

        internal FormatConfigurator _Format;

        public ColumnHeaderConfigurator Text(string text)
        {
            _text = text;
            return this;
        }

        public ColumnHeaderConfigurator Format(Action<FormatConfigurator> formatConfigurator)
        {
            _Format = new FormatConfigurator();
            formatConfigurator.Invoke(_Format);

            return this;
        }



    }
}
