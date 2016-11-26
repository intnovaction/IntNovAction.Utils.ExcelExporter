using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    public class TitleConfigurator
    {
        internal string _TitleText;
        internal FormatConfigurator _Format;

        public TitleConfigurator Text(string text)
        {
            _TitleText = text;
            return this;
        }

        public TitleConfigurator Format(Action<FormatConfigurator> formatConfigurator)
        {
            _Format = new FormatConfigurator();
            formatConfigurator.Invoke(_Format);

            return this;
        }
    }
}
