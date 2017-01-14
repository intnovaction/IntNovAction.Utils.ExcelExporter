using System;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    public class TitleConfigurator
    {
        internal FormatConfigurator _Format;
        internal string _TitleText;

        public TitleConfigurator Format(Action<FormatConfigurator> formatConfigurator)
        {
            _Format = new FormatConfigurator();
            formatConfigurator.Invoke(_Format);

            return this;
        }

        public TitleConfigurator Text(string text)
        {
            _TitleText = text;
            return this;
        }
    }
}