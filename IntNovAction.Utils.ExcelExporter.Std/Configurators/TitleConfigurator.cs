using System;

namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    public class TitleConfigurator
    {
        internal FormatConfigurator _Format { private set; get; }
        internal string _TitleText { private set; get; }

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