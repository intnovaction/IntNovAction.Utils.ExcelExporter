namespace IntNovAction.Utils.ExcelExporter.Utils
{

    /// <summary>
    /// Configura un formato para un rango de celdas
    /// </summary>
    public class FormatConfigurator
    {
        internal bool _bold = false;
        internal int? _fontSize = null;
        internal bool _italic = false;
        internal bool _underline = false;

        public FormatConfigurator Bold()
        {
            _bold = true;
            return this;
        }

        public FormatConfigurator FontSize(int fontSize)
        {
            _fontSize = fontSize;
            return this;
        }

        public FormatConfigurator Italic()
        {
            _italic = true;
            return this;
        }

        public FormatConfigurator Underline()
        {
            _underline = true;
            return this;
        }
    }
}