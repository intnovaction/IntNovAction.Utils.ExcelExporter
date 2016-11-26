namespace IntNovAction.Utils.ExcelExporter.Utils
{

    /// <summary>
    /// Configura un formato para un rango de celdas
    /// </summary>
    public class FormatConfigurator
    {
        internal bool _bold = false;
        internal ColorInfo _color;
        internal int? _fontSize = null;
        internal bool _italic = false;
        internal bool _underline = false;

        internal class ColorInfo
        {
            public ColorInfo(int red, int green, int blue)
            {
                Red = red;
                Green = green;
                Blue = blue;
            }

            public int Red { get; private set; }
            public int Green { get; private set; }
            public int Blue { get; private set; }
        }

        public FormatConfigurator Color(int red, int green, int blue)
        {
            _color = new ColorInfo(red, green, blue);
            
            return this;
        }

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