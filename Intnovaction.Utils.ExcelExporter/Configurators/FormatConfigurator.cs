namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    /// <summary>
    /// Configura un formato para un rango de celdas
    /// </summary>
    public class FormatConfigurator
    {
        internal bool _bold { private set; get; } = false;
        internal bool _bottomBorder { private set; get; } = false;
        internal ColorInfo _color { private set; get; } = null;
        internal int? _fontSize { private set; get; } = null;
        internal bool _italic { private set; get; } = false;
        internal bool _underline { private set; get; } = false;

        public FormatConfigurator Bold()
        {
            return Bold(true);
        }

        public FormatConfigurator Bold(bool value)
        {
            _bold = value;
            return this;
        }

        public FormatConfigurator BottomBorder()
        {
            return BottomBorder(true);
        }

        public FormatConfigurator BottomBorder(bool value)
        {
            _bottomBorder = value;
            return this;
        }

        public FormatConfigurator Color(int red, int green, int blue)
        {
            _color = new ColorInfo(red, green, blue);

            return this;
        }

        public FormatConfigurator FontSize(int fontSize)
        {
            _fontSize = fontSize;
            return this;
        }

        public FormatConfigurator Italic()
        {
            return Italic(true);
        }

        public FormatConfigurator Italic(bool value)
        {
            _italic = value;
            return this;
        }

        public FormatConfigurator Underline()
        {
            _underline = true;
            return this;
        }

        /// <summary>
        /// La info de color a poner
        /// </summary>
        internal class ColorInfo
        {
            public ColorInfo(int red, int green, int blue)
            {
                Red = red;
                Green = green;
                Blue = blue;
            }

            public int Blue { get; private set; }
            public int Green { get; private set; }
            public int Red { get; private set; }
        }
    }
}