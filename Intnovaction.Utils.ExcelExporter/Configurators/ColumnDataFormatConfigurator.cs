namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    /// <summary>
    /// Configurador del formato de una columna.
    /// Se aplica ese formato antes del de las filas, que tiene prioridad
    /// </summary>
    public class ColumnDataFormatConfigurator : FormatConfigurator
    {
        internal double? _width { get; private set; }

        /// <summary>
        /// Establece el ancho de la columna
        /// </summary>
        /// <param name="width"></param>
        /// <returns></returns>
        public FormatConfigurator Width(double width)
        {
            _width = width;
            return this;
        }
    }
}