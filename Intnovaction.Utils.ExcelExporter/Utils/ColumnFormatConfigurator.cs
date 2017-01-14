namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// Configurador del formato de una columna.
    /// Se aplica ese formato antes del de las filas, que tiene prioridad
    /// </summary>
    public class ColumnFormatConfigurator : FormatConfigurator
    {
        internal double? _width;

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