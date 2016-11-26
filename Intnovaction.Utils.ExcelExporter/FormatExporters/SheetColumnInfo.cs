using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.FormatExporters
{
    internal class SheetColumnInfo
    {
        /// <summary>
        /// Nombre a mostrar
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Orden en el que se muestra
        /// </summary>
        public int Order { get; set; }

        /// <summary>
        /// Propiedad de la clase correspondiente a la columna
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }
    }
}