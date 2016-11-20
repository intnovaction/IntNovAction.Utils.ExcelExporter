using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.FormatExporters
{
    internal class PropInfo
    {
        public string DisplayName { get; set; }
        public int Order { get; set; }
        public PropertyInfo PropertyInfo { get; set; }
    }
}