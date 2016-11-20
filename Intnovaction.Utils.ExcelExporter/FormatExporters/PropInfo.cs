using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.FormatExporters
{
    internal class PropInfo
    {
        public int Order { get; set; }

        public string DisplayName { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}
