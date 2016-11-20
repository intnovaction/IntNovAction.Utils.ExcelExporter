using IntNovAction.Utils.ExcelExporter.Exceptions;
using System.Collections.Generic;
using System.Linq;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// La colección de hojas dentro del excel
    /// </summary>
    internal class SheetCollection : List<SheetConfiguratorBase>
    {
        /// <summary>
        /// Añade una hoja con datos
        /// </summary>
        /// <param name="sheetConfigurator">El configurador de la hoja</param>
        public new void Add(SheetConfiguratorBase sheetConfigurator)
        {
            if (this.Any(p => p._name == sheetConfigurator._name))
            {
                throw new DuplicatedSheetNameException(sheetConfigurator._name);
            }

            var sheetNumber = this.Count() + 1;
            string sheetName = $"Sheet {sheetNumber}";

            base.Add(sheetConfigurator);
        }
    }
}