using IntNovAction.Utils.ExcelExporter.Exceptions;
using System.Collections.Generic;
using System.Linq;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    internal class SheetCollection : List<SheetConfiguratorBase>
    {
        /// <summary>
        /// Añade una hoja con datos
        /// </summary>
        /// <param name="dato"></param>
        public new void Add(SheetConfiguratorBase dato)
        {
            if (this.Any(p => p._name == dato._name))
            {
                throw new DuplicatedSheetNameException(dato._name);
            }

            var sheetNumber = this.Count() + 1;
            string sheetName = $"Sheet {sheetNumber}";

            base.Add(dato);
        }
    }
}