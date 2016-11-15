using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    internal class SheetCollection<ListItem> : Dictionary<SheetInfo, IEnumerable<ListItem>>
    {

        /// <summary>
        /// Añade una hoja con datos
        /// </summary>
        /// <param name="dato"></param>
        public void Add(IEnumerable<ListItem> datos)
        {
            var sheetNumber = this.Count() + 1;
            string sheetName = $"Sheet {sheetNumber}";

            Add(sheetName, datos);

        }


        /// <summary>
        /// Añade una hoja con datos
        /// </summary>
        /// <param name="dato"></param>
        public void Add(string sheetName, IEnumerable<ListItem> datos)
        {
            var sheetNumber = this.Count() + 1;

            var info = new SheetInfo()
            {
                Name = sheetName,
                Order = sheetNumber
            };

            this.Add(info, datos);
        }
    }
}
