using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Exceptions
{
    public class DuplicatedSheetNameException : ApplicationException
    {
        public DuplicatedSheetNameException() : base() {}


        public DuplicatedSheetNameException(string sheetName) : base()
        {
            DuplicatedName = sheetName;
        }

        public String DuplicatedName { get; set; }
    }
}
