using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    public abstract class SheetConfiguratorBase
    {

        internal string _name { get; set; }

        internal int _order { get; set; }

        internal string _title;

        internal bool _jumpHeaders;
    }
}
