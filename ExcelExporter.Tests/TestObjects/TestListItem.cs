using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Tests.TestObjects
{
    public class TestListItem
    {
        [Display(Name = "Pepe", Order = 2)]
        public string PropA { get; set; }

        [Display(Order = 33)]
        public string PropB { get; set; }

        public int PropC { get; set; }
    }
}
