using System.ComponentModel.DataAnnotations;

namespace IntNovAction.Utils.ExcelExporter.Tests.TestObjects
{
    public class TestListItem
    {
        public const string PropATitle = "Pepe";

        [Display(Name = PropATitle, Order = 2)]
        public string PropA { get; set; }

        [Display(Order = 33)]
        public string PropB { get; set; }

        public int PropC { get; set; }
    }
}