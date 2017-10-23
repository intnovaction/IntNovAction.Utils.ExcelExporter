using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class PrintAreaTests
    {
        [TestMethod]
        [TestCategory(Categories.PrintArea)]
        public void If_I_set_PrintInOnePage_It_should_honor_coordinates()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items)
                    .SetCoordinates(3, 3)
                    .Name(sheetName)
                    .Title(t => t.Format(f => f.Bold()))
                    .PrintInOnePage()
                );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.PageSetup.PrintAreas.Should().NotBeNullOrEmpty();

                var printArea = firstSheet.PageSetup.PrintAreas.First();

                printArea.FirstColumnUsed().ColumnNumber().Should().Be(3);
                printArea.LastColumnUsed().ColumnNumber().Should().Be(5);

                printArea.FirstRowUsed().RowNumber().Should().Be(3);
                printArea.LastRowUsed().RowNumber().Should().Be(7, "Title + Headers + 3 data");
            }
        }

        [TestMethod]
        [TestCategory(Categories.PrintArea)]
        public void If_I_set_PrintInOnePage_It_should_include_only_used_cells()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items)
                    .Name(sheetName)
                    .Title(t => t.Format(f => f.Bold()))
                    .PrintInOnePage()
                );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.PageSetup.PrintAreas.Should().NotBeNullOrEmpty();

                var printArea = firstSheet.PageSetup.PrintAreas.First();

                printArea.FirstColumnUsed().ColumnNumber().Should().Be(1);
                printArea.LastColumnUsed().ColumnNumber().Should().Be(3);

                printArea.FirstRowUsed().RowNumber().Should().Be(1);
                printArea.LastRowUsed().RowNumber().Should().Be(5, "Title + Headers + 3 data");
            }
        }
    }
}