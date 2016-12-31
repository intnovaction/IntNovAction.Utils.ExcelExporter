using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class TitleTests
    {

        [TestMethod]
        [TestCategory(Categories.Title)]
        public void ShowTitleText_Default()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).Title());

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(1, 1).Value.Should().Be(sheetName);
                firstSheet.LastRowUsed().RowNumber().Should().Be(items.Count + 2);
            }
        }


        [TestMethod]
        [TestCategory(Categories.Title)]
        public void ShowTitleText_ExplicitTitle()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";
            var sheetTitle = "Title";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).Title(h => h.Text(sheetTitle)));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(1, 1).Value.Should().Be(sheetTitle);
                firstSheet.LastRowUsed().RowNumber().Should().Be(items.Count + 2);
            }
        }

        [TestMethod]
        [TestCategory(Categories.Title)]
        public void ShowTitleText_Format()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";


            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName)
                .Title(t => t.Format(f => f.Bold()))
                );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(1, 1).Style.Font.Bold.Should().Be(true);
                firstSheet.Cell(2, 1).Style.Font.Bold.Should().Be(false);
                firstSheet.LastRowUsed().RowNumber().Should().Be(items.Count + 2);
            }
        }

        [TestMethod]
        [TestCategory(Categories.Title)]
        public void ShowTitleText_Coordinates()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items)
                    .SetCoordinates(2,3)
                    .Name(sheetName)
                    .Title(t => t.Text(sheetName)));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(2, 3).Value.Should().Be(sheetName);
            }
        }
    }
}
