using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using FluentAssertions;
using System.Reflection;

namespace ExcelExporter.Tests
{
    [TestClass]
    public class ExporterTest
    {
        [TestMethod]
        [TestCategory("ExcelExporter")]
        public void TestAddData()
        {

            var exporter = new Exporter<TestListItem>();

            var items = GenerateItems(3);
            exporter.SetData(items);

            var item2 = GenerateItems(3);
            exporter.SetData(item2);


        }


        [TestMethod]
        [TestCategory("ExcelExporter")]
        public void TestAddDataNamedSheet()
        {

            var exporter = new Exporter<TestListItem>();

            var items = GenerateItems(3);
            exporter.SetData("1-Sheet", items);

            var item2 = GenerateItems(3);
            exporter.SetData(item2);


        }

        [TestMethod]
        [TestCategory("ExcelExporter")]
        public void TestExport()
        {


            var sheetTitle = "1-Sheet";
            var items = GenerateItems(3);

            var exporter = new Exporter<TestListItem>()
                .SetData(sheetTitle, items)
                .AddFormat(p => p.PropC == 3, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Bold)
                .AddFormat(p => p.PropC == 2, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Italic)
                .AddFormat(p => p.PropC == 1, 20);


            var result = exporter.Export();


            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                workbook.Worksheets.Count.Should().Be(1);
                var firstSheet = workbook.Worksheets.Worksheet(1);
                firstSheet.Name.Should().Be(sheetTitle);

                firstSheet.LastRowUsed().RowNumber()
                    .Should().Be(items.Count + 1, $"Hay {items.Count} datos y una mas de cabecera");

            }

        }

        [TestMethod]
        [TestCategory("ExcelExporter")]
        public void TestUseExistingExcel()
        {
            var items = GenerateItems(3);

            var excelFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("IntNovAction.Utils.ExcelExporter.Tests.Test.xlsx");

            var exporter = new Exporter<TestListItem>(excelFileStream)
                .SetData(items)
                .JumpHeaders()
                .AddFormat(p => p.PropC == 3, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Bold)
                .AddFormat(p => p.PropC == 2, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Italic)
                .AddFormat(p => p.PropC == 1, 20);

            var result = exporter.Export();

            //System.IO.File.WriteAllBytes(@"d:\pp.xlsx", result);

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                workbook.Worksheets.Count.Should().Be(1);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Name
                    .Should()
                    .Be("Hoja1", "el nombre de la hoja 1 en el excel de ejemplo es Hoja 1");

                firstSheet.LastRowUsed().RowNumber()
                    .Should().Be(items.Count + 1, $"el excel de ejemplo tiene cabecera y hay {items.Count} datos");
            }

        }


        private List<TestListItem> GenerateItems(int numItems)
        {

            var result = new List<TestListItem>();
            for (int i = 0; i < numItems; i++)
            {
                result.Add(new TestListItem()
                {
                    PropA = $"PropA - {i}",
                    PropB = $"PropB - {i}",
                    PropC = i,
                });
            }

            return result;
        }
    }
}
