using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter;
using IntNovAction.Utils.ExcelExporter.Tests;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelExporter.Tests.Integration
{
    [TestClass]
    public class ExporterTest
    {
        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void TestAddData()
        {
            var exporter = new Exporter();

            var items = GenerateItems(3);

            exporter.AddSheet<TestListItem>(c =>
                c.SetData(items).Name("Hoja 1")
            );
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void TestExport()
        {
            var sheetTitle = "1-Sheet";
            var items = GenerateItems(3);

            var exporter = new Exporter()
                .AddSheet<TestListItem>(c =>
                    c.SetData(items).Name(sheetTitle));

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
        [TestCategory(Categories.Integration)]
        public void TestExportWithFormat()
        {
            var sheetTitle = "1-Sheet";
            var items = GenerateItems(1);

            var exporter = new Exporter()
                .AddSheet<TestListItem>(c =>
                    c.SetData(items).Name(sheetTitle)
                        .AddFormatRule(p => p.PropA == items.First().PropA, f => f.Bold().Italic()));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastRowUsed().RowNumber()
                    .Should().Be(items.Count + 1, $"Hay {items.Count} datos y una mas de cabecera");

                firstSheet.Cell(2, 1).Style.Font.Bold.Should().BeTrue();
                firstSheet.Cell(2, 1).Style.Font.Italic.Should().BeTrue();
                firstSheet.Cell(2, 1).Style.Font.Underline.Should().Be(XLFontUnderlineValues.None);
            }
        }


        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void TestMultipleDataSheet()
        {
            var items = GenerateItems(3);
            var items2 = GenerateItems(3);

            var exporter = new Exporter()
                .AddSheet<TestListItem>(c => c.SetData(items).Name("Hoja 1"))
                .AddSheet<TestListItem>(c => c.SetData(items2).Name("Hoja 2"));
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void TestUseExistingExcel()
        {
            var items = GenerateItems(3);

            var excelFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("IntNovAction.Utils.ExcelExporter.Tests.Test.xlsx");

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                workbook.Worksheets.Count.Should().Be(1);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Name
                    .Should()
                    .Be(sheetName, "el nombre de la hoja 1 en el excel de ejemplo es Hoja 1");

                firstSheet.LastRowUsed().RowNumber()
                    .Should().Be(items.Count + 1, $"el excel de ejemplo tiene cabecera y hay {items.Count} datos");
            }
        }

        private List<TestListItem> GenerateItems(int numItems)
        {
            var dataToExport = new List<TestListItem>();
            for (int i = 0; i < numItems; i++)
            {
                dataToExport.Add(new TestListItem()
                {
                    PropA = $"PropA - {i}",
                    PropB = $"PropB - {i}",
                    PropC = i,
                });
            }

            return dataToExport;
        }
    }
}