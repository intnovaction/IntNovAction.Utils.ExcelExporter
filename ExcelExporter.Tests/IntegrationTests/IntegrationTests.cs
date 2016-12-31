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

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class ExporterTest
    {
        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void AddData()
        {
            var exporter = new Exporter();

            var items = IntegrationTestsUtils.GenerateItems(3);

            exporter.AddSheet<TestListItem>(c =>
                c.SetData(items).Name("Hoja 1")
            );
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void Export()
        {
            var sheetTitle = "1-Sheet";
            var items = IntegrationTestsUtils.GenerateItems(3);

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
        public void ExportWithFormat()
        {
            var sheetTitle = "1-Sheet";
            var items = IntegrationTestsUtils.GenerateItems(1);

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
        public void MultipleDataSheet()
        {
            var sheet1Name = "Hoja 1";
            var sheet2Name = "Hoja 2";
            var sheet1Rows = 2;
            var sheet2Rows = 3;

            var items = IntegrationTestsUtils.GenerateItems(sheet1Rows);
            var items2 = IntegrationTestsUtils.GenerateItems(sheet2Rows);

            var exporter = new Exporter()
                .AddSheet<TestListItem>(c => c.SetData(items).Name(sheet1Name))
                .AddSheet<TestListItem>(c => c.SetData(items2).Name(sheet2Name));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {

                var workbook = new XLWorkbook(stream);

                workbook.Worksheets.Count.Should().Be(2);

                workbook.Worksheets.Worksheet(1).Name.Should().Be(sheet1Name);
                workbook.Worksheets.Worksheet(1).LastRowUsed().RowNumber().Should().Be(sheet1Rows + 1);


                workbook.Worksheets.Worksheet(2).Name.Should().Be(sheet2Name);
                workbook.Worksheets.Worksheet(2).LastRowUsed().RowNumber().Should().Be(sheet2Rows + 1);

            }
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void UseExistingExcel()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var excelFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("IntNovAction.Utils.ExcelExporter.Tests.Test.xlsx");

            var exporter = new Exporter(excelFileStream)
               .AddSheet<TestListItem>(c => c.SetData(items));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {

                var sheetName = "ExistingHoja1";
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

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void SetCoordinates()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).SetCoordinates(3, 2));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(1, 1).Value.Should().Be(string.Empty);
                firstSheet.Cell(3, 2).Value.Should().NotBeNull();
            }
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void HideColumnHeaders()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).HideColumnHeaders());

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastRowUsed().RowNumber().Should().Be(items.Count);

            }
        }

    }
}