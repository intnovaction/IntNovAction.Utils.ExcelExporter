using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        public void If_I_add_data_It_should_not_fail()
        {
            var exporter = new Exporter();

            var items = IntegrationTestsUtils.GenerateItems(3);

            exporter.AddSheet<TestListItem>(c =>
                c.SetData(items).Name("Hoja 1")
            );
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void If_I_export_multiple_datasets_All_sheets_should_be_created()
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
        public void If_I_export_Mustshow_SheetTitle_and_headers_and_all_data_should_be_exported()
        {
            var sheetTitle = "1-Sheet";
            var items = IntegrationTestsUtils.GenerateItems(300);

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
        public void If_I_set_a_condtionalFormat_It_should_be_honored()
        {
            var sheetTitle = "1-Sheet";
            var items = IntegrationTestsUtils.GenerateItems(1);

            var exporter = new Exporter()
                .AddSheet<TestListItem>(c =>
                    c.SetData(items).Name(sheetTitle)
                        .AddFormatRule(p => p.PropA == items.First().PropA,
                            f => f.Bold().Italic()));

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

                firstSheet.Cell(3, 1).Style.Font.Bold.Should().BeFalse();
                firstSheet.Cell(3, 1).Style.Font.Italic.Should().BeFalse();
                firstSheet.Cell(3, 1).Style.Font.Underline.Should().Be(XLFontUnderlineValues.None);
            }
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void If_I_set_coordinates_Data_should_start_in_place()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items)
                .Name(sheetName)
                .SetCoordinates(3, 2));

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Cell(1, 1).GetValue<string>().Should().Be(string.Empty);
                firstSheet.Cell(3, 2).Value.Should().NotBeNull();
            }
        }

        [TestMethod]
        [TestCategory(Categories.Integration)]
        public void If_I_use_an_existing_excel_It_should_use_the_origina_sheet()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var excelFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("IntNovAction.Utils.ExcelExporterTests.Test.xlsx");

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
        public void If_I_add_a_custom_cell_it_should_be_added()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(
                    c => c.SetData(items)
                          .Name(sheetName)
                          .SetCustomContent(2, 1, "Custom Cell Content")
                          .SetCustomContent(3, 1, "Formatted Content", format => format.Bold().Italic())
                          .SetCoordinates(4, 1)
               );

            var result = exporter.Export();

            // Guardar el archivo para inspección visual
            var outputPath = Path.Combine(Path.GetTempPath(), "CustomCellTest.xlsx");
            File.WriteAllBytes(outputPath, result);
            System.Diagnostics.Debug.WriteLine($"Excel guardado en: {outputPath}");

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(sheetName);

                // Verificar celda sin formato
                firstSheet.Cell(2, 1).GetValue<string>().Should().Be("Custom Cell Content");

                // Verificar celda con formato
                firstSheet.Cell(3, 1).GetValue<string>().Should().Be("Formatted Content");
                firstSheet.Cell(3, 1).Style.Font.Bold.Should().BeTrue();
                firstSheet.Cell(3, 1).Style.Font.Italic.Should().BeTrue();

                // Verificar que los datos empiezan en la fila 4
                firstSheet.Cell(4, 1).Value.Should().NotBeNull();
            }
        }
    }
}