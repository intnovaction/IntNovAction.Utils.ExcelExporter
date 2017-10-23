using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Configurators;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter.Tests.Utils;
using IntNovAction.Utils.ExcelExporter.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class RowFormatTests
    {
        [TestMethod]
        [TestCategory(Categories.RowFormat)]
        public void If_I_do_not_format_a_row_It_should_not_be_formatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            Action<FormatConfigurator> formatAct = (f) =>
               f.Bold()
              .Italic()
              .Underline()
              .Color(255, 0, 0);

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).HideColumnHeaders()
                .AddFormatRule(p => p.PropC == 0, formatAct)
               );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                var fontFormat = firstSheet.Row(2).Cell(1).Style.Font;
                fontFormat.Underline.Should().Be(XLFontUnderlineValues.None);
                fontFormat.Bold.Should().Be(false);
                fontFormat.Italic.Should().Be(false);
                fontFormat.FontColor.Color.ToHex().Should().Be("FF000000");
            }
        }

        [TestMethod]
        [TestCategory(Categories.RowFormat)]
        public void If_I_format_a_row_Only_that_row_should_be_formatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            Action<FormatConfigurator> formatAct = (f) =>
             f.Bold()
            .Italic()
            .Underline()
            .Color(255, 0, 0);

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).HideColumnHeaders()
                .AddFormatRule(p => p.PropC == 0, formatAct)
               );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                var cellA1 = firstSheet.Row(1).Cell(1);

                FormatChecker.CheckFormat(cellA1, formatAct);
            }
        }
    }
}