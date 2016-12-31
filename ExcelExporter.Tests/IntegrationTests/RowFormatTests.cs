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
    public class RowFormatTests
    {

        [TestMethod]
        [TestCategory(Categories.RowFormat)]
        public void FormatRow_RowShouldBeFormatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).HideColumnHeaders()
                .AddFormatRule(p => p.PropC == 0, f => f.Bold().Italic().Underline().Color(255, 0, 0))
               );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                var cellA1 = firstSheet.Row(1).Cell(1);

                cellA1.Style.Font.Underline.Should().Be(XLFontUnderlineValues.Single);
                cellA1.Style.Font.Bold.Should().Be(true);
                cellA1.Style.Font.Italic.Should().Be(true);
                cellA1.Style.Font.FontColor.Color.ToHex().Should().Be("FFFF0000");

            }

        }

        [TestMethod]
        [TestCategory(Categories.RowFormat)]
        public void FormatRow_RowShouldNotBeFormatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(c => c.SetData(items).Name(sheetName).HideColumnHeaders()
                .AddFormatRule(p => p.PropC == 0, f => f.Bold().Italic().Underline().Color(255, 0, 0))
               );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);


                var cellA2 = firstSheet.Row(2).Cell(1);

                cellA2.Style.Font.Underline.Should().Be(XLFontUnderlineValues.None);
                cellA2.Style.Font.Bold.Should().Be(false);
                cellA2.Style.Font.Italic.Should().Be(false);
                cellA2.Style.Font.FontColor.Color.ToHex().Should().Be("FF000000");
            }

        }
    }
}
