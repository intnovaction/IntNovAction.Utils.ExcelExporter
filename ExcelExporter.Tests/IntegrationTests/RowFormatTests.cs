using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter.Tests.Utils;
using IntNovAction.Utils.ExcelExporter.Utils;
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

        [TestMethod]
        [TestCategory(Categories.RowFormat)]
        public void FormatRow_RowShouldNotBeFormatted()
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


                var fornFormat = firstSheet.Row(2).Cell(1).Style.Font;
                fornFormat.Underline.Should().Be(XLFontUnderlineValues.None);
                fornFormat.Bold.Should().Be(false);
                fornFormat.Italic.Should().Be(false);
                fornFormat.FontColor.Color.ToHex().Should().Be("FF000000");
            }

        }
    }
}
