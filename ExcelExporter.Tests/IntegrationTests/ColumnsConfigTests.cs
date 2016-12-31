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
    public class ColumnsConfigTests
    {

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_ExplicitColumns()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA);
                        cols.AddColumn(prop => prop.PropA).Title("Prop a (2)");
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastColumnUsed().ColumnNumber().Should().Be(2);
                firstSheet.Cell(1, 1).Value.Should().Be(TestListItem.PropATitle);
                firstSheet.Cell(1, 2).Value.Should().Be("Prop a (2)");
            }
        }



        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_ExplicitColumnsWithTitle()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items)
                    .Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA);
                        cols.AddColumn(prop => prop.PropA).Title("PropC Inc");
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastColumnUsed().ColumnNumber().Should().Be(2);
                firstSheet.Cell(1, 1).Value.Should().Be(TestListItem.PropATitle);
                firstSheet.Cell(1, 2).Value.Should().Be("PropC Inc");
            }
        }

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_ExplicitColumnsWithTransform()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA);
                        cols.AddColumnExpr(prop => prop.PropC + 1, "Plus 2");
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);



                firstSheet.Cell(1, 1).Value.Should().Be(TestListItem.PropATitle);
                firstSheet.Cell(1, 2).Value.Should().Be("Plus 2");

                for (int excelRow = 2; excelRow <= items.Count + 1; excelRow++)
                {
                    var originalItem = items[excelRow - 2];

                    var secondValue = firstSheet.Cell(excelRow, 2).Value;
                    secondValue.CastTo<int>().Should().Be(originalItem.PropC + 1);
                }

                firstSheet.LastColumnUsed().ColumnNumber().Should().Be(2);
                firstSheet.LastRowUsed().RowNumber().Should().Be(items.Count + 1);
            }
        }

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_HideColumns()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.HideColumn(prop => prop.PropB);
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastColumnUsed().ColumnNumber().Should().Be(2);
                firstSheet.Cell(1, 1).Value.Should().Be(TestListItem.PropATitle);
                firstSheet.Cell(1, 2).Value.Should().Be(nameof(TestListItem.PropC));
            }
        }

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_FormatColumn()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA).Format(f => f.Bold().Color(255, 0, 0));
                        cols.AddColumn(prop => prop.PropB).Format(f => f.Italic());
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                var cellStyle = firstSheet.Cell(2, 1).Style;
                cellStyle.Font.Bold.Should().BeTrue();
                cellStyle.Font.Italic.Should().BeFalse();

                cellStyle = firstSheet.Cell(3, 1).Style;
                cellStyle.Font.Bold.Should().BeTrue();
                cellStyle.Font.Italic.Should().BeFalse();

                cellStyle = firstSheet.Cell(2, 2).Style;
                cellStyle.Font.Bold.Should().BeFalse();
                cellStyle.Font.Italic.Should().BeTrue();

                cellStyle = firstSheet.Cell(3, 2).Style;
                cellStyle.Font.Bold.Should().BeFalse();
                cellStyle.Font.Italic.Should().BeTrue();
            }
        }


        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void ColumnConfig_ColumnWidth()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA).Format(f => f.Width(150));
                        cols.AddColumn(prop => prop.PropB).Format(f => f.Width(10));
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.Column(1).Width.Should().Be(150);
                firstSheet.Column(2).Width.Should().Be(10);

            }
        }
    }
}
