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
    public class ColumnsConfigTests
    {
        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void If_I_hide_a_colum_It_should_be_hidden()
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
        public void If_I_set_a_column_format_Only_rows_should_be_formatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            Action<FormatConfigurator> firstColumnFormat = (f) => f.Bold().Color(255, 0, 0);
            Action<FormatConfigurator> secondColumnFormat = (f) => f.Italic();

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA).DataFormat(firstColumnFormat);
                        cols.AddColumn(prop => prop.PropB).DataFormat(secondColumnFormat);
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                // Titulo
                FormatChecker.CheckFormat(firstSheet.Cell(1, 1), DefaultStyles.GetCellDefaultStyle());
                FormatChecker.CheckFormat(firstSheet.Cell(1, 2), DefaultStyles.GetCellDefaultStyle());

                // Cabecera
                FormatChecker.CheckFormat(firstSheet.Cell(2, 1), firstColumnFormat);
                FormatChecker.CheckFormat(firstSheet.Cell(2, 2), secondColumnFormat);

                for (int i = 3; i <= 5; i++)
                {
                    FormatChecker.CheckFormat(firstSheet.Cell(i, 1), firstColumnFormat);
                    FormatChecker.CheckFormat(firstSheet.Cell(i, 2), secondColumnFormat);
                }

                // Fila de despues
                FormatChecker.CheckFormat(firstSheet.Cell(6, 1), DefaultStyles.GetCellDefaultStyle());
                FormatChecker.CheckFormat(firstSheet.Cell(6, 2), DefaultStyles.GetCellDefaultStyle());
            }
        }

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void If_I_set_a_column_title_and_transform_It_should_be_honored()
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
        public void If_I_set_a_column_with_It_should_be_honored()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA).DataFormat(f => f.Width(150));
                        cols.AddColumn(prop => prop.PropB).DataFormat(f => f.Width(10));
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

        [TestMethod]
        [TestCategory(Categories.ColumnsConfig)]
        public void If_I_set_column_titles_They_must_be_shown()
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
                        cols.AddColumn(prop => prop.PropA).Header("PropC Inc");
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
        public void If_I_specify_columns_Only_those_should_be_shown()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items).Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA);
                        cols.AddColumn(prop => prop.PropA).Header("Prop a (2)");
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
        public void If_I_set_header_format_They_must_be_honored()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter().AddSheet<TestListItem>(sheet =>
                sheet.SetData(items)
                    .Name(sheetName)
                    .Columns(cols =>
                    {
                        cols.Clear();
                        cols.AddColumn(prop => prop.PropA).Header(t => t.Format(f => f.Bold()));
                        cols.AddColumn(prop => prop.PropA).Header(t => t.Text("PropC Inc").Format(f => f.Italic()));
                    })
            );

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);
                var firstSheet = workbook.Worksheets.Worksheet(1);

                firstSheet.LastColumnUsed().ColumnNumber().Should().Be(2);

                firstSheet.Cell(1, 1).Value.Should().Be(TestListItem.PropATitle);
                firstSheet.Cell(1, 1).Style.Font.Bold.Should().Be(true, "should be bold");

                firstSheet.Cell(1, 2).Value.Should().Be("PropC Inc");
                firstSheet.Cell(1, 2).Style.Font.Italic.Should().Be(true, "should be italic");
            }
        }


    }
}