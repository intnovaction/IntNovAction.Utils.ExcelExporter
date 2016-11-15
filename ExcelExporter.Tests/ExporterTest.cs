using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using FluentAssertions;

namespace ExcelExporter.Tests
{
    [TestClass]
    public class ExporterTest
    {
        [TestMethod]
        public void TestAdd()
        {

            var exporter = new Exporter<TestListItem>();

            var items = GenerateItems(3);
            exporter.SetData(items);

            var item2 = GenerateItems(3);
            exporter.SetData(item2);


        }


        [TestMethod]
        public void TestAdd2()
        {

            var exporter = new Exporter<TestListItem>();

            var items = GenerateItems(3);
            exporter.SetData("1-Sheet", items);

            var item2 = GenerateItems(3);
            exporter.SetData(item2);


        }

        [TestMethod]
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

            System.IO.File.WriteAllBytes(@"d:\pp.xlsx", result);

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                workbook.Worksheets.Count.Should().Be(1);
                var firstSheet = workbook.Worksheets.Worksheet(1);
                firstSheet.Name.Should().Be(sheetTitle);

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
