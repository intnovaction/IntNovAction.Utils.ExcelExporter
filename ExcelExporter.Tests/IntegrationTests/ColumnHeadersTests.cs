using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using System.IO;
using ClosedXML.Excel;
using FluentAssertions;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class ColumnHeadersTests
    {

        [TestMethod]
        [TestCategory(Categories.ColumnHeaders)]
        public void If_I_hide_headers_They_should_not_be_shown()
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
