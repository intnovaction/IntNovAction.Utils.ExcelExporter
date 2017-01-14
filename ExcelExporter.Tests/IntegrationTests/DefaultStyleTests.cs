using ClosedXML.Excel;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter.Tests.Utils;
using IntNovAction.Utils.ExcelExporter.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    [TestClass]
    public class DefaultStyleTests
    {
        [TestMethod]
        [TestCategory(Categories.DefaultFormat)]
        public void If_I_apply_DefaultStyles_Title_and_headers_should_be_formatted()
        {
            var items = IntegrationTestsUtils.GenerateItems(3);

            var sheetName = "Hoja 1";

            var exporter = new Exporter()
               .AddSheet<TestListItem>(
                    c => c.SetData(items)
                        .ApplyDefaultStyles()
                        .Name(sheetName)
                        .Title());

            var result = exporter.Export();

            using (var stream = new MemoryStream(result))
            {
                var workbook = new XLWorkbook(stream);

                var firstSheet = workbook.Worksheets.Worksheet(1);

                var titleDefaultStyle = DefaultStyles.GetTitleDefaultStyle();
                FormatChecker.CheckFormat(firstSheet.Cell(1, 1), titleDefaultStyle);

                var headerDefaultStyle = DefaultStyles.GetHeadersDefaultStyle();
                for (int i = firstSheet.FirstColumnUsed().ColumnNumber();
                    i < firstSheet.LastColumnUsed().ColumnNumber(); i++)
                {
                    FormatChecker.CheckFormat(firstSheet.Cell(2, i), headerDefaultStyle);
                }
            }
        }
    }
}