using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelExporter.Tests.IntegrationTests
{
    internal static class IntegrationTestsUtils
    {
        public static List<TestListItem> GenerateItems(int numItems)
        {
            var dataToExport = new List<TestListItem>();
            for (int i = 0; i < numItems; i++)
            {
                dataToExport.Add(new TestListItem()
                {
                    PropA = $"PropA - {i}",
                    PropB = $"PropB - {i}",
                    PropC = i,
                });
            }

            return dataToExport;
        }
    }
}