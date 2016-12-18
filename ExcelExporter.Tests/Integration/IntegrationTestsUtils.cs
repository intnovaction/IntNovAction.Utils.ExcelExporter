using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Tests.Integration
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
