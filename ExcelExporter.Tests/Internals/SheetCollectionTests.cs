using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Exceptions;
using IntNovAction.Utils.ExcelExporter.Tests.TestObjects;
using IntNovAction.Utils.ExcelExporter.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace IntNovAction.Utils.ExcelExporter.Tests.Internals
{
    [TestClass]
    public class SheetCollectionTests
    {
        [TestCategory(Categories.SheetCollection)]
        [TestMethod]
        public void DuplicatedSheetName()
        {
            var sheetCollection = new Utils.SheetCollection();

            var name = "Sheet1";

            var sheetConfig1 = new SheetConfigurator<TestListItem>();
            sheetConfig1.Name(name);

            var duplicateNameSheetConfig = new SheetConfigurator<TestListItem>();
            duplicateNameSheetConfig.Name(name);

            sheetCollection.Add(sheetConfig1);

            Action action = () => sheetCollection.Add(duplicateNameSheetConfig);

            action.ShouldThrow<DuplicatedSheetNameException>()
                .Where(p => p.DuplicatedName == name);
        }
    }
}