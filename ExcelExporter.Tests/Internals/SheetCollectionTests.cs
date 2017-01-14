using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Configurators;
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
        public void If_I_add_a_duplicated_sheet_name_It_should_raise_an_exception()
        {
            var sheetCollection = new SheetCollection();

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