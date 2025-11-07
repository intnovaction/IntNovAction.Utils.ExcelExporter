using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace IntNovAction.Utils.ExcelExporter.ExcelWriters
{
    public class ExcelFileWriter
    {
        public byte[] Export(Exporter exporter)
        {
            XLWorkbook workbook = null;

            try
            {
                if (exporter._existingFileStream != null)
                {
                    workbook = new XLWorkbook(exporter._existingFileStream);
                }
                else
                {
                    workbook = new XLWorkbook();
                }

                var orderedSheets = exporter._sheets.OrderBy(p => p._order).ToList();

                for (int i = 0; i < orderedSheets.Count; i++)
                {
                    var sheetConfigurator = orderedSheets[i];

                    var itemType = sheetConfigurator.GetType().GetGenericArguments().First();

                    var newType = typeof(ExcelSheetGenerator<>);
                    var genericType = newType.MakeGenericType(itemType);

                    var sheetExporter = Activator.CreateInstance(genericType);

                    genericType.InvokeMember("WriteSheet",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public,
                        Type.DefaultBinder, sheetExporter, new object[] { workbook, sheetConfigurator });
                }

                workbook.Properties.Created = DateTime.Now;
                workbook.Properties.Comments = "Created with IntNovAction ExcelExporter";

                using (var ms = new MemoryStream())
                {
                    workbook.CalculationOnSave = false;
                    workbook.FullCalculationOnLoad = true;
                    workbook.SaveAs(ms);
                    return ms.ToArray();
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Dispose();
                }
            }
        }
    }
}