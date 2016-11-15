using ClosedXML.Excel;
using IntNovAction.Utils.ExcelExporter;
using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.FormatExporters
{
    internal class Excel<TItem> : BaseExporter<TItem>
        where TItem : new()
    {
        public Excel() : base()
        {

        }

        public override byte[] Export(Exporter<TItem> exporter)
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

                var orderedSheets = exporter._data.OrderBy(p => p.Key.Order).ToList();

                for (int i = 0; i < orderedSheets.Count; i++)
                {
                    var sheetInfo = orderedSheets[i];
                    // Creamos el worksheet
                    ProcessExcelSheet(exporter, workbook, sheetInfo, i);
                }

                using (var ms = new MemoryStream())
                {
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

        private void ProcessExcelSheet(Exporter<TItem> exporter, XLWorkbook workbook, KeyValuePair<SheetInfo, IEnumerable<TItem>> sheetItem, int sheetIndex)
        {

            IXLWorksheet worksheet = null;
            if (workbook.Worksheets.Count <= sheetIndex)
            {
                worksheet = workbook.Worksheets.Add(sheetItem.Key.Name);
            }
            else
            {
                worksheet = workbook.Worksheets.ElementAt(sheetIndex);
            }

            var initRow = 1;


            // Formato de la cabecera (si es necesario)
            if (!exporter._hideHeaders)
            {

                for (var column = 1; column <= _classPropInfo.Count; column++)
                {
                    var columnToDisplay = _classPropInfo[column - 1];
                    var cell = worksheet.Cell(initRow, column);
                    cell.Value = columnToDisplay.DisplayName;
                }

                initRow++;
            }
            else if (!exporter._jumpHeaders)
            {
                initRow++;
            }


            var finalRow = initRow + sheetItem.Value.Count();

            for (var row = initRow; row < finalRow; row++)
            {

                var rowDataItem = sheetItem.Value.ElementAt(row - initRow);

                for (var column = 1; column <= _classPropInfo.Count; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    var propToDisplay = _classPropInfo[column - 1].PropertyInfo;

                    cell.Value = propToDisplay.GetValue(rowDataItem);
                }

                FormatRow(worksheet.Row(row), rowDataItem, exporter);

            }
        }

        private void FormatRow(IXLRow excelRow, TItem data, Exporter<TItem> exporter)
        {
            FormatFontStyle(excelRow, data, exporter);
            FormatFontSize(excelRow, data, exporter);
        }

        private void FormatFontSize(IXLRow excelRow, TItem data, Exporter<TItem> exporter)
        {
            foreach (var filter in exporter._fontSizeFormatters)
            {
                if (filter.Item1(data))
                {
                    excelRow.Style.Font.FontSize = filter.Item2;
                }
            }
        }

        private void FormatFontStyle(IXLRow excelRow, TItem data, Exporter<TItem> exporter)
        {
            foreach (var filter in exporter._fontFormatters)
            {
                if (filter.Item1(data))
                {
                    switch (filter.Item2)
                    {
                        case FontFormat.Bold:
                            {
                                excelRow.Style.Font.Bold = true;
                                break;
                            }
                        case FontFormat.Italic:
                            {
                                excelRow.Style.Font.Italic = true;
                                break;
                            }
                        case FontFormat.Underline:
                            {
                                excelRow.Style.Font.Underline = XLFontUnderlineValues.Single;
                                break;
                            }
                    }
                }
            }
        }



    }
}
