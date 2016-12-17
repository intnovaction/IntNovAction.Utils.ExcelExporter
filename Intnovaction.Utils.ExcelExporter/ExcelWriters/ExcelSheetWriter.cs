using ClosedXML.Excel;
using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.ExcelWriters
{
    internal class ExcelSheetGenerator<TDataItem>
         where TDataItem : new()
    {
        public ExcelSheetGenerator() : base()
        {
        }

        public void WriteSheet(XLWorkbook workbook, SheetConfigurator<TDataItem> sheetConfig)

        {
            var _classPropInfo = sheetConfig._columnsConfig;

            var columns = sheetConfig._columnsConfig.GetColumns();

            IXLWorksheet worksheet = null;
            if (workbook.Worksheets.Count < sheetConfig._order + 1)
            {
                worksheet = workbook.Worksheets.Add(sheetConfig._name);
            }
            else
            {
                worksheet = workbook.Worksheets.ElementAt(sheetConfig._order);
            }

            var initRow = sheetConfig._initialRow;

            if (sheetConfig._title != null)
            {
                var cell = worksheet.Cell(initRow, 1);
                cell.SetValue(sheetConfig._title._TitleText);

                if (sheetConfig._title._Format != null)
                {
                    ApplyFormat(cell.Style, sheetConfig._title._Format);
                }

                initRow++;
            }

            // Formato de la cabecera (si es necesario)
            if (!sheetConfig._hideColumnHeaders)
            {
                for (var column = initRow; column <= columns.Count; column++)
                {
                    var columnToDisplay = columns[column - 1];
                    var cell = worksheet.Cell(initRow, column);
                    cell.Value = columnToDisplay._title;
                }

                initRow++;
            }


            var finalRow = initRow + sheetConfig._data.Count();

            for (var row = initRow; row < finalRow; row++)
            {
                var rowDataItem = sheetConfig._data.ElementAt(row - initRow);

                for (var column = 1; column <= columns.Count; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    
                    if (columns[column - 1].Expression == null)
                    {
                        var propToDisplay = columns[column - 1].PropertyInfo;
                        cell.Value = propToDisplay.GetValue(rowDataItem);
                    }
                    else
                    {
                        cell.Value = columns[column - 1].Expression.Invoke(rowDataItem);
                    }
                }

                FormatRow(worksheet.Row(row), rowDataItem, sheetConfig);
            }
        }



        private static void ApplyFormat(IXLStyle style, FormatConfigurator configurator)
        {
            if (configurator._bold)
            {
                style.Font.Bold = true;
            }
            if (configurator._underline)
            {
                style.Font.Underline = XLFontUnderlineValues.Single;
            }
            if (configurator._italic)
            {
                style.Font.Italic = true;
            }

            if (configurator._color != null)
            {
                style.Font.FontColor = XLColor.FromArgb(configurator._color.Red, configurator._color.Green, configurator._color.Blue);
            }


            if (configurator._fontSize.HasValue)
            {
                style.Font.FontSize = configurator._fontSize.Value;
            }
        }

        private void FormatRow(IXLRow excelRow, TDataItem data, SheetConfigurator<TDataItem> configurator)
        {
            foreach (var filter in configurator._fontFormatters)
            {
                if (filter.Item1(data))
                {
                    ApplyFormat(excelRow.Style, filter.Item2);
                }
            }
        }

    }
}