using ClosedXML.Excel;
using IntNovAction.Utils.ExcelExporter.Configurators;
using IntNovAction.Utils.ExcelExporter.Utils;
using System.Linq;

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
            var initColumn = sheetConfig._initialColumn;

            if (sheetConfig._title != null)
            {
                var cell = worksheet.Cell(initRow, initColumn);
                cell.SetValue(sheetConfig._title._TitleText);

                if (sheetConfig._applyDefaultStyle)
                {
                    ApplyTitleDefaultStyle(cell);
                }

                if (sheetConfig._title._Format != null)
                {
                    ApplyFormat(cell.Style, sheetConfig._title._Format);
                }

                initRow++;
            }

            // Textos de la cabecera (si es necesario)
            if (!sheetConfig._hideColumnHeaders)
            {
                for (var column = initColumn; column < initColumn + columns.Count; column++)
                {
                    // La columna
                    var columnToDisplay = columns[column - initColumn];

                    // Ponemos el titulo
                    var cell = worksheet.Cell(initRow, column);
                    cell.Value = columnToDisplay._columnHeaderFormat._text;

                    if (sheetConfig._applyDefaultStyle)
                    {
                        ApplyHeaderCellDefaultStyle(cell);
                    }
                    if (columnToDisplay._columnHeaderFormat._Format != null)
                    {
                        ApplyFormat(cell.Style, columnToDisplay._columnHeaderFormat._Format);
                    }
                }

                initRow++;
            }

            var firstRowWithData = initRow;
            var lastRowWithData = initRow + sheetConfig._data.Count();

            // Formato de las columnas
            for (var column = initColumn; column < initColumn + columns.Count; column++)
            {
                var columnToDisplay = columns[column - initColumn];
                // Establecemos el formato
                if (columnToDisplay._columnFormat != null)
                {
                    var columnToFormat = worksheet.Column(column);
                    ApplyFormat(columnToFormat, firstRowWithData, lastRowWithData, columnToDisplay._columnFormat);
                }
            }

            for (var row = initRow; row < lastRowWithData; row++)
            {
                var rowDataItem = sheetConfig._data.ElementAt(row - initRow);

                for (var column = initColumn; column < initColumn + columns.Count; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    var excelColumn = columns[column - initColumn];

                    if (excelColumn.Expression == null)
                    {
                        var propToDisplay = excelColumn.PropertyInfo;
                        cell.Value = XLCellValue.FromObject(propToDisplay.GetValue(rowDataItem));
                    }
                    else
                    {
                        cell.Value = XLCellValue.FromObject(excelColumn.Expression.Invoke(rowDataItem));
                    }
                }
                FormatRow(worksheet.Row(row), rowDataItem, sheetConfig);
            }

            foreach (var customCell in sheetConfig._customCells)
            {
                worksheet.Cell(customCell.Row, customCell.Column).Value = XLCellValue.FromObject(customCell.Value);
            }

            if (sheetConfig._fitInOnePage)
            {
                worksheet.PageSetup.PrintAreas.Add(sheetConfig._initialRow, sheetConfig._initialColumn, lastRowWithData, initColumn + columns.Count);
            }
        }

        /// <summary>
        /// Aplica un formato a una columan de datos
        /// </summary>
        /// <param name="column">La columna (para poner el ancho)</param>
        /// <param name="firstRow">Primera fila con datos</param>
        /// <param name="lastRow">Ultima fila con datos</param>
        /// <param name="configurator">El formato</param>
        private void ApplyFormat(IXLColumn column, int firstRow, int lastRow, ColumnDataFormatConfigurator configurator)
        {
            if (configurator._width.HasValue)
            {
                column.Width = configurator._width.Value;
            }

            var range = column.Cells(firstRow, lastRow);

            ApplyFormat(range.Style, (FormatConfigurator)configurator);
        }

        private void ApplyFormat(IXLStyle style, FormatConfigurator configurator)
        {
            if (configurator._bold)
            {
                style.Font.Bold = true;
                style.Font.SetBold();
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

            if (configurator._backcolor != null)
            {
                style.Fill.BackgroundColor = XLColor.FromArgb(configurator._backcolor.Red, configurator._backcolor.Green, configurator._backcolor.Blue);
            }

                if (configurator._fontSize.HasValue)
            {
                style.Font.FontSize = configurator._fontSize.Value;
            }

            if (configurator._bottomBorder)
            {
                style.Border.BottomBorder = XLBorderStyleValues.Medium;
            }
        }

        private void FormatRow(IXLRow excelRow, TDataItem data, SheetConfigurator<TDataItem> configurator)
        {
            foreach (var rule in configurator._rowFormatRules)
            {
                if (rule.Item1(data))
                {
                    ApplyFormat(excelRow.Style, rule.Item2);
                }
            }
        }

        #region Default Style Formatters

        private void ApplyHeaderCellDefaultStyle(IXLCell cell)
        {
            ApplyFormat(cell.Style, Utils.DefaultStyles.GetHeadersDefaultStyle());
        }

        private void ApplyTitleDefaultStyle(IXLCell cell)
        {
            ApplyFormat(cell.Style, Utils.DefaultStyles.GetTitleDefaultStyle());
        }

        #endregion Default Style Formatters
    }
}