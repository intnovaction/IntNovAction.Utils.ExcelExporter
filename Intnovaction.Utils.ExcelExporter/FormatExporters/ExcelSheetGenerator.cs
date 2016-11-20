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

namespace IntNovAction.Utils.ExcelExporter.FormatExporters
{
    internal class ExcelSheetGenerator<TDataItem>
         where TDataItem : new()
    {
        public ExcelSheetGenerator() : base()
        {

        }


        public void WriteSheet(XLWorkbook workbook, SheetConfigurator<TDataItem> sheetConfig)
           
        {
            var _classPropInfo = ReadClassInfo();



            IXLWorksheet worksheet = null;
            if (workbook.Worksheets.Count <= sheetConfig._order)
            {
                worksheet = workbook.Worksheets.Add(sheetConfig._name);
            }
            else
            {
                worksheet = workbook.Worksheets.ElementAt(sheetConfig._order);
            }

            var initRow = 1;


            // Formato de la cabecera (si es necesario)
            if (!sheetConfig._hideHeaders)
            {

                for (var column = 1; column <= _classPropInfo.Count; column++)
                {
                    var columnToDisplay = _classPropInfo[column - 1];
                    var cell = worksheet.Cell(initRow, column);
                    cell.Value = columnToDisplay.DisplayName;
                }

                initRow++;
            }
            else if (!sheetConfig._jumpHeaders)
            {
                initRow++;
            }


            var finalRow = initRow + sheetConfig._data.Count();

            for (var row = initRow; row < finalRow; row++)
            {

                var rowDataItem = sheetConfig._data.ElementAt(row - initRow);

                for (var column = 1; column <= _classPropInfo.Count; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    var propToDisplay = _classPropInfo[column - 1].PropertyInfo;

                    cell.Value = propToDisplay.GetValue(rowDataItem);
                }

                FormatRow(worksheet.Row(row), rowDataItem, sheetConfig);

            }
        }

        private void FormatRow(IXLRow excelRow, TDataItem data, SheetConfigurator<TDataItem> configurator)
        {
            FormatFontStyle(excelRow, data, configurator);
            FormatFontSize(excelRow, data, configurator);
        }

        private void FormatFontSize(IXLRow excelRow, TDataItem data, SheetConfigurator<TDataItem> configurator)
        {
            foreach (var filter in configurator._fontSizeFormatters)
            {
                if (filter.Item1(data))
                {
                    excelRow.Style.Font.FontSize = filter.Item2;
                }
            }
        }

        private void FormatFontStyle(IXLRow excelRow, TDataItem data, SheetConfigurator<TDataItem> configurator)
        {
            foreach (var filter in configurator._fontFormatters)
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

        
        List<PropInfo> ReadClassInfo()
        {
            var type = typeof(TDataItem);


            var result = new List<PropInfo>();

            var allProps = type.GetProperties();
            foreach (var prop in allProps)
            {
                var attr = prop.GetCustomAttribute<DisplayAttribute>();

                if (attr != null)
                {
                    result.Add(new PropInfo()
                    {
                        DisplayName = attr.GetName() ?? prop.Name,
                        Order = attr.GetOrder() ?? int.MaxValue,
                        PropertyInfo = prop,
                    });
                }
                else
                {
                    result.Add(new PropInfo()
                    {
                        DisplayName = prop.Name,
                        Order = Int16.MaxValue,
                        PropertyInfo = prop,
                    });
                }
            }

            // Ordenamos
            result = result.OrderBy(p => p.Order).ThenBy(p => p.DisplayName).ToList();


            return result;
        }



       
    }
}
