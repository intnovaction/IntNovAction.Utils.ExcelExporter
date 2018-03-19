using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelExporter.Configurators;
using IntNovAction.Utils.ExcelExporter.Utils;
using System;

namespace IntNovAction.Utils.ExcelExporter.Tests.Utils
{
    internal class FormatChecker
    {
        public static void CheckFormat(IXLCell cellToCheck, FormatConfigurator theoricFormat)
        {
            var fontStyle = cellToCheck.Style.Font;
            fontStyle.Bold.Should().Be(theoricFormat._bold, "Incorrect bold");
            fontStyle.Italic.Should().Be(theoricFormat._italic, "Incorrect Italic");

            var underline = theoricFormat._underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
            fontStyle.Underline.Should().Be(underline, "Incorrect underline");

            if (theoricFormat._fontSize.HasValue)
            {
                fontStyle.FontSize.Should().Be(theoricFormat._fontSize);
            }

            if (theoricFormat._color != null)
            {
                fontStyle.FontColor.Color.R.Should().Be((byte)theoricFormat._color.Red);
                fontStyle.FontColor.Color.G.Should().Be((byte)theoricFormat._color.Green);
                fontStyle.FontColor.Color.B.Should().Be((byte)theoricFormat._color.Blue);
            }

            if (theoricFormat._backcolor != null)
            {
                cellToCheck.Style.Fill.BackgroundColor.Color.B.Should().Be((byte)theoricFormat._backcolor.Blue);
                cellToCheck.Style.Fill.BackgroundColor.Color.G.Should().Be((byte)theoricFormat._backcolor.Green);
                cellToCheck.Style.Fill.BackgroundColor.Color.R.Should().Be((byte)theoricFormat._backcolor.Red);
            }


            var border = theoricFormat._bottomBorder ? XLBorderStyleValues.Medium : XLBorderStyleValues.None;
            cellToCheck.Style.Border.BottomBorder.Should().Be(border);
        }

        public static void CheckFormat(IXLCell cellToCheck, Action<FormatConfigurator> formatAction)
        {
            var formatConfig = new FormatConfigurator();
            formatAction.Invoke(formatConfig);

            CheckFormat(cellToCheck, formatConfig);
        }
    }
}