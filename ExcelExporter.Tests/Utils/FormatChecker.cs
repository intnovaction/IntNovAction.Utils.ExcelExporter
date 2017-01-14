using ClosedXML.Excel;
using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;

namespace IntNovAction.Utils.ExcelExporter.Tests.Utils
{
    class FormatChecker
    {
        public static void CheckFormat(IXLCell cellToCheck, FormatConfigurator theoricFormat)
        {
            cellToCheck.Style.Font.Bold.Should().Be(theoricFormat._bold, "Incorrect bold");
            cellToCheck.Style.Font.Italic.Should().Be(theoricFormat._italic, "Incorrect Italic");

            var underline = theoricFormat._underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
            cellToCheck.Style.Font.Underline.Should().Be(underline, "Incorrect underline");

            if (theoricFormat._fontSize.HasValue)
            {
                cellToCheck.Style.Font.FontSize.Should().Be(theoricFormat._fontSize);
            }

            if (theoricFormat._color != null)
            {
                cellToCheck.Style.Font.FontColor.Color.R.Should().Be((byte)theoricFormat._color.Red);
                cellToCheck.Style.Font.FontColor.Color.G.Should().Be((byte)theoricFormat._color.Green);
                cellToCheck.Style.Font.FontColor.Color.B.Should().Be((byte)theoricFormat._color.Blue);
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
