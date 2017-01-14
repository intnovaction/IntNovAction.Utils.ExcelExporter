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
        public static void CheckFormat(IXLCell appliedFormat, FormatConfigurator theoricFormat)
        {
            appliedFormat.Style.Font.Bold.Should().Be(theoricFormat._bold);
            appliedFormat.Style.Font.Italic.Should().Be(theoricFormat._italic);

            var underline = theoricFormat._underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
            appliedFormat.Style.Font.Underline.Should().Be(underline);

            if (theoricFormat._fontSize.HasValue)
            {
                appliedFormat.Style.Font.FontSize.Should().Be(theoricFormat._fontSize);
            }

            if (theoricFormat._color != null)
            {
                appliedFormat.Style.Font.FontColor.Color.R.Should().Be((byte)theoricFormat._color.Red);
                appliedFormat.Style.Font.FontColor.Color.G.Should().Be((byte)theoricFormat._color.Green);
                appliedFormat.Style.Font.FontColor.Color.B.Should().Be((byte)theoricFormat._color.Blue);
            }

            var border = theoricFormat._bottomBorder ? XLBorderStyleValues.Medium : XLBorderStyleValues.None;
            appliedFormat.Style.Border.BottomBorder.Should().Be(border);

        }


        public static void CheckFormat(IXLCell appliedFormat, Action<FormatConfigurator> formatAction)
        {
            var formatConfig = new FormatConfigurator();
            formatAction.Invoke(formatConfig);

            CheckFormat(appliedFormat, formatConfig);
        }
    }
}
