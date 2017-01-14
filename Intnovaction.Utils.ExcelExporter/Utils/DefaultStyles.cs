using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    internal class DefaultStyles
    {
        internal static FormatConfigurator GetTitleDefaultStyle()
        {
            var formatConfig = new FormatConfigurator()
                .Bold()
                .Italic()
                .FontSize(18);

            return formatConfig;
        }

        internal static FormatConfigurator GetHeadersDefaultStyle()
        {
            var formatConfig = new FormatConfigurator()
                .Bold()
                .BottomBorder()
                .FontSize(12);

            return formatConfig;
        }
    }
}
