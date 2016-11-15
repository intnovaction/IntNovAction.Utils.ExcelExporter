using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IntNovAction.Utils.ExcelExporter;
using System.Reflection;
using System.ComponentModel.DataAnnotations;

namespace IntNovAction.Utils.FormatExporters
{
    internal abstract class BaseExporter<TItem>
        where TItem : new()
    {

        public BaseExporter()
        {
            _classPropInfo = ReadClassInfo();
        }

        public abstract byte[] Export(Exporter<TItem> exporter);


        protected List<PropInfo> _classPropInfo;

        List<PropInfo> ReadClassInfo()
        {
            var type = typeof(TItem);


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

        protected class PropInfo
        {
            public int Order { get; set; }

            public string DisplayName { get; set; }

            public PropertyInfo PropertyInfo { get; set; }
        }
    }
}
