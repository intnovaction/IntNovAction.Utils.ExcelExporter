using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    internal static class ReflectionHelper<TDataItem>
    {

        internal static ColumnConfigurator<TDataItem> GetColumnFromPropertyInfo(PropertyInfo prop)
        {
            ColumnConfigurator<TDataItem> dataItem = null;
            var attr = prop.GetCustomAttribute<DisplayAttribute>();
            if (attr != null)
            {
                dataItem = new ColumnConfigurator<TDataItem>()
                {
                    _columnTitle = attr.GetName() ?? prop.Name,
                    _orderFromMetadata = attr.GetOrder() ?? int.MaxValue,
                    PropertyInfo = prop,
                };
            }
            else
            {
                dataItem = new ColumnConfigurator<TDataItem>()
                {
                    _columnTitle = prop.Name,
                    _orderFromMetadata = Int16.MaxValue,
                    PropertyInfo = prop,
                };
            }

            return dataItem;
        }
    }
}
