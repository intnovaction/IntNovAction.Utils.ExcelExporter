using IntNovAction.Utils.ExcelExporter.Configurators;
using System;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

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
                    _orderFromMetadata = attr.GetOrder() ?? int.MaxValue,
                    PropertyInfo = prop,
                };
                dataItem._columnHeaderFormat.Text(attr.GetName() ?? prop.Name);
            }
            else
            {
                dataItem = new ColumnConfigurator<TDataItem>()
                {
                    _orderFromMetadata = Int16.MaxValue,
                    PropertyInfo = prop,
                };
                dataItem._columnHeaderFormat.Text(prop.Name);
            }

            return dataItem;
        }
    }
}