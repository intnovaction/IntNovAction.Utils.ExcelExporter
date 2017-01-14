using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// La colección de columnas dentro de una hoja
    /// </summary>
    public class ColumnCollection<TDataItem>
    {
        private List<ColumnConfigurator<TDataItem>> _columnCol;

        public ColumnCollection()
        {
            _columnCol = new List<Utils.ColumnConfigurator<TDataItem>>();

            var type = typeof(TDataItem);
            var allProps = type.GetProperties();
            foreach (var prop in allProps)
            {
                var dataItem = ReflectionHelper<TDataItem>.GetColumnFromPropertyInfo(prop);
                AddColumn(dataItem);
            }
            if (_columnCol.Select(p => p._orderFromMetadata).Distinct().Count() > 1)
            {
                _columnCol = _columnCol.OrderBy(p => p._orderFromMetadata).ToList();
            }
        }

        /// <summary>
        /// Añade una columna
        /// </summary>
        /// <typeparam name="TProp"></typeparam>
        /// <param name="expression">Expresión lambda con la propiedad que se pintará en la columna</param>
        /// <returns></returns>
        public ColumnConfigurator<TDataItem> AddColumn<TProp>(Expression<Func<TDataItem, TProp>> expression)
        {
            PropertyInfo propertyInfo = ExtractPropertyInfoFromExpression(expression);

            var columnConfigurator = ReflectionHelper<TDataItem>.GetColumnFromPropertyInfo(propertyInfo);

            _columnCol.Add(columnConfigurator);

            return columnConfigurator;
        }

        /// <summary>
        /// Añade una column
        /// </summary>
        /// <param name="sheetConfigurator">El configurador de la hoja</param>
        public ColumnConfigurator<TDataItem> AddColumnExpr(Func<TDataItem, object> expr, string title)
        {
            var columnConfigurator = new ColumnConfigurator<TDataItem>();

            columnConfigurator.Expression = expr;
            columnConfigurator._columnTitle = title;
            columnConfigurator._orderFromMetadata = int.MaxValue;
            _columnCol.Add(columnConfigurator);

            return columnConfigurator;
        }

        /// <summary>
        /// Borra todas las columnas
        /// </summary>
        public void Clear()
        {
            _columnCol.Clear();
        }

        /// <summary>
        /// Oculta una columna
        /// </summary>
        /// <typeparam name="TProp"></typeparam>
        /// <param name="expression">Expresión lambda con la propiedad correspondiente a la columna que no se pintará</param>
        /// <returns></returns>
        public void HideColumn<TProp>(Expression<Func<TDataItem, TProp>> expression)
        {
            PropertyInfo propertyInfo = ExtractPropertyInfoFromExpression(expression);

            _columnCol.RemoveAll(p => p.PropertyInfo != null && p.PropertyInfo.Name == propertyInfo.Name);

            return;
        }

        internal ColumnConfigurator<TDataItem> AddColumn(ColumnConfigurator<TDataItem> column)
        {
            _columnCol.Add(column);
            return column;
        }

        internal IList<ColumnConfigurator<TDataItem>> GetColumns()
        {
            return _columnCol;
        }

        private PropertyInfo ExtractPropertyInfoFromExpression<TProp>(Expression<Func<TDataItem, TProp>> expression)
        {
            var body = expression.Body as MemberExpression;
            if (body == null)
            {
                throw new ArgumentException("'expression' should be a member expression");
            }
            var propertyInfo = (PropertyInfo)body.Member;
            return propertyInfo;
        }
    }
}