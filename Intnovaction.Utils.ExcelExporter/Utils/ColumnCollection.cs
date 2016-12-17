using IntNovAction.Utils.ExcelExporter.Exceptions;
using System.Collections.Generic;
using System.Linq;
using System;

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
        }

        /// <summary>
        /// Añade una column
        /// </summary>
        /// <param name="sheetConfigurator">El configurador de la hoja</param>
        public ColumnConfigurator<TDataItem> AddColumnExpr(Func<TDataItem, object> expr, string title)
        {
            var columnConfigurator = new ColumnConfigurator<TDataItem>();

            columnConfigurator.Expression = expr;
            columnConfigurator._title = title;
            columnConfigurator.Order = int.MaxValue;
            _columnCol.Add(columnConfigurator);

            return columnConfigurator;

        }

        public ColumnConfigurator<TDataItem> AddColumn<TProp>(System.Linq.Expressions.Expression<Func<TDataItem, TProp>> expression)
        {
            var body = expression.Body as System.Linq.Expressions.MemberExpression;
            if (body == null)
            {
                throw new ArgumentException("'expression' should be a member expression");
            }

            var propertyInfo = (System.Reflection.PropertyInfo)body.Member;

            var columnConfigurator = ReflectionHelper<TDataItem>.GetColumnFromPropertyInfo(propertyInfo);

            _columnCol.Add(columnConfigurator);

            return columnConfigurator;
        }

        internal ColumnConfigurator<TDataItem> AddColumn(ColumnConfigurator<TDataItem> column)
        {
            _columnCol.Add(column);
            return column;
        }

        /// <summary>
        /// Borra todas las columnas
        /// </summary>
        public void Clear()
        {
            _columnCol.Clear();
        }

        public IList<ColumnConfigurator<TDataItem>> GetColumns()
        {
            // Ordenamos
            if (_columnCol.Select(p => p.Order).Distinct().Count() != 1)
            {
                return _columnCol
                    .OrderBy(p => p.Order)
                    .ThenBy(p => p._title)
                    .ToList();
            }
            else
            {
                return _columnCol;
            }
        }

    }
}