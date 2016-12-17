using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelExporter.Utils
{
    /// <summary>
    /// Configurador de una hoja
    /// </summary>
    public class SheetConfigurator<TDataItem> : SheetConfiguratorBase
        where TDataItem : new()
    {
        internal List<ColumnConfigurator> _columnsConfig;
        internal IEnumerable<TDataItem> _data;

        internal List<Tuple<Func<TDataItem, bool>, FormatConfigurator>> _fontFormatters;

        internal bool _hideColumnHeaders = false;
        public SheetConfigurator()
        {
            _fontFormatters = new List<Tuple<Func<TDataItem, bool>, FormatConfigurator>>();

            _columnsConfig = ReadClassColumns();
        }

        /// <summary>
        /// Añade un formato condicional de fila a una hoja
        /// </summary>
        /// <param name="condition">Condición para aplicar o no el formato</param>
        /// <param name="formatConfigurator">Expresión para especificar el formato de la fila</param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> AddFormatRule(Func<TDataItem, bool> condition, Action<FormatConfigurator> formatConfigurator)
        {
            var configurator = new FormatConfigurator();
            formatConfigurator.Invoke(configurator);

            _fontFormatters.Add(new Tuple<Func<TDataItem, bool>, FormatConfigurator>(condition, configurator));

            return this;
        }

        /// <summary>
        /// Indica que no se muestren los headers de las columnas
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> HideColumnHeaders()
        {
            _hideColumnHeaders = true;
            return this;
        }

        /// <summary>
        /// Establece el nombre de la hoja
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> Name(string name)
        {
            _name = name;
            return this;
        }

        /// <summary>
        /// Establece el orden en el que la hoja se añadirá al Excel
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> Order(int order)
        {
            _order = order;
            return this;
        }

        /// <summary>
        /// Establece las coordenadas de donde empieza a pintar los datos
        /// </summary>
        /// <param name="initialRow"></param>
        /// <param name="initialColumn"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetCoordinates(int initialRow, int initialColumn)
        {
            if (initialRow < 1 || initialColumn < 1)
            {
                throw new ArgumentOutOfRangeException("The minimum coordinates are 1, 1");
            }

            _initialRow = initialRow;
            _initialColumn = initialColumn;
            return this;
        }
        /// <summary>
        /// Establece los datos a mostrar en la hoja
        /// </summary>
        /// <param name="data">Datos a pintar en la hoja</param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetData(IEnumerable<TDataItem> data)
        {
            _data = data;
            return this;
        }

        


        /// <summary>
        /// Establece la cabecera de la hoja
        /// </summary>
        /// <typeparam name="TDataItem">El tipo de datos que se va a poner en la hoja</typeparam>
        /// <param name="config">Expresioón para trabajar con el configurador de la cabecera</param>
        /// <returns>Exportador</returns>
        public SheetConfigurator<TDataItem> Title(Action<TitleConfigurator> config)            
        {
            var configurator = new TitleConfigurator();
            configurator.Text(_name);

            config.Invoke(configurator);

            this._title = configurator;
            

            return this;
        }

        /// <summary>
        /// Añade la cabecera por defecto a la hoja
        /// </summary>
        /// <typeparam name="TDataItem">El tipo de datos que se va a poner en la hoja</typeparam>
        /// <returns>Exportador</returns>
        public SheetConfigurator<TDataItem> Title()
        {
            var configurator = new TitleConfigurator();
            
            this._title = configurator;
            _title.Text(_name);

            return this;
        }


        /// <summary>
        /// Lee las propiedades de la clase de la que vamos a pintar el excel y rellena una lista
        /// </summary>
        /// <returns></returns>
        private List<ColumnConfigurator> ReadClassColumns()
        {
            var type = typeof(TDataItem);

            var result = new List<ColumnConfigurator>();

            var allProps = type.GetProperties();
            foreach (var prop in allProps)
            {
                var attr = prop.GetCustomAttribute<DisplayAttribute>();

                if (attr != null)
                {
                    result.Add(new ColumnConfigurator()
                    {
                        DisplayName = attr.GetName() ?? prop.Name,
                        Order = attr.GetOrder() ?? int.MaxValue,
                        PropertyInfo = prop,
                    });
                }
                else
                {
                    result.Add(new ColumnConfigurator()
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