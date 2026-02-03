using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    /// <summary>
    /// Configurador de una hoja
    /// </summary>
    public class SheetConfigurator<TDataItem> : SheetConfiguratorBase
        where TDataItem : new()
    {
        /// <summary>
        /// La información de las columnas
        /// </summary>
        internal ColumnCollection<TDataItem> _columnsConfig;

        /// <summary>
        /// Los datos a pintar
        /// </summary>
        internal IEnumerable<TDataItem> _data;

        /// <summary>
        /// Los formateadores de las filas
        /// </summary>
        internal List<Tuple<Func<TDataItem, bool>, FormatConfigurator>> _rowFormatRules;

        public SheetConfigurator()
        {
            _rowFormatRules = new List<Tuple<Func<TDataItem, bool>, FormatConfigurator>>();
            _columnsConfig = new ColumnCollection<TDataItem>();
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

            _rowFormatRules.Add(new Tuple<Func<TDataItem, bool>, FormatConfigurator>(condition, configurator));

            return this;
        }

        /// <summary>
        /// Indica que se apliquen los estilos por defecto
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> ApplyDefaultStyles()
        {
            return ApplyDefaultStyles(true);
        }

        /// <summary>
        /// Indica que se apliquen o no los estilos por defecto
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> ApplyDefaultStyles(bool value)
        {
            _applyDefaultStyle = value;
            return this;
        }

        /// <summary>
        /// Configura las columnas de la hoja
        /// </summary>
        /// <typeparam name="TDataItem">El tipo de datos que se va a poner en la hoja</typeparam>
        /// <param name="config">Expresión para trabajar la coleccion de columnas</param>
        /// <returns>Exportador</returns>
        public SheetConfigurator<TDataItem> Columns(Action<ColumnCollection<TDataItem>> config)
        {
            config.Invoke(_columnsConfig);
            return this;
        }

        /// <summary>
        /// Indica que no se muestren los headers de las columnas
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> HideColumnHeaders()
        {
            return HideColumnHeaders(true);
        }

        /// <summary>
        /// Indica que se oculten o no los headers de las columnas
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> HideColumnHeaders(bool value)
        {
            _hideColumnHeaders = value;
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
        /// Imprime la hoja en una sola página
        /// </summary>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> PrintInOnePage()
        {
            _fitInOnePage = true;
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
        /// <param name="config">Expresión para trabajar con el configurador de la cabecera</param>
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
        /// Establece el contenido de una celda personalizada
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetCustomContent(int row, int column, string value)
        {
            return SetCustomContent(row, column, value, null);
        }

        /// <summary>
        /// Establece el contenido de una celda personalizada con formato
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <param name="formatConfigurator">Expresión para especificar el formato de la celda</param>
        /// <returns></returns>
        public SheetConfigurator<TDataItem> SetCustomContent(int row, int column, string value, Action<FormatConfigurator> formatConfigurator)
        {
            if (row < 1 || column < 1)
                throw new ArgumentOutOfRangeException("The minimum coordinates are 1, 1");

            FormatConfigurator format = null;
            if (formatConfigurator != null)
            {
                format = new FormatConfigurator();
                formatConfigurator.Invoke(format);
            }

            _customCells.Add(new CustomCell
            {
                Row = row,
                Column = column,
                Value = value,
                Format = format
            });

            return this;
        }
    }
}