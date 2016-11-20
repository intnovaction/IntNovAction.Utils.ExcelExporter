using IntNovAction.Utils.ExcelExporter.FormatExporters;
using IntNovAction.Utils.ExcelExporter.Utils;
using System;
using System.IO;

namespace IntNovAction.Utils.ExcelExporter
{
    /// <summary>
    /// Genera un excel con los elementos de tipo ListItem como filas
    /// </summary>
    public class Exporter
    {
        internal Stream _existingFileStream;
        internal SheetCollection _sheets;

        public Exporter()
        {
            _sheets = new SheetCollection();
        }

        public Exporter(Stream existingFileStream) : this()
        {
            _existingFileStream = existingFileStream;
        }

        /// <summary>
        /// Añade una hoja al excel
        /// </summary>
        /// <typeparam name="TDataItem">El tipo de datos que se va a poner en la hoja</typeparam>
        /// <param name="config">Configurador de la hoja</param>
        /// <returns></returns>
        public Exporter AddSheet<TDataItem>(Action<SheetConfigurator<TDataItem>> config)
            where TDataItem : new()
        {
            var configurator = new SheetConfigurator<TDataItem>();
            config.Invoke(configurator);

            this._sheets.Add(configurator);

            return this;
        }

        public byte[] Export()
        {
            var excelExporter = GetFormatter();

            var elems = excelExporter.Export(this);

            return elems;
        }

        private ExcelGenerator GetFormatter()
        {
            return new ExcelGenerator();
        }
    }
}