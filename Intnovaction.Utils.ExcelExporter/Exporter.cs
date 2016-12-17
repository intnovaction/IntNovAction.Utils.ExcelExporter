using IntNovAction.Utils.ExcelExporter.ExcelWriters;
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

        /// <summary>
        /// Constructor para la creación de una nueva hoja excel
        /// </summary>
        public Exporter()
        {
            _sheets = new SheetCollection();
        }

        /// <summary>
        /// Constructor para trabajar sobre un excel existente
        /// </summary>
        /// <param name="existingFileStream">Stream con el contenido del excel sobre el que se va a trabajar</param>
        public Exporter(Stream existingFileStream) : this()
        {
            _existingFileStream = existingFileStream;
        }

        /// <summary>
        /// Añade una hoja al excel
        /// </summary>
        /// <typeparam name="TDataItem">El tipo de datos que se va a poner en la hoja</typeparam>
        /// <param name="config">Expresioón para trabajar con el configurador de la hoja</param>
        /// <returns>Exportador</returns>
        public Exporter AddSheet<TDataItem>(Action<SheetConfigurator<TDataItem>> config)
            where TDataItem : new()
        {
            var configurator = new SheetConfigurator<TDataItem>();
            config.Invoke(configurator);

            configurator._order = this._sheets.Count;
            this._sheets.Add(configurator);


            return this;
        }

        /// <summary>
        /// Devuelve el array de bytes con el contenido del excel configurado
        /// </summary>
        /// <returns></returns>
        public byte[] Export()
        {
            var excelExporter = GetFormatter();

            var elems = excelExporter.Export(this);

            return elems;
        }

        private ExcelFileWriter GetFormatter()
        {
            return new ExcelFileWriter();
        }
    }
}