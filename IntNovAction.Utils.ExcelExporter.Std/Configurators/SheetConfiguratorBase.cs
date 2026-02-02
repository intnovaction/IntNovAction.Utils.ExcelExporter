using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelExporter.Configurators
{
    /// <summary>
    /// Tiene que haber una clase base para poder manejar la colección
    /// </summary>
    public abstract class SheetConfiguratorBase
    {
        internal bool _applyDefaultStyle = false;
        internal bool _fitInOnePage = false;
        internal bool _hideColumnHeaders = false;

        internal int _initialColumn = 1;
        internal int _initialRow = 1;
        internal TitleConfigurator _title = null;
        internal string _name { get; set; }

        internal int _order { get; set; }

        internal List<CustomCell> _customCells = new List<CustomCell>();
    }
}