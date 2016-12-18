namespace IntNovAction.Utils.ExcelExporter.Utils
{

    /// <summary>
    /// Tiene que haber una clase base para poder manejar la colección
    /// </summary>
    public abstract class SheetConfiguratorBase
    {

        internal bool _hideColumnHeaders = false;
        
        internal string _name { get; set; }
        
        internal int _order { get; set; }

        internal int _initialRow = 1;

        internal int _initialColumn = 1;


        internal TitleConfigurator _title = null;


    }
}