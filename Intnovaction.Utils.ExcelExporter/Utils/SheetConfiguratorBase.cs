namespace IntNovAction.Utils.ExcelExporter.Utils
{
    public abstract class SheetConfiguratorBase
    {


        internal string _name { get; set; }


        internal int _order { get; set; }

        internal int _initialRow = 1;

        internal int _initialColumn = 1;


        internal TitleConfigurator _title = null;


    }
}