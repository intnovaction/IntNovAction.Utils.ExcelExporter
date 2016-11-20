namespace IntNovAction.Utils.ExcelExporter.Utils
{
    public abstract class SheetConfiguratorBase
    {
        internal bool _jumpHeaders;
        internal string _title;
        internal string _name { get; set; }

        internal int _order { get; set; }
    }
}