using System;

namespace IntNovAction.Utils.ExcelExporter.Exceptions
{
    public class DuplicatedColumnTitleException : ApplicationException
    {
        public DuplicatedColumnTitleException() : base()
        {
        }

        public DuplicatedColumnTitleException(string columnName) : base()
        {
            DuplicatedTitle = columnName;
        }

        public String DuplicatedTitle { get; set; }
    }
}