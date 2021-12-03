using System;
using System.Collections.Generic;
using System.Text;

namespace ExportToExcelConsoleApp
{
    public class DataSetJson
    {
        public string Name { get; set; }
        public List<DataTableNew> DataTables { get; set; }
    }

    public class DataTableNew
    {
        public string Name { get; set; }
        public DataTableMetaData MetaData { get; set; }
        public List<Dictionary<int, string>> Items { get; set; }
    }

    public class DataTableMetaData
    {
        public string Description { get; set; }
        public string DateCreated { get; set; }
        public List<MetaDataField> Fields { get; set; }
    }

    public class MetaDataField
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public List<string> AdditionalInfo { get; set; }
    }
}
