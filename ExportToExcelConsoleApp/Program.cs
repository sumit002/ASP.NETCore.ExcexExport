using Newtonsoft.Json;
using System;
using System.Data;
using System.Linq;

namespace ExportToExcelConsoleApp
{
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Welcome to Export to Excel.");

            var closedXml = new ExportUsingClosedXml();
            IRawData rawData = new RawData();
            //var result = closedXml.ExportToExcel(rawData.GetDataSet());

            closedXml.ExportToExcelFile(JsonToDataTable(rawData.GetJson()));

            Console.WriteLine("ExportToExcel Successful.");
        }

        private static DataSet JsonToDataTable(string json)
        {
            var ds = new DataSet();

            var dataSetJson = JsonConvert.DeserializeObject<DataSetJson>(json);
            foreach (var datatable in dataSetJson.DataTables)
            {
                var dt = new DataTable(datatable.Name);

                // Taking the Fields index and adding to table Column
                for (var i = 0; i < datatable.MetaData.Fields.Count; i++)
                {
                    var columnName = datatable.MetaData.Fields.FirstOrDefault(x => x.Index == i)?.Name;
                    if (columnName != null && columnName.Contains("I18N|", StringComparison.InvariantCultureIgnoreCase))
                    {
                        // translate key and use instead

                    }

                    dt.Columns.Add(columnName);
                }

                // Taking the Items index and adding to table Row
                for (var r = 0; r < datatable.Items.Count; r++)
                {
                    var dr = dt.NewRow();
                    for (var i = 0; i < datatable.Items[r].Count; i++)
                        dr[i] = datatable.Items[r][i].Trim();
                    dt.Rows.Add(dr);
                }

                ds.Tables.Add(dt);
            }

            return ds;
        }
    }
}
