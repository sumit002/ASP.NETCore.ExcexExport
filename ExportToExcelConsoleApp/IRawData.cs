using System.Data;

namespace ExportToExcelConsoleApp
{
    internal interface IRawData
    {
        /// <summary>
        /// Get DataSet Raw Data
        /// </summary>
        /// <returns>Sample DataSet</returns>
        DataSet GetDataSet();
        string GetJson();
        object GetDataObject();
    }
}
