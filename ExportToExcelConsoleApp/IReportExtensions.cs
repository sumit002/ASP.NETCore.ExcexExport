using System.Data;

namespace ExportToExcelConsoleApp
{
    internal interface IReportExtensions
    {
        /// <summary>
        /// Export DataSet To Excel(.xlsx)
        /// </summary>
        /// <param name="ds">DataSet</param>
        /// <returns>ByteArray MemoryStream of the Excel</returns>
        byte[] ExportToExcel(DataSet ds);
    }
}
