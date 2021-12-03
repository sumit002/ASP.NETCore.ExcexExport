using System;
using ClosedXML.Excel;
using System.Data;
using System.IO;

namespace ExportToExcelConsoleApp
{
    internal class ExportUsingClosedXml : IReportExtensions
    {
        public ExportUsingClosedXml()
        {
            Console.WriteLine("ExportUsingClosedXml Loaded!");
        }

        public byte[] ExportToExcel(DataSet ds)
        {
            var workbookBytes = new byte[0];
            try
            {
                var xLWorkbook = new XLWorkbook();
                foreach (DataTable dt in ds.Tables)
                    xLWorkbook.Worksheets.Add(dt, dt.TableName);

                using var ms = new MemoryStream();
                xLWorkbook.SaveAs(ms);
                workbookBytes = ms.ToArray();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in ExportToExcel", ex.ToString());
            }

            return workbookBytes;
        }

        public void ExportToExcelFile(DataSet ds)
        {
            try
            {
                var xLWorkbook = new XLWorkbook();
                foreach (DataTable dt in ds.Tables)
                    xLWorkbook.Worksheets.Add(dt, dt.TableName);

                using var ms = new MemoryStream();
                xLWorkbook.SaveAs(ms);

                string folderPath = "C:\\Excel\\";
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                xLWorkbook.SaveAs(folderPath + $"sample{DateTime.Now.Ticks}.xlsx");
                //Return xlsx Excel File  
                //return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Sample23");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in ExportToExcel", ex.ToString());
            }

        }
    }
}
