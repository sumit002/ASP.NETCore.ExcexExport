using System;
using System.Data;

namespace ExportToExcelConsoleApp
{
    public class RawData : IRawData
    {
        public object GetDataObject()
        {
            throw new NotImplementedException();
        }

        public DataSet GetDataSet()
        {
            #region DataTable Creation
            DataTable dt = new DataTable("Users");
            dt.Clear();
            dt.Columns.Add("Id");
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Email");
            dt.Columns.Add("PhoneNumber");
            #endregion

            Random rnd = new Random();
            int rows = rnd.Next(1, 10);

            for (int i = 1; i <= rows; i++)
            {
                DataRow drUserData = dt.NewRow();
                drUserData["Id"] = Guid.NewGuid().ToString();
                drUserData["FirstName"] = $"FirstName{i}";
                drUserData["LastName"] = "LastName";
                drUserData["Email"] = $"{drUserData["FirstName"]}{i}@gmail.com";
                drUserData["PhoneNumber"] = $"786787899{i}";
                dt.Rows.Add(drUserData);
            }

            DataTable dtGraphData = new DataTable("Graph");
            dtGraphData.Clear();
            dtGraphData.Columns.Add("TimeSpan");
            dtGraphData.Columns.Add("Measurement1");
            dtGraphData.Columns.Add("Temp");
            dtGraphData.Columns.Add("Arsenic");

            for (int i = 1; i <= rows; i++)
            {
                DataRow drGraphData = dtGraphData.NewRow();
                drGraphData["TimeSpan"] = DateTime.Now.AddDays(-i).Ticks.ToString();
                drGraphData["Measurement1"] = "23.670089766";
                drGraphData["Temp"] = "223.5670000";
                drGraphData["Arsenic"] = "45.785670000";
                dtGraphData.Rows.Add(drGraphData);
            }

            // Adding DataTable
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.Tables.Add(dtGraphData);

            return ds;
        }

        public string GetJson()
        {
            string jsonData = "{\"Name\":\"UserTenant\",\"DataTables\":[{\"Name\":\"Tenant Data\",\"MetaData\":{\"Description\":\"SqlDataAdapter\",\"DateCreated\":\"12/03/2020 14:28:14\",\"Fields\":[{\"Index\":0,\"Name\":\"TenantId\",\"Type\":\"Guid\",\"AdditionalInfo\":[]},{\"Index\":1,\"Name\":\"Street\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":2,\"Name\":\"Postal Code\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":3,\"Name\":\"City\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":4,\"Name\":\"Country\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":5,\"Name\":\"Plant Name\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":6,\"Name\":\"Email Address 1\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":7,\"Name\":\"Email Address 2\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":8,\"Name\":\"Contract Number\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":9,\"Name\":\"Contact Name\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":10,\"Name\":\"Account Name\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":11,\"Name\":\"Phone Number\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":12,\"Name\":\"Mobile Number\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":13,\"Name\":\"Account Number\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":14,\"Name\":\"Start Date\",\"Type\":\"DateTime\",\"AdditionalInfo\":[]},{\"Index\":15,\"Name\":\"End Date\",\"Type\":\"DateTime\",\"AdditionalInfo\":[]},{\"Index\":16,\"Name\":\"Creation Date\",\"Type\":\"DateTime\",\"AdditionalInfo\":[]}]},\"Items\":[{\"0\":\"e86c76fa-14fe-4158-ae18-c6ba0ec46fee\",\"1\":\"400 Britannia Rd E \",\"2\":\"L4Z 1X9\",\"3\":\"Mississauga\",\"4\":\"Canada\",\"5\":\"Ola Canada\",\"6\":\"melissa123_456@ola.com\",\"7\":\"paul123_456@ola.com\",\"8\":\"22446688\",\"9\":\"Melissa S\",\"10\":\"Melissa S\",\"11\":\"4567811134\",\"12\":\"4567811134\",\"13\":\"135792468\",\"14\":\"11/1/2020 6:39:00 PM\",\"15\":\"11/30/2020 6:39:00 PM\",\"16\":\"12/3/2020 2:28:14 PM\"}]},{\"Name\":\"Measured Data\",\"MetaData\":{\"Description\":\"MeasurementDataAdapter\",\"DateCreated\":\"12/03/2020 14:28:14\",\"Fields\":[{\"Index\":0,\"Name\":\"TimeSeries\",\"Type\":\"String\",\"AdditionalInfo\":[]},{\"Index\":1,\"Name\":\"PARAMETERTYPE.MEASUREMENT\",\"Type\":\"String\",\"AdditionalInfo\":[\"?g/L\"]}]},\"Items\":[{\"0\":\"637410816000000000\",\"1\":\"305.26373291015625\"},{\"0\":\"637410916000000000\",\"1\":\"300.26373291015625\"},{\"0\":\"637310916000000000\",\"1\":\"310.26373291015625\"}]}]}";
            return jsonData;
        }
    }
}
