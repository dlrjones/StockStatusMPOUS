using System;
using System.Data;
using OleDBDataManager;
using LogDefault;
using System.Collections.Specialized;
using System.Configuration;

namespace StockStatusMPOUS
{
    class DataManager
    {
        private string connectString = "";
        private NameValueCollection ConfigData = null;
        private LogManager lm = LogManager.GetInstance();
        protected ODMDataFactory ODMDataSetFactory = new ODMDataFactory();

        public DataSet GetStockStatus(string location)
        {
            if(connectString.Length == 0)
                GetConnectString();
            string sql = BuildQuery(location);
            return BuildDataSet(sql);
        }

        private string BuildQuery(string loc)
        {
            string sqlQuery = "exec up_rpt_StockStatus;1 '', '', '" + loc + "', '', '', '', '', '', '', '', ''";
            return sqlQuery;
        }

        private DataSet BuildDataSet(string sql)
        {
            DataSet dSet = new DataSet();
            ODMRequest Request = new ODMRequest();
            Request.ConnectString = connectString;
            Request.CommandType = CommandType.Text;
            Request.Command = sql;
     //       string itemNmbr = "";
            try
            {
                dSet = ODMDataSetFactory.ExecuteDataSetBuild(ref Request);
            }
            catch (Exception ex)
            {
                lm.Write("DataManager.GetData: " + ex.Message);
            }
            return dSet;
        }

        private void GetConnectString()
        {
         //   ConfigData = (NameValueCollection)ConfigurationSettings.GetConfig("appSettings");
            ConfigData = (NameValueCollection)ConfigurationManager.GetSection("appSettings");
            connectString = ConfigData.Get("cnctMPOUS");
        }

    }
}
