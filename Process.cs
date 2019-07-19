using System.Data;
using System.Collections;


namespace StockStatusMPOUS
{
    class Process
    {
        private ArrayList locations = new ArrayList();
        private string colHeaders = "";

        public Process(string itemLocs)
        {
            string[] location = itemLocs.Split(",".ToCharArray());
            SetColHeaders();
            foreach (string item in location)
                locations.Add(item);
            FillStockStatus();
        }

        private void FillStockStatus()
        {
            DataSet ds;
            DataManager dm = new DataManager();
            OutputManager om = new OutputManager();
            om.ColHeaders = colHeaders;
            foreach(string loc in locations)
            {
                ds = new DataSet();
                ds = dm.GetStockStatus(loc);
                om.DSet = ds;
                om.Location = loc; //this will be the name for the worksheet
                om.CreateSpreadsheet();            
            }
            om.SaveSpreadSheet();
        }

        private void SetColHeaders()
        {
            colHeaders = "Item" + "|";
            colHeaders += "Description" + "|";
            colHeaders += "On Hand" + "|";      
            colHeaders += "Min" + "|"; 
            colHeaders += "Max" + "|"; 
            colHeaders += "IUOM" + "|"; 
            colHeaders += "Billable" + "|"; 
            colHeaders += "Active" + "|"; 
            colHeaders += "Consigned" + "|";
            colHeaders += "MFR Cat No" + "|"; 
            colHeaders += "Issue Cost" + "|";
            colHeaders += "Total Cost" + "|";
            colHeaders += "RFID" + "|";
            colHeaders += "Primary Bin" + "|";
        }
    }
}
