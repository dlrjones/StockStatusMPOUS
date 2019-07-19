using System;
using System.Data;
using OleDBDataManager;
using LogDefault;
using DTUtilities;
using KeyMaster;
using System.Collections.Specialized;
using System.Configuration;
using SpreadsheetLight;
using System.IO;
using System.Collections;
using System.Net.Mail;

namespace StockStatusMPOUS
{
    class OutputManager
    {
        private static LogManager lm = LogManager.GetInstance();
        private ODMDataFactory ODMDataSetFactory = null;
        private NameValueCollection ConfigSettings = null;
        private static ArrayList skipCols = new ArrayList();
        private DateTimeUtilities dtu = new DateTimeUtilities();
        private string outFilePath = "";
        private string outFileName = "";
        private string location = "";
    //    private string sheetName = "";
        private string emailList = "";
        private string emailCopyTo = "";
        private string emailReplyTo = "";
        private string firstName = "";
        private DataSet dSet = new DataSet();
        private string colHeaders = "";
        private int worksheetCount = 0;
   //     private  OutputManager outMngr = null;
        private  SLDocument sldStockStatus = null;

        public string Location
        {
            set { location = value; }
        }

        public string ColHeaders
        {
            set { colHeaders = value; }
        }       

        public DataSet DSet
        {
            set { dSet = value; }
        }

        public OutputManager()
        {
            ODMDataSetFactory = new ODMDataFactory();
            //ConfigSettings = (NameValueCollection)ConfigurationSettings.GetConfig("appSettings");
            ConfigSettings = (NameValueCollection)ConfigurationManager.GetSection("appSettings");
            lm.LogFile = ConfigSettings.Get("logFile");
            lm.LogFilePath = ConfigSettings.Get("logFilePath");
            outFilePath = ConfigSettings.Get("outFilePath");
            emailList = ConfigSettings.Get("email_to_list");
            emailCopyTo = ConfigSettings.Get("email_copy_to");
            emailReplyTo = ConfigSettings.Get("email_reply_to"); 
            firstName = ConfigSettings.Get("email_to_first_name");
            SetExcludedColsList();
        }
 
        private void SetExcludedColsList()
        {
            for (int x = 0; x < 26; x++) {
                switch (x)
                {
                    case 0:
                        skipCols.Add(x);
                        break;
                    case 1:
                        skipCols.Add(x);
                        break;
                    case 2:
                        skipCols.Add(x);
                        break;
                    case 3:
                        skipCols.Add(x);
                        break;
                    case 4:
                        skipCols.Add(x);
                        break;
                    case 5:
                        skipCols.Add(x);
                        break;
                    case 7:
                        skipCols.Add(x);
                        break;
                    case 19:
                        skipCols.Add(x);
                        break;
                    case 20:
                        skipCols.Add(x);
                        break;
                    case 21:
                        skipCols.Add(x);
                        break;
                    case 23:
                        skipCols.Add(x);
                        break;
                    case 24:
                        skipCols.Add(x);
                        break;
                    case 25:
                        skipCols.Add(x);
                        break;
                }
            }
                    
        }
        public void CreateSpreadsheet()
        {
            int dataColNo = 0;
            int colNo = 1;
            int rowNo = 1;
            if(sldStockStatus == null)
                sldStockStatus = new SLDocument();
            try
            {
                if (worksheetCount++ > 0)
                    sldStockStatus.AddWorksheet(location);
                else
                    sldStockStatus.RenameWorksheet(SLDocument.DefaultFirstSheetName, location);

                SetColHeaders();
                foreach (DataRow dRow in dSet.Tables[0].Rows)
                {
                    dataColNo = 0;
                    colNo = 1;
                    rowNo++;
                    foreach (object colData in dRow.ItemArray)
                    {
                        if (!skipCols.Contains(dataColNo++)){
                            sldStockStatus.SetCellValue(rowNo, colNo++, colData.ToString());
                        }
                    }
                }                                        
              //  sldStockStatus.SaveAs(outFilePath);
            }
            catch (Exception ex)
            {
                lm.Write(location + "  OutputManager.CreateSpreadsheet() " + ex.Message);
            }
        }
        public void SaveSpreadSheet()
        {
            outFileName = outFilePath + "StockStatus" + dtu.DateTimeCoded() + ".xlsx";
            sldStockStatus.SaveAs(outFileName);
            SendMail();
        }

        private void SetColHeaders()
        {
            int rowNo = 1;
            int colNo = 1;
            string[] colNames = colHeaders.Split("|".ToCharArray());
            foreach (string cname in colNames)
            {
                sldStockStatus.SetCellValue(rowNo, colNo++, cname);
            }
        }

        private void SendMail()
        {
            string[] mailList = emailList.Split(";".ToCharArray());
            string[] ccList = emailCopyTo.Split(";".ToCharArray());
            try
            {
                foreach (string recipient in mailList)
                {
                    if (recipient.Trim().Length > 0)
                    {
                        MailMessage mail = new MailMessage();
                        SmtpClient SmtpServer = new SmtpClient("smtp.uw.edu");
                        mail.To.Add(recipient);
                        mail.From = new MailAddress("pmmhelp@uw.edu");
                        if (emailCopyTo.Length > 0)
                        {
                            foreach (string cc in ccList)
                            {
                                mail.CC.Add(cc);
                            }
                        }
                        mail.Subject = "Stock Status Report for " + dtu.DateTimeToShortDate(DateTime.Now);
                        mail.Body = (firstName.Length > 0
                            ? firstName + "," + Environment.NewLine + Environment.NewLine
                            : "") +
                                    "Here's the MPOUS Stock Status Report for " + dtu.DateTimeToShortDate(DateTime.Now) +
                                    Environment.NewLine +
                                    Environment.NewLine +
                                    Environment.NewLine +
                                    Environment.NewLine +
                                    Environment.NewLine +
                                    "PMMHelp" + Environment.NewLine +
                                    "UW Medicine Harborview Medical Center" + Environment.NewLine +
                                    "Supply Chain Management Informatics" + Environment.NewLine +
                                    "206-598-0044" + Environment.NewLine +
                                    "pmmhelp@uw.edu";
                        mail.ReplyToList.Add(emailReplyTo);

                        Attachment attachment;
                        attachment =
                            new System.Net.Mail.Attachment(outFileName);

                        mail.Attachments.Add(attachment);

                        SmtpServer.Port = 587;
                        SmtpServer.Credentials = new System.Net.NetworkCredential("pmmhelp", GetKey());
                        SmtpServer.EnableSsl = true;
                        SmtpServer.Send(mail);
                        lm.Write("Process/SendMail:  " + recipient);
                        if (emailCopyTo.Length > 0)
                            lm.Write("Process/Send_Mail/CC:  " + emailCopyTo);
                    }
                }
            }
            catch (Exception ex)
            {
                string mssg = ex.Message;
                lm.Write("Process/SendMail_:  " + mssg);
            }
        }

        protected string GetKey()
        {
            string[] key = File.ReadAllLines(outFilePath + "status.txt");
            return StringCipher.Decrypt(key[0], "pmmhelp");
        }
    }
}
