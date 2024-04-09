using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace PrintTestLabel
{

    [ComVisible(true)]
    [ProgId("PrintTestLabel.PrintTestLabel")]
    public class PrintTestLable : IWorkflowExtension
    {

        INautilusServiceProvider sp;
        private const string Type = "3";
        private int _port = 9100;
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private IDataLayer dal;
        public void Execute(ref LSExtensionParameters Parameters)
        {


            try
            {


                #region param
                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                var workstationId = Parameters["WORKSTATION_ID"];
                var resultName = rs.Fields["NAME"].Value;
                var resultId = (long)rs.Fields["RESULT_ID"].Value;
                #endregion

                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);


                dal = new DataLayer();
                dal.Connect();
                Workstation ws = dal.getWorkStaitionById(workstationId);
                ReportStation reportStation = dal.getReportStationByWorksAndType(ws.NAME, Type);
                //            string ip = GetIp(printerName);
                string goodIp = ""; //removeBadChar(ip);
                string printerName = "";
                if (reportStation != null)
                {


                    if (reportStation.Destination != null)
                    {
                        //מביא את הIP של המדפסת להדפסה הזאת
                        goodIp = reportStation.Destination.ManualIP;
                    }

                    if (reportStation.Destination != null && reportStation.Destination.RawTcpipPort != null)
                    {
                        //מביא את הפורט רק במקרה שהוא דיפולט
                        _port = (int)reportStation.Destination.RawTcpipPort;
                    }

                    Result result = dal.GetResultById(resultId);
                    Aliquot aliquot = result.Test.Aliquot;
                    var sampleName = aliquot.Sample.Name;
                    // var sampleID = aliq.Sample.Name;//TODO : לבדוק אם זה הערך הנכון
                    var testcode = "";
                    testcode = getTestCode(aliquot, testcode);
                    var mihol = result.DilutionFactor;

                    Print(sampleName, sampleName, testcode, mihol.ToString(), goodIp);

                }
                else
                    MessageBox.Show("לא הוגדרה מדפסת עבור תחנה זו.");


            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }

        }

        private string removeBadChar(string ip)
        {
            string ret = "";
            foreach (var c in ip)
            {
                int ascii = (int)c;
                if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
                    ret += c;
            }
            return ret;
        }

        private string getTestCode(Aliquot aliq, string testcode)
        {
            if (aliq.Parent.Count == 0)//&& aliq.U_CHARGE == "T")
            {
                testcode = aliq.ShortName;
            }
            if (aliq.Parent.Count != 0)
            {
                getTestCode(aliq.Parent.FirstOrDefault(), testcode);
            }
            return testcode;
        }

        public string GetIp(string printerName)
        {
            string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
            string ret = "";
            var searcher = new ManagementObjectSearcher(query);
            var coll = searcher.Get();
            foreach (ManagementObject printer in coll)
            {
                foreach (PropertyData property in printer.Properties)
                {
                    if (property.Name == "PortName")
                    {
                        ret = property.Value.ToString();
                    }
                }
            }
            return ret;
        }

        private static string ReverseString(string s)
        {
            Regex reg = new Regex(".*[א-ת].*");
            var replacedString = Regex.Replace(s, @".*[א-ת].*", m => new string(m.Value.Reverse().ToArray()));
            return replacedString;
        }

        private void Print(string name, string ID, string testcode, string mihol, string ip)
        {
            string ipAddress = ip;


            // ZPL Command(s)
            string ntxt = name;
            string tctxt = testcode;
            string mtxt = mihol;
            string itxt = ID;


            string ZPLString =
                 "^XA" +
                 "^LH0,0" +
                 "^FO20,10" +
                 "^A@N30,30" +
                string.Format("^FD{0}^FS", ntxt) + //שם
                 "^FO10,60" +
                 "^A@N30,30" +

                 string.Format("^FD{0}^FS", mtxt) +
                "^FO100,60" +
                 "^A@N30,30" +

                 string.Format("^FD{0}^FS", tctxt) +
                "^FO320,0" + "^BQN,4,3" +
                //string.Format("^FD   {0}^FS", itxt) + //ברקוד
                    string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
                "^XZ";
            try
            {
                // Open connection
                var client = new System.Net.Sockets.TcpClient();
                client.Connect(ipAddress, _port);
                // Write ZPL String to connection
                var writer = new StreamWriter(client.GetStream());
                writer.Write(ZPLString.Trim());
                writer.Flush();
                // Close Connection
                writer.Close();
                client.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
