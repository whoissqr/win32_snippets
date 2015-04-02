@@ -1,387 +0,0 @@
﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
using Npgsql;
using System.Reflection;

namespace XapTesterStatus
{
    public partial class Service1 : ServiceBase
    {
        public string eventSource;
        public EventLog eventLog;
        public Timer timer;
        private bool done4ThisWeek = false;
        EmailHelper emailHelper;
        bool debug = false;

        public string CurrentVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public Service1()
        {
            InitializeComponent();
            
            eventLog = new EventLog();
            eventSource = ConfigurationManager.AppSettings["eventSource"].ToString();
            if (!EventLog.SourceExists(eventSource))
            {
                EventLog.CreateEventSource(eventSource, eventSource);
            }
            eventLog.Source = eventSource;
            eventLog.EnableRaisingEvents = true;
            eventLog.WriteEntry("OnInit: ReportGenerator v " + CurrentVersion, EventLogEntryType.Warning);
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                string mode = ConfigurationManager.AppSettings["mode"].ToString();
                debug = (mode == "test") ? true : false;

                emailHelper = new EmailHelper();
                int interval = 60;
                if (int.TryParse(ConfigurationManager.AppSettings["Interval"].ToString(), out interval))
                {
                    eventLog.WriteEntry("OnStart with interval of every " + ConfigurationManager.AppSettings["Interval"].ToString() + " minutes;", EventLogEntryType.Information);
                    int dueTime = 1000*30; //service starts in 30 secs upon installation
                    interval = interval * 60 * 1000;
                    timer = new Timer(new TimerCallback(timer_Elapsed));
                    timer.Change(dueTime, interval);                    
                }
                else
                {
                    eventLog.WriteEntry("Interval setting in config file is not correct!", EventLogEntryType.Warning);
                    OnStop();
                }
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.Source + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, EventLogEntryType.Error);
                OnStop();
            }
        }

        protected override void OnStop()
        {
            eventLog.WriteEntry("OnStop", EventLogEntryType.Information);
        }

        void timer_Elapsed(object sender)
        {
            try
            {
                if (!debug)
                {
                    if (DateTime.Today.DayOfWeek != DayOfWeek.Monday) { done4ThisWeek = false; return; }          //only runs on Monday each week
                    if ((DateTime.Now.Hour < 9) || (DateTime.Now.Hour > 14)) { done4ThisWeek = false; return; }

                    eventLog.WriteEntry("Elapse at " + DateTime.Now.ToShortTimeString(), EventLogEntryType.Information);
                
                    if (done4ThisWeek) return;
                }

                eventLog.WriteEntry("Elapse; going to run report;", EventLogEntryType.Information);   

                if (DoWork())
                {
                    done4ThisWeek = true;
                    eventLog.WriteEntry("Elapse; done everything for this week;", EventLogEntryType.Information);
                }
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.Source + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, EventLogEntryType.Error);
                OnStop();
            }
            finally
            {
             
            }
        }
        
        //The scheduled task to run every Monday, 9am - 2pm.
        private bool DoWork(){
            bool success = false;
            DateTime reportStartDt = getStartOfLastWeek();
            DateTime reportEndDt = reportStartDt.AddDays(7);
            DBHelper dbHelper = new DBHelper();
            var OSAT_report = new List<KeyValuePair<string, string>>();

            //$reportDir here is a default setting
            //App.config takes priority if exists
            string reportDir = @"D:\reports\MPRS_Auto_weekly_report\";
            if (ConfigurationManager.AppSettings["reportDirectory"] != null)
                reportDir = ConfigurationManager.AppSettings["reportDirectory"].ToString();

            if (!reportDir.EndsWith(@"\")) reportDir = reportDir + @"\";

            try
            {
                for (int siteIndex = 0; siteIndex < 4; siteIndex++)
                {
                    OSAT osat = (OSAT)siteIndex;
                    int weekNoOfLastWeek = GetWeekOfYear(DateTime.Now) - 1;
                    string reportName = reportDir + osat.ToString() + "-Weekly-Report-W" + weekNoOfLastWeek + "-Y" + (DateTime.Now.Year-2000) + ".xlsx";
                    //implementation of StdReportCreator omitted
                    List<string> testerList = dbHelper.GetTesterListBySite(siteIndex, reportStartDt);
                    StdReportCreator stdReportCreator = new StdReportCreator(reportStartDt, reportEndDt, 7, testerList, reportName, osat);
                    stdReportCreator.CreateStandardProductionReport();
                    OSAT_report.Add(new KeyValuePair<string, string>(osat.ToString(), reportName));
                }
                eventLog.WriteEntry("Elapse; going to send email;", EventLogEntryType.Information);
                SendNotificationEmails(OSAT_report);
                success = true;
            }
            catch (Exception ex)
            {
                success = false;
                eventLog.WriteEntry(ex.Source + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, EventLogEntryType.Error);
                OnStop();             
            }
            
            return success;
        }

        private void SendNotificationEmails(List<KeyValuePair<string, string>> OSAT_report)
        {
            //implementation of ExcelReader omitted
            ExcelReader er = new ExcelReader();
            
            string subject = "*** Automatic Weekly Standard OEE Report ***";
            bool ishtml = true;

            string body = "";
            body += ("Good morning, " + "<br/><br/>");
            body += ("This is an automatic generated weekly OEE report email.<br/><br/>");
            body += "<html>";
            body += "<head>";
            body += "<style type=\"text/css\">";
            body += ".myTable th, td { border: none solid black; border-collapse: collapse; padding-left: 5px; padding-right: 10px;}";
            body += ".myTable th { color:white; }";

            body += "</style>";
            body += "</head>";
            body += "<body>";
            body += "<table class=\"myTable\">";
            body += "<tr>";
            body += "<th bgcolor=\"#0000CC\">Factory</th>";
            body += "<th bgcolor=\"#0000CC\">Tester</th>";
            body += "<th bgcolor=\"#006600\">Earn Hrs</th>";
            body += "<th bgcolor=\"#660099\">Mfg Hrs</th>";
            body += "<th bgcolor=\"#660099\">Rt Hrs</th>";
            body += "<th bgcolor=\"#660099\">Verification</th>";
            body += "<th bgcolor=\"#660099\">Qce Hrs</th>";
            body += "<th bgcolor=\"#660099\">Setup</th>";
            body += "<th bgcolor=\"#660099\">Down</th>";
            body += "<th bgcolor=\"#660099\">PM</th>";
            body += "<th bgcolor=\"#006600\">Others</th>";
            body += "<th bgcolor=\"#006600\">MTE</th>";
            body += "<th bgcolor=\"#006600\">PTE</th>";
            body += "<th bgcolor=\"#006600\">Idle</th>";
            body += "<th bgcolor=\"#660099\">Shutdown</th>";
            body += "<th bgcolor=\"#660099\">Unknown</th>";
            body += "<th bgcolor=\"#0000CC\">Total H</th>";
            body += "<th bgcolor=\"#0000CC\">OEE %</th>";
            body += "<th bgcolor=\"#006600\">xOEE %</th>";
            body += "</tr>";

            char[] delimiterChars = { ';' };
            List<string> reportList = new List<string>();
            
            foreach (var osat_report_pair in OSAT_report)
            {
                KeyValuePair<string, string> pair = osat_report_pair;
                string osat = pair.Key;
                string reportName = pair.Value;
                reportList.Add(reportName);

                List<string> rows = new List<string>();

                //step1. extract data from excel
                er.readRows(reportName, 2, 2, 4, rows);

                //step2. upload data to DB
                string years = "";
                string ww = "";
                FileInfo f = new FileInfo(reportName);
                getOSATinfoFromFileName(f.Name, ref osat, ref years, ref ww);
                upload2DB(osat, years, ww, rows);

                //step3. send summary to email 
                bool OsatCellDone = false;
                foreach (string row in rows)
                {
                   
                    if (!OsatCellDone)
                    {
                        body += "<tr bgcolor=\"#99CCFF\">";
                        body += "<td>" + osat + "</td>";
                        OsatCellDone = true;
                    }
                    else 
                    {
                        body += "<tr bgcolor:\"#FFFFFF\">";
                        body += "<td> </td>";                    
                    }
                    string[] cells = row.Split(delimiterChars);
                    foreach (string cell in cells)
                    {
                        body += "<td>" + cell + "</td>";
                    }
                    body += "</tr>";
                }
            }
 
            body += "</table></body></html><br/><br/>";
            body += ("Please feedback if there\'s any issue<br/><br/>");
            body += ("Best Regards,<br/>Administrator");

            emailHelper.SendEmailsXAP(subject, body, ishtml, reportList);
        }

        public bool upload2DB(string osat, string years, string ww, List<string> rows) {
            string[] field = new string[21];
            field[0] = osat;
            field[1] = years;
            field[2] = ww;
            foreach (string row in rows)
            {
                string[] str = row.Split(new char[] { ';' });
                if (str[0].ToUpper().Contains("93K"))
                {
                    str[0] = "93K";
                }
                else if (str[0].ToUpper().Contains("T2K"))
                {
                    str[0] = "T2K";
                }

                for (int i = 3; i < field.Length; i++)
                {
                    field[i] = str[i - 3];
                }

                string sqlstr = getSQL2Insert(field);
                if (ExecuteSQL(sqlstr))
                {
                    Debug.WriteLine(sqlstr + ": success!");
                }
                else
                {
                    Debug.WriteLine(sqlstr + ": fail!");
                    eventLog.WriteEntry("Fail to execute SQL statement " + sqlstr, EventLogEntryType.Error);
                }
            }//ends foreach

            return true;
        }

        static void getOSATinfoFromFileName(string fileName, ref string osat, ref  string years, ref string ww) {
            //e.g. fileName = "ATK-Weekly-Report-W2-15.xlsx";
            string[] split = fileName.Split(new Char[] { '-' });

            osat = split[0].ToUpper();

            //Y2015.xlsx
            years = split[4];
            years = years.Remove(years.IndexOf('.'));
            if (years.ToUpper().StartsWith("Y")) years = years.Substring(1);
            if (years.Length == 2) years = "20" + years;              //unit test

            //W2
            ww = split[3];
            if (ww.ToUpper().StartsWith("W")) ww = ww.Substring(1);
            ww = ww.TrimStart('0');
        }

        public DateTime getStartOfLastWeek() {
            DayOfWeek weekStart = DayOfWeek.Sunday; // or Sunday, or whenever
            DateTime startingDate = DateTime.Today;

            while (startingDate.DayOfWeek != weekStart)
                startingDate = startingDate.AddDays(-1);

            return startingDate.Date.AddDays(-7).AddHours(7);          
        }

        public int GetWeekOfYear(DateTime time)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;

            return cal.GetWeekOfYear(time, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
        }

        string getSQL2Insert(string[] strArray) {
            string sqlstr = "INSERT INTO OEE values (";
            foreach(string str in strArray){
                string tempStr = str;
                if (tempStr.Contains('%'))
                {
                    tempStr = tempStr.Trim(new char[] {'%'});
                    tempStr = float.Parse(tempStr).ToString("N4");
                }
                sqlstr += "\'" + tempStr + "\',";
            }

            sqlstr = sqlstr.Trim(new char[]{','});
            sqlstr += ")";
            return sqlstr;
        }

        bool ExecuteSQL(string sqlstr) {
            bool done = false;
 
            NpgsqlConnection conn = new NpgsqlConnection(DBHelper.getDBConnString("serverName"));
            NpgsqlCommand comm = null;
                            
            try
            {
                conn.Open();
                comm = new NpgsqlCommand(sqlstr, conn);
                comm.ExecuteNonQuery();
                done = true;
            }
            catch (NpgsqlException e)
            {
                if (e.Code != "23505")
                {
                    Debug.WriteLine(e.Message);
                    done = false;
                }
                else 
                {
                    done = true;  //when duplicate copy exists
                }                 
            }
            finally
            {
                if (comm != null)
                {
                    comm.Dispose();
                }
                if (conn != null)
                {
                    conn.Dispose();
                }
            }
            
            return done;
        }
    }
}
