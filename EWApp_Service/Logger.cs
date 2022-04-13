using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWApp_Service
{
    public interface ILogger
    {
        bool WriteLog(string strMessage, EventLogEntryType Eventtype);
    }

    public class FileLogger: ILogger
    {
        public bool WriteLog(string strMessage , EventLogEntryType Eventtype)
        {
            try
            {
                string strFileName = "Logs.txt";
                FileStream objFilestream = new FileStream(string.Format("{0}\\{1}", ConfigurationManager.AppSettings["DBLocation"], strFileName), FileMode.Append, FileAccess.Write);
                StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
                objStreamWriter.WriteLine(strMessage);
                objStreamWriter.Close();
                objFilestream.Close();

                return true;

            }
            catch (Exception ex)
            {
                EventLogger FailedLog = new EventLogger();
                FailedLog.WriteLog(String.Format("{0} @ {1}", DateTime.Now,"Failed write Log into LogFile,  Error : "+ex.Message), EventLogEntryType.Error);
                return false;
                
            }
        }
    }

    public class EventLogger: ILogger
    { 
        public bool WriteLog(string message, EventLogEntryType Eventtype)
        {
            try
            {
               // string message = strMessage;
                if (!EventLog.SourceExists("EWApp"))
                {
                    EventLog.CreateEventSource("EWApp", "EWLog");
                    EventLog Elog = new EventLog("EWLog");
                    Elog.Source = "EWApp";
                    Elog.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, "LogEntry Created "), EventLogEntryType.Information);
                    Elog.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, message), Eventtype);
                }
                else
                {
                    EventLog Elog = new EventLog("EWLog");
                    Elog.Source = "EWApp";
                    Elog.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, message), Eventtype);
                }


                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, message), Eventtype);
                }

                return true;

            }catch (Exception ex)
            {
                return false ;
            }
        }
    }
}
