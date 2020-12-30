using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bunker.Web.Services
{
	public static class XtraLogHelper
	{
		//EventLog eventLog = new EventLog(
		private static string LOG_SOURCE = "XtraPM";
		private const string LOG_NAME = "Application";

		private static void RegisterEventSource()
		{
			if(!EventLog.SourceExists(LOG_SOURCE))
			{
				EventLog.CreateEventSource(LOG_SOURCE, LOG_NAME);
			}
		}
		private static void RegisterEventSource(string src)
		{
			if(!EventLog.SourceExists(src))
			{
				EventLog.CreateEventSource(src, LOG_NAME);
			}
		}

		private static void WriteLog(string message, EventLogEntryType logEntryType)
		{
			RegisterEventSource();
			EventLog.WriteEntry(LOG_SOURCE, message, logEntryType);
		}

		public static void UseApplicationLog(bool isDefault = true)
		{
			if(!isDefault)
			{
				LOG_SOURCE = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
			}
			RegisterEventSource();
		}

		public static void ErrorLog(string message)
		{
			WriteLog(message, EventLogEntryType.Error);
		}

		public static void WarningLog(string message)
		{
			WriteLog(message, EventLogEntryType.Warning);
		}

		public static void InfoLog(string message)
		{
			WriteLog(message, EventLogEntryType.Information);
		}

		public static void ExeceptionLog(Exception ex,bool defaultSource = true)
		{
			string message = string.Format("{0} StackTrace: {1}", ex.Message, ex.StackTrace);
			if(defaultSource)
			{
				WriteLog(message, EventLogEntryType.Error);
			}
			else
			{
				RegisterEventSource(ex.Source);
				EventLog.WriteEntry(ex.Source, message, EventLogEntryType.Error);
			}
		}
	}
}
