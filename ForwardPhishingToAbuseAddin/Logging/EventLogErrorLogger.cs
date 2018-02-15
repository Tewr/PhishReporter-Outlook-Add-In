using System;
using System.Diagnostics;
using ForwardPhishingToAbuseAddin.Services;

namespace ForwardPhishingToAbuseAddin.Logging
{
	public class EventLogErrorLogger : IErrorLogger
	{
		private static readonly IApplicationInfo AppInfo = ServiceProvider.AppInfo;

		public void LogError(Func<string> constructErrorMessage, Exception exception)
		{
			LogError(constructErrorMessage, exception, 101);
		}
		public void LogError(Func<string> constructErrorMessage, Exception exception, ushort errorCode)
		{
			string message;
			try
			{
				message = constructErrorMessage();
			}
			catch (Exception e)
			{
				message = $"(Additional error when trying to construct error message: {e})";
			}

			using (EventLog eventLog = new EventLog("Application"))
			{
				var consistentCode = message.GetHashCode();
				message = $"{AppInfo.ApplicationProduct} {AppInfo.ApplicationVersion}: {Environment.NewLine} {exception.GetBaseException().GetType()} {message}{Environment.NewLine}{exception}";
				eventLog.Source = "Application";
				eventLog.WriteEntry(message, EventLogEntryType.Error, errorCode, 1);
			}
		}
	}
}