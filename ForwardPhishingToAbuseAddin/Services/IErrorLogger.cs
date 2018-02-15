using System;

namespace ForwardPhishingToAbuseAddin.Services
{
	public interface IErrorLogger
	{
		void LogError(Func<string> constructErrorMessage, Exception exception);

		void LogError(Func<string> constructErrorMessage, Exception exception, ushort errorCode);
		
	}
}