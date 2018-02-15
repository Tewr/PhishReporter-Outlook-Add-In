using ForwardPhishingToAbuseAddin.Config;
using ForwardPhishingToAbuseAddin.Logging;
using ForwardPhishingToAbuseAddin.Properties;

namespace ForwardPhishingToAbuseAddin.Services
{
	public static class ServiceProvider
	{
		private static IPhisingReporterConfig _config;
		private static IErrorLogger _log;
		private static IApplicationInfo _appInfo;

		public static IApplicationInfo AppInfo => _appInfo ?? (_appInfo = new AssemblyProperties());

		public static IPhisingReporterConfig Config =>
			_config ?? (_config = new RegeditReporterConfig(new ResourcesPhishingReporterConfig()));

		public static IErrorLogger Log => _log ?? (_log = new EventLogErrorLogger());
	}
}