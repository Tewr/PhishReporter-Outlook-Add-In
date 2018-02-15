using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ForwardPhishingToAbuseAddin.Config;
using ForwardPhishingToAbuseAddin.Logging;
using ForwardPhishingToAbuseAddin.Services;

namespace DeploymentTools
{
	public static class Tools
	{
		public static void GenerateRegistryFileFromdefaultValues()
		{
			var registryTargetKey = RegeditReporterConfig.RegistryKey;
			var registryTargetValues = RegeditReporterConfig.GetFallbackValuesAsDictionary(new ResourcesPhishingReporterConfig());
			var appInfo = ServiceProvider.AppInfo;
			var sb = new StringBuilder();
			sb.AppendLine(@"Windows Registry Editor Version 5.00");
			sb.AppendLine($"[{registryTargetKey}]");
			sb.AppendLine(
				$";This file was generated on {DateTime.Now:F} using the Command ({typeof(Tools).Assembly.GetName().Name}) {nameof(GenerateRegistryFileFromdefaultValues)}");
			sb.AppendLine($";Values generated from version {appInfo.ApplicationProduct} v{appInfo.ApplicationVersion}");
			foreach (var registryTargetValue in registryTargetValues)
			{
				sb.AppendLine();
				sb.AppendLine($@"""{registryTargetValue.Key}""=""{Escape(registryTargetValue.Value)}""");
			}
			
			Console.Out.Write(sb.ToString());
		}

		public static void PrintErrorCodes()
		{
			var sb = new StringBuilder();
			sb.AppendLine(@"Possible error codes");
			sb.AppendLine(@"Code	Name");
			foreach (var registryTargetValue in ErrorCodes.All)
			{
				sb.AppendLine($@"{registryTargetValue.Key}	{registryTargetValue.Value}");
			}

			Console.Out.Write(sb.ToString());
		}

		private static string Escape(object registryTargetValue)
		{
			return registryTargetValue.ToString().Replace("\\", "\\\\").Replace("\"", "\\\"");
		}
	}
}
