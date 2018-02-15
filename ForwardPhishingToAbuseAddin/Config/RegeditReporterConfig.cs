using System;
using System.Diagnostics.CodeAnalysis;
using ForwardPhishingToAbuseAddin.Services;
using Microsoft.Win32;

namespace ForwardPhishingToAbuseAddin.Config
{
	[SuppressMessage("ReSharper", "UnassignedGetOnlyAutoProperty", Justification = "All interface properties are instanciated with reflection")]
	[SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Local", Justification = "All interface properties are instanciated with reflection")]
	public class RegeditReporterConfig : IPhisingReporterConfig
	{
		private static readonly IApplicationInfo AppInfo = ServiceProvider.AppInfo;
		private static readonly IErrorLogger Log = ServiceProvider.Log;

		private static readonly string RegeditApplicationPathInLocalMachine = $@"SOFTWARE\{GetCompany()}\{GetProductName()}\";

		public RegeditReporterConfig(IPhisingReporterConfig fallback)
		{
			foreach (var interfaceProperty in typeof(IPhisingReporterConfig).GetProperties())
			{
				GetType().GetProperty(interfaceProperty.Name)?.SetValue(this, 
					GetValue(interfaceProperty.Name, interfaceProperty.GetValue(fallback, null)), null);
			}
		}

		public string ReportingButtonLabel { get; private set; }
		public string ReportingButtonScreenTip { get; private set; }
		public string ReportingButtonSuperTip { get; private set; }
		public string RibbonGroupName { get; private set; }
		public string SecurityTeamEmail { get; private set; }
		public string ReportingEmailSubject { get; private set; }
		public string ReportingConfirmationMessage { get; private set; }
		public string ReportingConfirmationTitle { get; private set; }
		public string SecurityTeamEmailBody { get; private set; }
		public string RunbookUrl { get; private set; }
		public string NoEmailSelectedMessage { get; private set; }
		public string NoEmailSelectedTitle { get; private set; }
		public string ProcessAddendumIfRunbookUrlConfigured { get; private set; }

		private object GetValue(string propertyName, object fallBack)
		{
			
			var registryKey = $@"{Registry.LocalMachine.Name}\{RegeditApplicationPathInLocalMachine}";

			string value = null;
			try
			{
				value = (string)Registry.GetValue(registryKey, propertyName, fallBack);
			}
			catch (Exception e)
			{
				Log.LogError(() => $"{nameof(RegeditReporterConfig)}.{nameof(GetValue)}({propertyName}, {fallBack}): " +
											$"Call to Registry.GetValue({registryKey}, {propertyName}, {fallBack}) failed", e);
			}

			return string.IsNullOrWhiteSpace(value) ? fallBack : value;
		}

		private static string GetProductName()
		{
			var assemblyProduct = AppInfo.ApplicationProduct;
			return !string.IsNullOrEmpty(assemblyProduct) ? assemblyProduct : "Outlook.ForwardPhishingToAbuseAddin";
		}

		private static string GetCompany()
		{
			var assemblyCompany = AppInfo.ApplicationCompany;
			return !string.IsNullOrEmpty(assemblyCompany) ? assemblyCompany : "MyCompany";
		}
	}
}
