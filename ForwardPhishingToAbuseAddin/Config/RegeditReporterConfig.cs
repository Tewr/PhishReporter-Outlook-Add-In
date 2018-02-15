using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ForwardPhishingToAbuseAddin.Logging;
using ForwardPhishingToAbuseAddin.Services;
using Microsoft.Win32;

namespace ForwardPhishingToAbuseAddin.Config
{
	[SuppressMessage("ReSharper", "UnassignedGetOnlyAutoProperty", Justification =
		"All interface properties are instanciated with reflection")]
	[SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Local", Justification =
		"All interface properties are instanciated with reflection")]
	public class RegeditReporterConfig : IPhisingReporterConfig
	{
		private static readonly IApplicationInfo AppInfo = ServiceProvider.AppInfo;
		private static readonly IErrorLogger Log = ServiceProvider.Log;

		public static readonly string RegistryKey = $@"{Registry.LocalMachine.Name}\SOFTWARE\{GetCompany()}\{GetProductName()}\";

		public RegeditReporterConfig(IPhisingReporterConfig fallback)
		{
			foreach (var interfaceProperty in GetFallbackValuesAsDictionary(fallback))
				GetType().GetProperty(interfaceProperty.Key)?.SetValue(this,
					GetValue(interfaceProperty.Key, interfaceProperty.Value), null);
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

		public static IDictionary<string, object> GetFallbackValuesAsDictionary(IPhisingReporterConfig dataSource)
		{
			return typeof(IPhisingReporterConfig).GetProperties().ToDictionary(x => x.Name, x => x.GetValue(dataSource, null));
		}

		private object GetValue(string propertyName, object fallBack)
		{
			string value = null;
			try
			{
				value = (string) Registry.GetValue(RegistryKey, propertyName, fallBack);
			}
			catch (Exception e)
			{
				Log.LogError(() => $"{nameof(RegeditReporterConfig)}.{nameof(GetValue)}({propertyName}, {fallBack}): " +
									 $"Call to Registry.GetValue({RegistryKey}, {propertyName}, {fallBack}) failed", e, ErrorCodes.FailedToReadRegistry);
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