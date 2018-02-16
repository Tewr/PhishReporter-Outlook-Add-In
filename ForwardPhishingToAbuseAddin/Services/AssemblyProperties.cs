using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;

namespace ForwardPhishingToAbuseAddin.Services
{
	public class AssemblyProperties : IApplicationInfo
	{
		private static readonly ConcurrentDictionary<string, string> Values = new ConcurrentDictionary<string, string>();

		public string ApplicationProduct =>
			GetAttributeValue<AssemblyProductAttribute>(nameof(ApplicationProduct), x => x.Product);

		public string ApplicationVersion =>
			Values.GetOrAdd(nameof(ApplicationVersion), key => CurrentAssembly.GetName().Version.ToString());

		public string ApplicationCompany =>
			GetAttributeValue<AssemblyCompanyAttribute>(nameof(ApplicationCompany), x => x.Company);

		public static T GetAttribute<T>()
		{
			return (T) CurrentAssembly.GetCustomAttributes(typeof(T), false).FirstOrDefault();
		}

		private static Assembly CurrentAssembly => typeof(AssemblyProperties).Assembly;
		
		private static string GetAttributeValue<T>(string cachedName, Func<T, string> member)
		{
			return Values.GetOrAdd(cachedName, key =>
			{
				var attribute = GetAttribute<T>();
				return attribute == null ? null : member(attribute);
			});
		}
	}
}
