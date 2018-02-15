using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ForwardPhishingToAbuseAddin;
using ForwardPhishingToAbuseAddin.Config;

namespace DeploymentTools
{
	class Program
	{
		static int Main(string[] args)
		{
			if (!args.Any())
			{
				Console.Error.WriteLine($"Specify a tool name. Possible tool names: {string.Join(",", GetPossibleOptions())}");
				return 1;
			}

			var getOption = GetPossibleOptions().FirstOrDefault(o => string.Equals(args[0], o, StringComparison.OrdinalIgnoreCase));
			if (getOption == null)
			{
				Console.Error.WriteLine($"Specify a valid tool name. Possible tool names: {string.Join(", ", GetPossibleOptions())}");
				return 2;
			}

			try
			{
				typeof(Tools).GetMethod(getOption)?.Invoke(null, null);
			}
			catch (Exception e)
			{
				Console.Error.WriteLine(e);
				return 3;
			}
			
			return 0;
		}

		private static IEnumerable<string> GetPossibleOptions()
		{
			return typeof(Tools).GetMethods(BindingFlags.Public | BindingFlags.Static).Where(m => m.ReturnType == typeof(void)).Select(x => x.Name);
		}
	}
}
