using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ForwardPhishingToAbuseAddin.Logging
{
	public static class ErrorCodes
	{
		public static IDictionary<ushort, string> All =>
			typeof(ErrorCodes).GetProperties(BindingFlags.Public | BindingFlags.Static)
				.Where(p => p.PropertyType == typeof(ushort))
				.ToDictionary(p => (ushort) p.GetValue(null, null), p => p.Name);

		public static ushort FailedToDeleteFromExplorerView => 102;
		public static ushort FailedToDeleteFromMailReadView => 103;
		public static ushort FailedToLoadPlugin => 104;

		public static ushort FailedToReadRegistry => 105;
	}
}