using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace ForwardPhishingToAbuseAddin.Tools
{
	public static class MailItemExtension
	{
		public static void Dispose(this MailItem source)
		{
			if (source != null)
			{
				Marshal.ReleaseComObject(source);
			}
		}
	}
}
