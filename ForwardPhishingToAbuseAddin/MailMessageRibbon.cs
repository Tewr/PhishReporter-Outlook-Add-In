using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ForwardPhishingToAbuseAddin.Logging;
using Microsoft.Office.Interop.Outlook;

namespace ForwardPhishingToAbuseAddin
{
	public class MailMessageRibbon: BaseReportingRibbon
	{
		public MailMessageRibbon() : base(OutlookTargets.TabReadMessage, OutlookTargets.RibbonTypeMailRead, "GroupZoom")
		{}

		protected override List<MailItem> GetSelectedMailItems()
		{
			var inspector = Globals.ThisAddIn.Application.ActiveInspector();
			return new[] { inspector.CurrentItem }.OfType<MailItem>().ToList();
		}

		protected override ushort FailClickErrorCode()
		{
			return ErrorCodes.FailedToDeleteFromMailReadView;
		}
	}
}
