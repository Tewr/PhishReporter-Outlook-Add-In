using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ForwardPhishingToAbuseAddin
{
	public class MailMessageRibbon: BaseReportingRibbon
	{
		public MailMessageRibbon() : base(OutlookTargets.TabReadMessage, OutlookTargets.RibbonTypeMailRead, "GroupZoom")
		{}
	}
}
