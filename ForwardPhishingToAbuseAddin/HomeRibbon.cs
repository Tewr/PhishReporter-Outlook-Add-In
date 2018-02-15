using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ForwardPhishingToAbuseAddin
{
	public class HomeRibbon : BaseReportingRibbon
	{
		public HomeRibbon() : base(OutlookTargets.TabMail, OutlookTargets.RibbonTypeExplorer, "GroupFind")
		{
		}
	}
}
