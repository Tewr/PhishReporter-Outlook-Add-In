using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ForwardPhishingToAbuseAddin.Logging;

namespace ForwardPhishingToAbuseAddin
{
	public class HomeRibbon : BaseReportingRibbon
	{
		public HomeRibbon() : base(OutlookTargets.TabMail, OutlookTargets.RibbonTypeExplorer, "GroupFind")
		{
		}

		protected override ushort FailClickErrorCode()
		{
			return ErrorCodes.FailedToDeleteFromExplorerView;
		}
	}
}
