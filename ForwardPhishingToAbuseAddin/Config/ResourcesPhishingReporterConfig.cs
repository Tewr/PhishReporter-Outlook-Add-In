using ForwardPhishingToAbuseAddin.Properties;
using ForwardPhishingToAbuseAddin.Services;

namespace ForwardPhishingToAbuseAddin.Config
{
	public class ResourcesPhishingReporterConfig : IPhisingReporterConfig
	{
		public string ReportingButtonLabel => DefaultConfigStrings.ReportingButtonLabel;
		public string ReportingButtonScreenTip => DefaultConfigStrings.ReportingButtonScreenTip;
		public string ReportingButtonSuperTip => DefaultConfigStrings.ReportingButtonSuperTip;
		public string RibbonGroupName => DefaultConfigStrings.RibbonGroupName;
		public string SecurityTeamEmail => DefaultConfigStrings.Mail_MailTo;
		public string ReportingEmailSubject => DefaultConfigStrings.ReportingEmailSubject;
		public string ReportingConfirmationMessage => DefaultConfigStrings.Messages_ReportingConfirmationMessage;
		public string ReportingConfirmationTitle => DefaultConfigStrings.Messages_ReportingConfirmationTitle;
		public string SecurityTeamEmailBody => DefaultConfigStrings.Mail_MailBody;
		public string RunbookUrl => string.Empty;
		public string NoEmailSelectedMessage => DefaultConfigStrings.NoEmailSelectedMessage;
		public string NoEmailSelectedTitle => DefaultConfigStrings.Messages_NoEmailSelectedTitle;
		public string ProcessAddendumIfRunbookUrlConfigured => DefaultConfigStrings.ProcessAddendumIfRunbookUrlConfigured;
	}
}