namespace ForwardPhishingToAbuseAddin.Services
{
	public interface IPhisingReporterConfig
	{
		string NoEmailSelectedMessage { get; }
		string NoEmailSelectedTitle { get; }
		string ProcessAddendumIfRunbookUrlConfigured { get; }
		string ReportingButtonLabel { get; }

		string ReportingButtonScreenTip { get; }
		string ReportingButtonSuperTip { get; }
		string ReportingConfirmationMessage { get; }

		string ReportingConfirmationTitle { get; }
		string ReportingEmailSubject { get; }
		string RibbonGroupName { get; }
		string RunbookUrl { get; }

		string SecurityTeamEmail { get; }
		string SecurityTeamEmailBody { get; }
	}
}