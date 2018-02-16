using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ForwardPhishingToAbuseAddin.Logging;
using ForwardPhishingToAbuseAddin.Services;
using ForwardPhishingToAbuseAddin.Tools;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Exception = System.Exception;

namespace ForwardPhishingToAbuseAddin
{
	public partial class BaseReportingRibbon
	{
		private string _confirmationMessage;
		private string _reportEmailBody;
		private static IApplicationInfo AppInfo => ServiceProvider.AppInfo;

		private static IPhisingReporterConfig Config => ServiceProvider.Config;

		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{
			try
			{
				_confirmationMessage = Config.ReportingConfirmationMessage.Replace("$SecurityTeamEmail$", Config.SecurityTeamEmail);
				_reportEmailBody = Config.SecurityTeamEmailBody.Replace("$PluginName$", AppInfo.ApplicationProduct);
				if (string.IsNullOrWhiteSpace(Config.RunbookUrl))
					_reportEmailBody = _reportEmailBody + ".";
				else
					_reportEmailBody = _reportEmailBody +
					                   Config.ProcessAddendumIfRunbookUrlConfigured.Replace("$RunbookUrl$", Config.RunbookUrl);
			}
			catch (Exception ex)
			{
				ServiceProvider.Log.LogError(() => $"Error when loading Ribbon {GetType().Name}", ex,
					ErrorCodes.FailedToLoadPlugin);
				throw;
			}
		}

		private void PhishingReportingButton_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				ReportCurrentSelection();
			}
			catch (Exception ex)
			{
				ServiceProvider.Log.LogError(() => $"Error when calling {nameof(PhishingReportingButton_Click)}", ex,
					FailClickErrorCode());
				throw;
			}
		}

		protected virtual ushort FailClickErrorCode()
		{
			return 101;
		}

		protected virtual List<MailItem> GetSelectedMailItems()
		{
			return Globals.ThisAddIn.Application.ActiveExplorer().Selection.OfType<MailItem>().ToList();
		}

		private void ReportCurrentSelection()
		{
			var selectedMailItems = GetSelectedMailItems();
			if (selectedMailItems.Any())
			{
				var response = MessageBox.Show(_confirmationMessage, Config.ReportingConfirmationTitle, MessageBoxButtons.YesNo,
					MessageBoxIcon.Question);
				if (response == DialogResult.Yes)
					SendAndDispose(selectedMailItems);
			}
			else
			{
				MessageBox.Show(Config.NoEmailSelectedMessage, Config.NoEmailSelectedTitle, MessageBoxButtons.OK,
					MessageBoxIcon.Information);
			}
		}

		private void SendAndDispose(List<MailItem> phishEmails)
		{
			MailItem reportEmail = null;
			try
			{
				reportEmail = (MailItem) Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
				foreach (var phishEmail in phishEmails)
					reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);

				reportEmail.Subject = Config.ReportingEmailSubject.Replace("$PluginName$", AppInfo.ApplicationProduct);
				reportEmail.To = Config.SecurityTeamEmail;
				reportEmail.Body = _reportEmailBody;

				reportEmail.Send();

				foreach (var phishEmail in phishEmails)
					phishEmail.Delete();
			}
			finally
			{
				reportEmail.Dispose();
				foreach (var phishEmail in phishEmails)
					phishEmail.Dispose();
			}
		}
	}
}
