using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ForwardPhishingToAbuseAddin.Properties;
using ForwardPhishingToAbuseAddin.Services;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace ForwardPhishingToAbuseAddin
{
	public partial class BaseReportingRibbon
	{
		private string _confirmationMessage;
		private string _reportEmailBody;

		private static IPhisingReporterConfig Config => ServiceProvider.Config;
		private static IApplicationInfo AppInfo => ServiceProvider.AppInfo;

		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{
			try
			{
				_confirmationMessage = Config.ReportingConfirmationMessage.Replace("$SecurityTeamEmail$", Config.SecurityTeamEmail);
				_reportEmailBody = Config.SecurityTeamEmailBody.Replace("$PluginName$", AppInfo.ApplicationProduct);
				if (string.IsNullOrWhiteSpace(Config.RunbookUrl))
				{
					_reportEmailBody = _reportEmailBody + ".";
				}
				else
				{
					_reportEmailBody = _reportEmailBody + Config.ProcessAddendumIfRunbookUrlConfigured.Replace("$RunbookUrl$", Config.RunbookUrl);
				}
			}
			catch (System.Exception ex)
			{
				ServiceProvider.Log.LogError(() => $"Error when loading Ribbon {GetType().Name}", ex);
				throw;
			}
			
		}

		private void PhishingReportingButton_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				ReportCurrentSelection();
			}
			catch (System.Exception ex)
			{
				ServiceProvider.Log.LogError(() => $"Error when calling {nameof(PhishingReportingButton_Click)}",ex);
				throw;
			}
		}

		private void ReportCurrentSelection()
		{
			var exp = Globals.ThisAddIn.Application.ActiveExplorer();
			if (exp.Selection.Count > 0)
			{
				var response = MessageBox.Show(_confirmationMessage, Config.ReportingConfirmationTitle, MessageBoxButtons.YesNo);
				if (response == DialogResult.Yes)
				{
					var reportEmail = (MailItem) Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
					foreach (MailItem phishEmail in exp.Selection)
					{
						reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);
					}

					reportEmail.Subject = Config.ReportingEmailSubject;
					reportEmail.To = Config.SecurityTeamEmail;
					reportEmail.Body = _reportEmailBody;

					reportEmail.Send();
					foreach (MailItem phishEmail in exp.Selection)
					{
						phishEmail.Delete();
					}
				}
			}
			else
			{
				MessageBox.Show(Config.NoEmailSelectedMessage, Config.NoEmailSelectedTitle, MessageBoxButtons.OK);
			}
		}
	}
}
