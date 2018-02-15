using ForwardPhishingToAbuseAddin.Services;
using Microsoft.Office.Core;

namespace ForwardPhishingToAbuseAddin
{

	public static class OutlookTargets
	{
		public static string TabMail = "TabMail";
		public static string TabReadMessage = "TabReadMessage";
		public static string RibbonTypeExplorer = "Microsoft.Outlook.Explorer";
		public static string RibbonTypeMailRead = "Microsoft.Outlook.Mail.Read";
	}

	partial class BaseReportingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		private string _officeId = OutlookTargets.TabMail;
		private string _ribbonType = OutlookTargets.RibbonTypeExplorer;
		private string _positionAfterOfficeId = "GroupFind";

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;


		public BaseReportingRibbon() : this(OutlookTargets.TabMail, OutlookTargets.RibbonTypeExplorer, "GroupFind")
		{
		}

		public BaseReportingRibbon(string officeId, string ribbonType, string positionAfterOfficeId)
			: base(Globals.Factory.GetRibbonFactory())
		{
			_officeId = officeId;
			_ribbonType = ribbonType;
			_positionAfterOfficeId = positionAfterOfficeId;
			InitializeComponent();
			this.genericTab.ControlId.OfficeId = _officeId;
			this.genericTab.Position = this.Factory.RibbonPosition.AfterOfficeId(_positionAfterOfficeId);
			this.reportSecurityIssuesRibbonGroup.Position = this.Factory.RibbonPosition.AfterOfficeId(_positionAfterOfficeId);
			this.RibbonType = _ribbonType;
			this.Name = GetType().Name;
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.genericTab = this.Factory.CreateRibbonTab();
			this.reportSecurityIssuesRibbonGroup = this.Factory.CreateRibbonGroup();
			this.phishingReportingButton = this.Factory.CreateRibbonButton();
			this.genericTab.SuspendLayout();
			this.SuspendLayout();
			// 
			// phishReporterTab
			// 
			this.genericTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			
			this.genericTab.Groups.Add(this.reportSecurityIssuesRibbonGroup);
			this.genericTab.Label = "TabMail";
			this.genericTab.Name = "genericTab";
			
			// 
			// reportSecurityIssuesRibbonGroup
			// 
			this.reportSecurityIssuesRibbonGroup.Items.Add(this.phishingReportingButton);
			this.reportSecurityIssuesRibbonGroup.Label = ServiceProvider.Config.RibbonGroupName;
			this.reportSecurityIssuesRibbonGroup.Name = "reportSecurityIssuesRibbonGroup";
			
			// 
			// phishingReportingButton
			// 
			this.phishingReportingButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.phishingReportingButton.Label = ServiceProvider.Config.ReportingButtonLabel;
			this.phishingReportingButton.OfficeImageId = "TrustCenter";
			this.phishingReportingButton.Name = "phishingReportingButton";
			this.phishingReportingButton.ScreenTip = ServiceProvider.Config.ReportingButtonScreenTip;
			this.phishingReportingButton.ShowImage = true;
			this.phishingReportingButton.SuperTip = ServiceProvider.Config.ReportingButtonSuperTip;
			this.phishingReportingButton.Click += PhishingReportingButton_Click;
			// 
			// HomeRibbon
			// 

			this.Tabs.Add(this.genericTab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
			this.genericTab.ResumeLayout(false);
			this.genericTab.PerformLayout();
			this.ResumeLayout(false);
		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab genericTab;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton phishingReportingButton;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup reportSecurityIssuesRibbonGroup;
	}

	partial class ThisRibbonCollection
	{
		internal HomeRibbon HomeRibbon
		{
			get { return this.GetRibbon<HomeRibbon>(); }
		}

		internal MailMessageRibbon MailMessageRibbon
		{
			get { return this.GetRibbon<MailMessageRibbon>(); }
		}
	}
}
