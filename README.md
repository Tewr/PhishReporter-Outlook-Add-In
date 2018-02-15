# PhishReporter-Outlook-Add-In
C# fork of PhishReporter Outlook Add-In utlook Add-In that allows users to report phishing e-mails to a specific e-mail address for further processing/investigation

This simple, yet efficient, Outlook Add-In adds a button to your Outlook Home Ribbon that allows users to simply select/highlight a phishing email and it will forward it to the appropriate mailbox/e-mail address as an attachment for further analysis.  Once the user has verified that they want to send this Phishing email, then the Outlook Add-In removes it from their inbox and places it in their “trash” folder.

There are tons of companies (now) offering this type of Add-In, but PhishReporter Outlook Add-In is completely free and can be completely customized.

## Requirements to Build/Customize the PhishReporter Outlook Add-In:

* PhishReporter Project Files - Clone this repo
* Visual Studio 2015 (Tested and working on Professional, source branches supposively worked on Community)
* Visual Studio Installer Projects Extension: https://marketplace.visualstudio.com/items?itemName=VisualStudioProductTeam.MicrosoftVisualStudio2015InstallerProjects
* Visual Studio Office Developer Tools

## Building
Create or get a new signing certificate from the c# Project properties. The pfx which is referenced in the project is not under source control.

## Configuring
Edit AssemblyInfo.cs with your company name and perhaps a product name to enable registry configuration. The optional registry setting location will be HKLM\Software\company\productname\ . 
To modify a property, add a new string value with the same name as IPhisingReporterConfig. Some of the strings have template variables on the form $name$, this can be kept(or not) in any replacement string.

For the installer, you should also edit Setup.vbproj company name (affects install location).

## Additions from source repo
* Rewritten in in C#
* Strings in resource files and optionally in registry
* Additional button in mailread view


