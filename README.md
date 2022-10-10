Requirements to run the EWS-SampleOAuth script

	- MSAL (Microsoft Authentication Library) module

	- EWS Managed API


Description

	- This script performs a search against a mailbox based on provided subject. 

	- The user will be prompted for credentials to obtain an OAuth token. 

	- The credentials can be either a standard user or a service account with impersonation rights.
           Note: Only searches against well known folder names


Parameters

	- MailboxName: Specifies the mailbox to connect

	- FolderName: Specifies the well known folder name to search

	- TenantId: Specifies the Exchange Online tenant where mailbox resides

	- Subject: Specifies the Subject for the search

	- UseImpersonation: Specifies whether or not EWS impersonation is used

	- EnableLogging: Specifies whether or not to enable EWS trace logging for troubleshooting. Trace log will be written in current directory.


Examples

	.\EWS-SampleOAuth.ps1 -MailboxName jmartin@contoso.com -FolderName Inbox -TenantId contoso.onmicrosoft.com -Subject "Test message" -UseImpersonation:$False
		This will perform a search against the Inbox folderfor the subject "Test message" using the authenticated user account.

	.\EWS-SampleOAuth.ps1 -MailboxName jmartin@contoso.com -FolderName Calendar -TenantId contoso.onmicrosoft.com -Subject "Confidential" -EnableLogging:$True
		This will perform a search against the Calendar folder for the subject "Confidential" and create a trace log for troubleshooting purposes.