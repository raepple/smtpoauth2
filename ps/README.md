PowerShell script to setup a Microsoft Entra tenant and Exchange Online to use OAuth 2.0 authentication. An external SMTP client can use OAuth client id and secret or a signed JSON Web Token (JWT) to authenticate against Exchange Online.

## Prerequisites
- PowerShell version 7 or higher

## Usage
`setupOAuth2.ps1 -appName <SMTP application name> -mailboxName <shared mailbox name> [-certFile <complete path to certificate file>] [-groupName <mail-enabled security group name>]"`

## Example
`setupEntra.ps1 -appName MySMTPApp -mailboxName shared@mail.com -certFile 'C:\path\to\cert.cer' -groupName MyMailEnabledGroup"`
