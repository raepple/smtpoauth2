Simple Java test client for SMTP-based access to M365 Exchange Online with SASL XOAUTH2 authentication

## Prerequisites
- Use the [PowerShell script](../ps/setupOAuth.ps1) in this repository to setup your Microsoft Entra and Exchange Online tenants.
- Ensure that your user's mail settings have [SMTP AUTH enabled](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission#enable-smtp-auth-for-specific-mailboxes). Notice: this requires PowerShell 7.x

## Build
`mvn clean package`

## Usage
`java -jar .\target\smtpoauth2test-1.0-SNAPSHOT.jar -c <clientId> -s <clientSecret> -t <AAD tenant ID> -e <email account>`