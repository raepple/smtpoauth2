Simple Java test client for SMTP-based access to M365 Exchange Online with SASL XOAUTH2 authentication

## Prerequisites
- Use the [PowerShell OAuth2 setup script](../ps/setupOAuth.ps1) in this repository to setup your Microsoft Entra and Exchange Online tenants.
- Ensure that your user's mail settings have [SMTP AUTH enabled](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission#enable-smtp-auth-for-specific-mailboxes). Notice: this requires PowerShell 7.x
- The test client can optionally send a test message from the shared Exchange Online mailbox to a recipient's email address specified by the `r` command line parameter
- [Apache Maven](https://maven.apache.org/) to build the executable JAR file for the test client.

## Known limitations
This test client does not support signed JSON Web Tokens (JWT) as an client authentication grant type. It only supports `client id` and `secret`. You can generate these credentials for the application service principal created by the [PowerShell OAuth2 setup script](../ps/setupOAuth.ps1) in the Entra admin center [as documented here](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-credentials?tabs=client-secret), or by **not** passing a certificate file with the `certFile` command line parameter.

## Build
Run `mvn clean package` in this directory to build the executable JAR file for the test client.

## Usage
Run `java -jar .\target\smtpoauth2test-1.0-SNAPSHOT.jar -c <clientId> -s <clientSecret> -t <Entra tenant id> -m <Exchange Online mailbox> [-r <test email recipient address>]` from this directory.

## Example
`java -jar .\target\smtpoauth2test-1.0-SNAPSHOT.jar -c 12345678-487a-4244-8f0b-bd20d4abcdef -s -lv8Q~GJIKw2zBIV5BsbUu~KMq5DLv55n8127631 -t 87654321-1234-5678-9876-de7dde1930a1 -m mailbox@contoso.onmicrosoft.com -r testrecipient@contoso.com`