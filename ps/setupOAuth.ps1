param (
    [string]$appName,
    [string]$mailboxName,
    [string]$certFile
)

function Show-Usage {
    Write-Host "Usage: setupOAuth2.ps1 -appName <SMTP application name> -mailboxName <shared mailbox name> [-certFile <complete path to certificate file>] [-groupName <mail-enabled security group name>]"
    Write-Host "Example: SetupEntra.ps1 -appName MySMTPApp -mailboxName shared@mail.com -certFile 'C:\path\to\cert.cer' -groupName MyMailEnabledGroup"
    exit 1
}

# Manual check before prompting
if (-not $PSBoundParameters.ContainsKey('appName') -or -not $PSBoundParameters.ContainsKey('mailboxName')) {
    Show-Usage
}

# -----------------------------------------------------------------------
# Get user input for the SMTP application name and mailbox
# -----------------------------------------------------------------------

if ('' -eq $certFile) {
    Write-Output "No certificate file provided. A client id and secret will be created instead."
} else {
    $certFileExists = Test-Path -Path $certFile
    if ($certFileExists -eq '')
    {
        Write-Output "File $($certFile) does not exist"
        exit 1
    } else
    {
        Write-Output "File $($certFile) found. "
    }
}

# -----------------------------------------------------------------------
# Check for required modules and install if not present 
# -----------------------------------------------------------------------

if (Get-Module -Name Microsoft.Entra -ListAvailable) {
    # Get the installed version of the Microsoft Entra PowerShell module
    $moduleVersion = (Get-Module -Name Microsoft.Entra -ListAvailable).Version
    Write-Output "Microsoft Entra PowerShell module version $moduleVersion is installed."
} else {
    Write-Output "Installing Microsoft Entra PowerShell module..."
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    Install-Module -Name Microsoft.Entra -Repository PSGallery -Scope CurrentUser -Force -AllowClobber
}

if (Get-Module -Name ExchangeOnlineManagement -ListAvailable) {
    # Get the installed version of the Exchange Online PowerShell module
    $moduleVersion = (Get-Module -Name ExchangeOnlineManagement -ListAvailable).Version
    Write-Output "Exchange Online PowerShell module version $moduleVersion is installed."
} else {
    Write-Output "Installing ExchangeOnlineManagement PowerShell module..."
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# -----------------------------------------------------------------------
# Step 1: Register the app representing the SAP system in Entra ID tenant 
# -----------------------------------------------------------------------

# Connect to Entra with required scopes
Write-Output "Connecting to Entra ID tenant in a browser window..."
# Note: The -ContextScope 'Process' parameter is used to ensure that the connection is scoped to the current PowerShell process.
Connect-Entra -ContextScope 'Process' -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'GroupMember.Read.All' -NoWelcome
$tenantId=(Get-EntraContext).TenantId
$userName=(Get-EntraContext).Account
Write-Output "Connected to Entra ID tenant $($tenantId) as $($userName)."

# Check if the mail-enabled security group exists if $groupName is provided
if ($groupName) {
    $group = Get-EntraGroup -Filter "DisplayName eq '$($groupName)'" -ErrorAction SilentlyContinue
    if (-not $group) {
        Write-Output "Mail-enabled security group $($groupName) does not exist. Please create the mail-enabled security group before running this script."
        exit 1
    } else {
        Write-Output "Mail-enabled security group $($groupName) found with ID $($group.Id)."
    }
}

# Check if the application already exists
$app = Get-EntraApplication -Filter "DisplayName eq '$($appName)'" -ErrorAction SilentlyContinue
if ($app) {
    Write-Output "Application $($app.DisplayName) already exists with ID $($app.AppId)."
} else {
    Write-Output "Registering application $($appName)..."
    $app = New-EntraApplication -DisplayName $appName
    Write-Output "Application $($app.DisplayName) registered with ID $($app.AppId)."
}

# Set the SMTP.SendAsApp API application permission for the application
# Note: 00000002-0000-0ff1-ce00-000000000000 is the App ID for OFfice 365 Exchange Online (https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/governance/verify-first-party-apps-sign-in)
$applicationPermission = 'SMTP.SendAsApp'
$exchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'
$exchangeOnlineServicePrincipal = Get-EntraServicePrincipal -Filter "AppId eq '$exchangeOnlineAppId'"

# Create resource access object
$resourceAccess = New-Object Microsoft.Open.MSGraph.Model.ResourceAccess
$resourceAccess.Id = ((Get-EntraServicePrincipal -ServicePrincipalId $exchangeOnlineServicePrincipal.ObjectId).AppRoles | Where-Object { $_.Value -eq $applicationPermission}).Id
$resourceAccess.Type = 'Role'

# Create required resource access object
$requiredResourceAccess = New-Object Microsoft.Open.MSGraph.Model.RequiredResourceAccess
$requiredResourceAccess.ResourceAppId = $exchangeOnlineAppId
$requiredResourceAccess.ResourceAccess = $resourceAccess

# Set application required resource access
Set-EntraApplication -ApplicationId $app.Id -RequiredResourceAccess $requiredResourceAccess
Write-Output "$($applicationPermission) permission set for application $($app.DisplayName) with ID $($app.AppId)."

# Check if a service principal already exists for the application
$appServicePrincipal = Get-EntraServicePrincipal -Filter "AppId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
if ($appServicePrincipal) {
    Write-Output "Entra service principal $($appServicePrincipal.DisplayName) with ID $($appServicePrincipal.Id) for Entra app $($app.DisplayName) already exists."
} else {
    # Create a service principal for the application
    Write-Output "Creating Entra service principal for Entra application $($app.DisplayName)..."
    $appServicePrincipal = New-EntraServicePrincipal -AppId $app.AppId
    Write-Output "Entra service principal with ID $($appServicePrincipal.Id) created for Entra application $($app.DisplayName)."
}

# Grant tenant-wide admin consent for the application
$exchangeServicePrincipal = Get-EntraServicePrincipal -Filter "AppId eq '$($exchangeOnlineAppId)'" -ErrorAction SilentlyContinue
# Check for existing admin consent
$roleAssignment = Get-EntraServicePrincipalAppRoleAssignment -ServicePrincipalId $appServicePrincipal.Id
if ($roleAssignment) {
    Write-Output "Admin consent already granted for permission $($applicationPermission) for Entra application $($app.DisplayName)."
} else {
    Write-Output "Granting tenant-wide admin consent for permission $($applicationPermission) for Entra application $($app.DisplayName)..."
    New-EntraServicePrincipalAppRoleAssignment -ServicePrincipalId $appServicePrincipal.Id -PrincipalId $appServicePrincipal.Id -Id $resourceAccess.Id -ResourceId $exchangeServicePrincipal.Id
}

# check for existing key credentials
$keyCredentials = Get-EntraApplicationKeyCredential -ApplicationId $app.ObjectId -ErrorAction SilentlyContinue
if ($keyCredentials) {
    Write-Output "Key credentials already exist for Entra application $($app.DisplayName) with ID $($app.AppId)."
} else {
    Write-Output "No key credentials found for Entra application $($app.DisplayName) with ID $($app.AppId)."
}

# Upload certificate if provided
if ($certFile) {

    $cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile) #create a new certificate object
    $bin = $cer.GetRawCertData()
    
    $keyCred = New-Object Microsoft.Open.MSGraph.Model.KeyCredential
    $keyCred.CustomKeyIdentifier = $base64Thumbprint
    $keyCred.Type = "AsymmetricX509Cert"
    $keyCred.Usage = "Verify"
    $keyCred.Key = $bin
       
    Set-EntraApplication -ApplicationId $app.ObjectId -KeyCredentials $keyCred

    Write-Output "Certificate uploaded for Entra application $($app.DisplayName) with ID $($app.AppId)."
} else {
    $clientSecret = New-EntraApplicationPasswordCredential -ApplicationId $app.ObjectId
    Write-Output "No certificate provided. Client secret generated for Entra application $($app.DisplayName) with client ID $($app.AppId)."
}

# -----------------------------------------------------------------------
# Step 2: Register the application service principal in Exchange Online  
# -----------------------------------------------------------------------

Write-Output "Connecting to Exchange Online in a browser window..."
Connect-ExchangeOnline -Organization $tenantId -ShowBanner:$false

# Allow using SMTP OAuth for the shared mailbox
Write-Output "Mailbox-specific setting"
Set-CASMailbox -Identity $mailboxName -SmtpClientAuthenticationDisabled $false

# Check if the service principal already exists in Exchange Online
if (Get-ServicePrincipal -Identity $appServicePrincipal.ObjectId -ErrorAction SilentlyContinue) {
    Write-Output "Entra service principal $($appServicePrincipal.DisplayName) already registered in Exchange Online with ID $($servicePrincipal.Id)."
} else {
    New-ServicePrincipal -AppId $appServicePrincipal.AppId -ObjectId $appServicePrincipal.ObjectId
    Write-Output "Entra service principal $($appServicePrincipal.ObjectId) registered in Exchange Online."
}

# -----------------------------------------------------------------------
# Step 3: Allow the application service principal to access the mailbox
# -----------------------------------------------------------------------
Write-Output "Granting full access to shared mailbox $($mailboxName) for application service principal $($appServicePrincipal.DisplayName)..."
Add-MailboxPermission -Identity $mailboxName -User $appServicePrincipal.ObjectId -AccessRights FullAccess

# -----------------------------------------------------------------------
# Step 4: Clean up and disconnect from Entra and Exchange Online  
# -----------------------------------------------------------------------

Write-Output "Disconnecting from Entra ID tenant $($tenantId) as $($userName)."
Disconnect-Entra

Write-Output "Disconnecting from Exchange Online as $($userName)."
Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Client ID: $($app.AppId)"
if ($certFile) {
    Write-Output "Certificate: $($certFile)"
} else {
    Write-Output "Client Secret: $($clientSecret.SecretText)"
}
Write-Output "Token endpoint URL: https://login.microsoftonline.com/$($tenantId)/oauth2/v2.0/token"
Write-Output "Tenant ID: $($tenantId)"
Write-Output "Mailbox Name: $($mailboxName)"