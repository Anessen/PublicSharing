# Save this as Test-ServiceAccountManager.ps1

#Requires -Version 5.1
#Requires -Modules ActiveDirectory, Microsoft.Graph.Users, Microsoft.Graph.Mail

# First, let's check PowerShell version and module availability
$requiredModules = @('ActiveDirectory', 'Microsoft.Graph.Users', 'Microsoft.Graph.Mail')
$missingModules = @()

Write-Host "Checking PowerShell version and required modules..."
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host "`nMissing required modules: $($missingModules -join ', ')" -ForegroundColor Yellow
    Write-Host "`nAttempting to install missing modules..."
    
    foreach ($module in $missingModules) {
        try {
            Write-Host "Installing $module..."
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
            Import-Module $module -Force
            Write-Host "Successfully installed and imported $module" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install $module. Error: $_"
            exit 1
        }
    }
}

# Try to import modules with error handling
foreach ($module in $requiredModules) {
    try {
        Import-Module $module -Force -ErrorAction Stop
        Write-Host "Successfully imported $module" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to import $module. Error: $_"
        Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
        Write-Host "1. Run PowerShell as Administrator"
        Write-Host "2. Try manually installing the module:"
        Write-Host "   Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser"
        Write-Host "3. If the issue persists, try removing and reinstalling the module:"
        Write-Host "   Uninstall-Module -Name $module -AllVersions"
        Write-Host "   Install-Module -Name $module -Force"
        exit 1
    }
}

# Try to connect to Microsoft Graph with error handling
try {
    Write-Host "`nConnecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Mail.Send", "Mail.ReadWrite" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Error: $_"
    Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
    Write-Host "1. Check your internet connection"
    Write-Host "2. Verify your credentials"
    Write-Host "3. Ensure you have the necessary permissions"
    exit 1
}

# Function to test a specific manager
function Test-SingleManager {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ManagerEmail,
        
        [Parameter(Mandatory=$true)]
        [string]$YourEmail
    )

    Write-Host "`nStarting test for manager: $ManagerEmail" -ForegroundColor Cyan

    # Get manager's DN from email
    try {
        $managerDN = (Get-ADUser -Filter "mail -eq '$ManagerEmail'" -ErrorAction Stop).DistinguishedName
        if (-not $managerDN) {
            Write-Error "Could not find manager with email: $ManagerEmail"
            return
        }
        Write-Host "Found manager DN: $managerDN" -ForegroundColor Green
    }
    catch {
        Write-Error "Error finding manager: $_"
        return
    }

    # Get all service accounts for this manager
    $serviceAccountOU = "OU=ALL DY Service Accounts,DC=Yurman,DC=com"
    Write-Host "Searching for service accounts in: $serviceAccountOU"
    
    try {
        $accounts = Get-ADUser -SearchBase $serviceAccountOU -Filter * -Properties Manager, 
            PasswordLastSet, Description, mail, userPrincipalName, Enabled -ErrorAction Stop |
            Where-Object { $_.Manager -eq $managerDN }

        if (-not $accounts) {
            Write-Warning "No service accounts found for manager: $ManagerEmail"
            return
        }
        
        Write-Host "Found $($accounts.Count) service accounts" -ForegroundColor Green
    }
    catch {
        Write-Error "Error retrieving service accounts: $_"
        return
    }

    # Process accounts
    $processedAccounts = $accounts | Select-Object @{
        Name = 'UserPrincipalName'
        Expression = { $_.UserPrincipalName }
    },
    @{
        Name = 'PasswordAge'
        Expression = { 
            if ($_.PasswordLastSet) {
                [math]::Round((New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).TotalDays)
            } else { "Never" }
        }
    },
    @{
        Name = 'PasswordLastChanged'
        Expression = { $_.PasswordLastSet }
    },
    @{
        Name = 'Description'
        Expression = { $_.Description }
    },
    @{
        Name = 'Disabled'
        Expression = { -not $_.Enabled }
    }

    # Create email body
    $accountTable = ""
    foreach ($account in $processedAccounts) {
        $accountTable += @"
UserPrincipalName: $($account.UserPrincipalName)
Password Age: $($account.PasswordAge) days
Last Password Change: $($account.PasswordLastChanged)
Account Status: $($account.Disabled -eq $false ? 'Enabled' : 'Disabled')
Description: $($account.Description)
----------------------------------------

"@
    }

    $emailBody = @"
===========================================================
SERVICE ACCOUNT PASSWORD ALERT - TEST EMAIL
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm")
===========================================================

This is a test email for manager: $ManagerEmail

The following service accounts are under this manager's supervision:

$accountTable

This is a test message. In production, this would be sent to the manager.
"@

    # Send test email
    Write-Host "Preparing to send test email to: $YourEmail"
    
    $messageParams = @{
        Message = @{
            Subject = "TEST - Service Account Manager Notification - $(Get-Date -Format 'yyyy-MM-dd')"
            Body = @{
                ContentType = "Text"
                Content = $emailBody
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $YourEmail
                    }
                }
            )
        }
        SaveToSentItems = $true
    }

    try {
        Send-MgUserMail -UserId $YourEmail -BodyParameter $messageParams
        Write-Host "Test email sent successfully to: $YourEmail" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to send test email: $_"
        Write-Host "Email parameters for troubleshooting:" -ForegroundColor Yellow
        $messageParams | ConvertTo-Json -Depth 3
    }
}

Write-Host "`nScript loaded successfully. You can now run:" -ForegroundColor Green
Write-Host 'Test-SingleManager -ManagerEmail "manager@davidyurman.com" -YourEmail "your.email@davidyurman.com"' -ForegroundColor Cyan

# Note: Don't disconnect from Graph here - let the user run their tests first