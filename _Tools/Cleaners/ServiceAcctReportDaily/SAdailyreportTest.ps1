# =========================================================================
# TITLE: ServiceAccountAudit_20250211_0800
# AUTHOR: System Administrator
# DATE: February 11, 2025
# VERSION: 1.1
# DESCRIPTION: Service Account Audit with Modern Authentication Email
# =========================================================================

# Set up logging first
$scriptPath = $PSScriptRoot
$logPath = Join-Path $scriptPath "Logs"
$logFile = Join-Path $logPath "ServiceAccountAudit_$(Get-Date -Format 'yyyyMMdd').log"

# Create log directory if it doesn't exist
if (-not (Test-Path $logPath)) {
    Write-Host "Creating log directory at: $logPath"
    New-Item -ItemType Directory -Path $logPath | Out-Null
}

# Start transcript logging
Start-Transcript -Path $logFile -Append

Write-Host "===================================================="
Write-Host "Script execution started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "===================================================="

# Import required modules
Write-Verbose "Importing required modules..."
try {
    Import-Module -SkipEditionCheck ActiveDirectory -ErrorAction Stop
    
    # Install and import Microsoft.Graph modules if not present
    if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Mail
} catch {
    Write-Error "Failed to import required modules. Error: $_"
    Stop-Transcript
    exit 1
}

# Configuration Variables
$emailRecipient = "jromaine@davidyurman.com"
$passwordAgeThreshold = 365
$serviceAccountOU = "OU=ALL DY Service Accounts,DC=Yurman,DC=com"

# Function to generate unique filename
function Get-UniqueFileName {
    param (
        [string]$basePath,
        [string]$fileName
    )
    
    $counter = 0
    $newFileName = $fileName
    
    while (Test-Path (Join-Path $basePath $newFileName)) {
        $counter++
        $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + 
                      $counter + [System.IO.Path]::GetExtension($fileName)
    }
    
    Write-Verbose "Generated unique filename: $newFileName"
    return $newFileName
}

# Function to get service account data
function Get-ServiceAccountData {
    Write-Verbose "Retrieving service account data..."
    
    try {
        $accounts = Get-ADUser -SearchBase $serviceAccountOU -SearchScope OneLevel -Filter * -Properties @(
            'Created',
            'Description',
            'DisplayName',
            'EmailAddress',
            'AccountExpirationDate',
            'GivenName',
            'LastLogonDate',
            'Surname',
            'Manager',
            'PasswordLastSet',
            'PasswordExpired',
            'PasswordNeverExpires',
            'PasswordNotRequired',
            'ProtectedFromAccidentalDeletion',
            'SID',
            'CannotChangePassword',
            'UserPrincipalName'
        ) | Select-Object @(
            'Name',
            @{Name='CreationDate';Expression={$_.Created}},
            'Description',
            @{Name='Disabled';Expression={-not $_.Enabled}},
            'DisplayName',
            'DistinguishedName',
            'EmailAddress',
            'AccountExpirationDate',
            'GivenName',
            'LastLogonDate',
            'Surname',
            'Manager',
            'PasswordExpired',
            'PasswordNeverExpires',
            'PasswordNotRequired',
            'ProtectedFromAccidentalDeletion',
            'SID',
            'CannotChangePassword',
            'SamAccountName',
            'UserPrincipalName',
            @{Name='PasswordAge';Expression={
                if($_.PasswordLastSet) {
                    (New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).Days
                } else {
                    0
                }
            }},
            @{Name='ParentContainer';Expression={
                ($_.DistinguishedName -split ',')[1..99] -join ','
            }},
            @{Name='ParentContainerReversed';Expression={
                ($_.DistinguishedName -split ',')[1..99] | Select-Object -First 1
            }},
            @{Name='MustChangePasswordAtNextLogon';Expression={
                $_.PasswordExpired
            }},
            @{Name='PasswordLastChanged';Expression={
                $_.PasswordLastSet
            }},
            @{Name='PasswordExpirationDate';Expression={
                if($_.PasswordNeverExpires) {
                    "Never"
                } else {
                    $_.PasswordLastSet.AddDays($passwordAgeThreshold)
                }
            }}
        )

        Write-Verbose "Retrieved $($accounts.Count) service accounts"
        return $accounts
        
    } catch {
        Write-Error "Failed to retrieve service account data. Error: $_"
        return $null
    }
}

# Generate Reports
Write-Verbose "Generating reports..."
$allAccounts = Get-ServiceAccountData

if ($null -eq $allAccounts) {
    Write-Error "No service account data retrieved. Exiting script."
    Stop-Transcript
    exit 1
}

$staleAccounts = $allAccounts | Where-Object { $_.PasswordAge -gt $passwordAgeThreshold }
Write-Verbose "Found $($staleAccounts.Count) accounts with stale passwords"

# Create CSV Files
$dateStamp = Get-Date -Format "MMddyyyy_HHmm"
$allAccountsFile = "ServiceAccounts_$dateStamp.csv"
$staleAccountsFile = "StaleServiceAccounts_$dateStamp.csv"

$allAccountsFile = Get-UniqueFileName -basePath $scriptPath -fileName $allAccountsFile
$staleAccountsFile = Get-UniqueFileName -basePath $scriptPath -fileName $staleAccountsFile

Write-Verbose "Exporting to CSV files..."
try {
    $allAccounts | Export-Csv -Path (Join-Path $scriptPath $allAccountsFile) -NoTypeInformation
    $staleAccounts | Export-Csv -Path (Join-Path $scriptPath $staleAccountsFile) -NoTypeInformation
} catch {
    Write-Error "Failed to export CSV files. Error: $_"
    Stop-Transcript
    exit 1
}

# Generate Email Content
$totalAccounts = $allAccounts.Count
$totalStale = $staleAccounts.Count
$totalEnabled = ($allAccounts | Where-Object { -not $_.Disabled }).Count
$totalDisabled = ($allAccounts | Where-Object { $_.Disabled }).Count

$emailBody = @"
===========================================================
SERVICE ACCOUNT AUDIT REPORT
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm")
$scriptPath
===========================================================

SUMMARY:
--------
Total Service Accounts: $totalAccounts
Accounts with Stale Passwords (>365 days): $totalStale
Enabled Accounts: $totalEnabled
Disabled Accounts: $totalDisabled

STALE PASSWORD ACCOUNTS:
-----------------------
"@

$staleAccounts | ForEach-Object {
    $emailBody += @"

UserPrincipalName: $($_.UserPrincipalName)
Password Age: $($_.PasswordAge) days
Manager: $($_.Manager)
Enabled: $($_.Disabled -eq $false)
----------------------------------------
"@
}

# Send Email using Microsoft Graph
Write-Verbose "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph -Scopes "Mail.Send", "Mail.ReadWrite" -NoWelcome
    
    # Handle file attachments
    $attachments = @()
    
    # Function to safely read and encode file
    function Convert-FileToAttachment {
        param (
            [string]$FilePath,
            [string]$FileName
        )
        
        try {
            $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
            $base64Content = [Convert]::ToBase64String($fileContent)
            
            return @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name = $FileName
                ContentBytes = $base64Content
            }
        } catch {
            Write-Error "Failed to process attachment $FileName. Error: $_"
            return $null
        }
    }
    
    # Process each attachment
    $allAccountsPath = Join-Path $scriptPath $allAccountsFile
    $staleAccountsPath = Join-Path $scriptPath $staleAccountsFile
    
    $allAccountsAttachment = Convert-FileToAttachment -FilePath $allAccountsPath -FileName $allAccountsFile
    $staleAccountsAttachment = Convert-FileToAttachment -FilePath $staleAccountsPath -FileName $staleAccountsFile
    
    if ($allAccountsAttachment -and $staleAccountsAttachment) {
        $attachments = @($allAccountsAttachment, $staleAccountsAttachment)
    } else {
        throw "Failed to prepare one or more attachments"
    }
    
    # Create email message with attachments
    $messageParams = @{
        Message = @{
            Subject = "Service Account Audit Report - $(Get-Date -Format 'yyyy-MM-dd')"
            Body = @{
                ContentType = "Text"
                Content = $emailBody
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $emailRecipient
                    }
                }
            )
            Attachments = $attachments
        }
        SaveToSentItems = $true
    }
    
    # Send the message
    Write-Verbose "Sending email report..."
    Send-MgUserMail -UserId $emailRecipient -BodyParameter $messageParams
    Write-Verbose "Email sent successfully"
    
} catch {
    Write-Error "Failed to send email. Error: $_"
} finally {
    Disconnect-MgGraph
}

Write-Host "===================================================="
Write-Host "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "===================================================="
Stop-Transcript