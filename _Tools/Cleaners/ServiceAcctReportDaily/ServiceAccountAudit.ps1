# =========================================================================
# TITLE: ServiceAccountAudit_20250211_0800
# AUTHOR: System Administrator
# DATE: February 11, 2025
# VERSION: 1.1
# =========================================================================

# Enable verbose output
$VerbosePreference = "Continue"

# Import required modules
Write-Verbose "Importing Active Directory module..."
Try {
    Import-Module ActiveDirectory -ErrorAction Stop
} Catch {
    Write-Error "Failed to import Active Directory module. Error: $_"
    exit 1
}

# Configuration Variables
$reportPath = $PSScriptRoot
$dateStamp = Get-Date -Format "MMddyyyy_HHmm"
$emailRecipient = "jromaine@davidyurman.com"
$smtpServer = "smtp.office365.com"
$fromAddress = "jromaine@davidyurman.com"
$passwordAgeThreshold = 365
$serviceAccountOU = "OU=ALL DY Service Accounts,DC=Yurman,DC=com"

Write-Verbose "Script initialized with following parameters:"
Write-Verbose "Report Path: $reportPath"
Write-Verbose "Service Account OU: $serviceAccountOU"

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
    
    Try {
        # Define selected properties
        $accounts = Get-ADUser -SearchBase $serviceAccountOU -SearchScope OneLevel -Filter * -Properties * |
            Select-Object @(
                'Name',
                'Created',
                'Description',
                'Enabled',
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
        
    } Catch {
        Write-Error "Failed to retrieve service account data. Error: $_"
        return $null
    }
}

# Generate Reports
Write-Verbose "Generating reports..."
$allAccounts = Get-ServiceAccountData

if ($null -eq $allAccounts) {
    Write-Error "No service account data retrieved. Exiting script."
    exit 1
}

$staleAccounts = $allAccounts | Where-Object { $_.PasswordAge -gt $passwordAgeThreshold }
Write-Verbose "Found $($staleAccounts.Count) accounts with stale passwords"

# Create CSV Files
$allAccountsFile = "ServiceAccounts_$dateStamp.csv"
$staleAccountsFile = "StaleServiceAccounts_$dateStamp.csv"

$allAccountsFile = Get-UniqueFileName -basePath $reportPath -fileName $allAccountsFile
$staleAccountsFile = Get-UniqueFileName -basePath $reportPath -fileName $staleAccountsFile

Write-Verbose "Exporting to CSV files..."
Try {
    $allAccounts | Export-Csv -Path (Join-Path $reportPath $allAccountsFile) -NoTypeInformation
    $staleAccounts | Export-Csv -Path (Join-Path $reportPath $staleAccountsFile) -NoTypeInformation
} Catch {
    Write-Error "Failed to export CSV files. Error: $_"
    exit 1
}

# Generate Email Body
$totalAccounts = $allAccounts.Count
$totalStale = $staleAccounts.Count
$totalEnabled = ($allAccounts | Where-Object { $_.Enabled -eq $true }).Count
$totalDisabled = ($allAccounts | Where-Object { $_.Enabled -eq $false }).Count

$emailBody = @"
===========================================================
           SERVICE ACCOUNT AUDIT REPORT
           Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm")
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
Enabled: $($_.Enabled)
----------------------------------------
"@
}

# Send Email
Write-Verbose "Sending email report..."
Try {
    $emailParams = @{
        From = $fromAddress
        To = $emailRecipient
        Subject = "Service Account Audit Report - $(Get-Date -Format 'yyyy-MM-dd')"
        Body = $emailBody
        SmtpServer = $smtpServer
        Attachments = (Join-Path $reportPath $allAccountsFile), (Join-Path $reportPath $staleAccountsFile)
    }

    Send-MailMessage @emailParams
    Write-Verbose "Email sent successfully"
} Catch {
    Write-Error "Failed to send email. Error: $_"
    exit 1
}

Write-Verbose "Script completed successfully"