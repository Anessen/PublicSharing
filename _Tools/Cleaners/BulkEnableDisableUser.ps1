# Hybrid User Account Disablement Script
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath
)

# Clear screen and set error preference
Clear-Host
$ErrorActionPreference = 'Stop'

# Function to validate and load modules
function Initialize-RequiredModules {
    $modules = @(
        'ActiveDirectory',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )
    
    foreach ($module in $modules) {
        try {
            if (!(Get-Module -Name $module)) {
                Import-Module $module -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Required module '$module' not found or failed to load. Error: $_"
            exit 1
        }
    }
}

# Function to get user details from both AD and Azure
function Get-UserStatus {
    param([string]$UserPrincipalName)
    
    $userStatus = [PSCustomObject]@{
        Username = $UserPrincipalName
        ExistsInAD = $false
        ExistsInCloud = $false
        ADEnabled = $null
        CloudEnabled = $null
        LastError = $null
    }
    
    # Check AD Status
    try {
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -Properties Enabled
        if ($adUser) {
            $userStatus.ExistsInAD = $true
            $userStatus.ADEnabled = $adUser.Enabled
        }
    }
    catch {
        $userStatus.LastError = "AD Error: $_"
    }
    
    # Check Cloud Status
    try {
        $mgUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -Property AccountEnabled
        if ($mgUser) {
            $userStatus.ExistsInCloud = $true
            $userStatus.CloudEnabled = $mgUser.AccountEnabled
        }
    }
    catch {
        if ($userStatus.LastError) {
            $userStatus.LastError += " | Cloud Error: $_"
        } else {
            $userStatus.LastError = "Cloud Error: $_"
        }
    }
    
    return $userStatus
}

# Main execution block
try {
    # Initialize modules
    Initialize-RequiredModules

    # Get CSV path if not provided
    if (-not $CsvPath) {
        $CsvPath = Read-Host "Don't forget the LOG saves to current Directory!`n`n   (CSV Header = UserPrincipalName)`n`nEnter the full path for the user list CSV"
    }

    # Validate CSV
    if (-not (Test-Path $CsvPath)) {
        throw "Invalid CSV path: $CsvPath"
    }

    # Setup logging
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile = Join-Path $PWD.Path "UserDisableLog_$timestamp.csv"

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All" -NoWelcome

    # Import and validate users
    $users = Import-Csv -Path $CsvPath
    if (-not ($users | Get-Member -Name "UserPrincipalName")) {
        throw "CSV must contain 'UserPrincipalName' column"
    }

    # Collect current status
    Write-Host "`nCollecting current status for all users..." -ForegroundColor Cyan
    $userStatuses = @()
    foreach ($user in $users) {
        $status = Get-UserStatus -UserPrincipalName $user.UserPrincipalName
        $userStatuses += $status
    }

    # Display current status
    Write-Host "`nCurrent User Statuses:" -ForegroundColor Green
    $userStatuses | Format-Table -AutoSize

    # Confirmation
    $proceed = Read-Host "`nDo you want to proceed with disabling these accounts? (Y/N)"
    if ($proceed -ne 'Y') {
        throw "Operation cancelled by user"
    }

    # Disable accounts
    $results = @()
    foreach ($status in $userStatuses) {
        $result = $status.PsObject.Copy()
        $result | Add-Member -NotePropertyName TimeStamp -NotePropertyValue (Get-Date)
        $result | Add-Member -NotePropertyName ActionTaken -NotePropertyValue "None"
        
        try {
            # Disable in AD if exists
            if ($status.ExistsInAD) {
                Disable-ADAccount -Identity ($status.Username.Split('@')[0])
                $result.ActionTaken = "AD"
            }
            
            # Disable in Cloud if exists
            if ($status.ExistsInCloud) {
                Update-MgUser -UserId $status.Username -AccountEnabled:$false
                $result.ActionTaken += if($result.ActionTaken -eq "AD"){"&Cloud"}else{"Cloud"}
            }

            # Get updated status
            Start-Sleep -Seconds 2  # Allow time for changes to propagate
            $newStatus = Get-UserStatus -UserPrincipalName $status.Username
            $result | Add-Member -NotePropertyName NewADEnabled -NotePropertyValue $newStatus.ADEnabled
            $result | Add-Member -NotePropertyName NewCloudEnabled -NotePropertyValue $newStatus.CloudEnabled
        }
        catch {
            $result.LastError = $_.Exception.Message
            $result | Add-Member -NotePropertyName NewADEnabled -NotePropertyValue "Error"
            $result | Add-Member -NotePropertyName NewCloudEnabled -NotePropertyValue "Error"
        }
        
        $results += $result
    }

    # Export and display results
    $results | Export-Csv -Path $logFile -NoTypeInformation
    Write-Host "`nFinal Results:" -ForegroundColor Green
    $results | Format-Table Username, ExistsInAD, ExistsInCloud, ADEnabled, CloudEnabled, ActionTaken, NewADEnabled, NewCloudEnabled -AutoSize
    Write-Host "`nLog file saved to: $logFile" -ForegroundColor Cyan
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}