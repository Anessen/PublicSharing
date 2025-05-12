# Hybrid User Account Disablement Script
[CmdletBinding(SupportsShouldProcess = $true)] # Added SupportsShouldProcess for -WhatIf support on native cmdlets
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [switch]$EnableAccounts, # Added switch to enable instead of disable

    [Parameter(Mandatory = $false)]
    [switch]$TestMode # Added switch for testing without making changes
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
                Write-Verbose "Importing module: $module"
                Import-Module $module -ErrorAction Stop
            } else {
                Write-Verbose "Module already loaded: $module"
            }
        }
        catch {
            Write-Error "Required module '$module' not found or failed to load. Please install it (e.g., Install-Module $module) and ensure you have permissions. Error: $_"
            exit 1
        }
    }
}

# Function to get user details from both AD and Azure
function Get-UserStatus {
    param([string]$UserPrincipalName)

    $userStatus = [PSCustomObject]@{
        Username        = $UserPrincipalName
        ExistsInAD      = $false
        ExistsInCloud   = $false
        ADEnabled       = $null
        CloudEnabled    = $null
        ADDistinguishedName = $null # Store DN for easier updates later
        CloudObjectId   = $null # Store Object ID for easier updates later
        LastError       = $null
    }

    # Check AD Status
    try {
        # Use SamAccountName derived from UPN for potential reliability in some environments
        $samAccountName = $UserPrincipalName.Split('@')[0]
        # Try searching by UPN first, then fall back to SamAccountName if necessary (adjust filter as needed for your environment)
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -Properties Enabled, DistinguishedName -ErrorAction SilentlyContinue
        if (-not $adUser) {
            $adUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -Properties Enabled, DistinguishedName -ErrorAction SilentlyContinue
        }

        if ($adUser) {
            $userStatus.ExistsInAD = $true
            $userStatus.ADEnabled = $adUser.Enabled
            $userStatus.ADDistinguishedName = $adUser.DistinguishedName
            Write-Verbose "Found AD User: $($userStatus.Username) (Enabled: $($userStatus.ADEnabled))"
        } else {
             Write-Verbose "User not found in AD: $UserPrincipalName"
        }
    }
    catch {
        $userStatus.LastError = "AD Error checking '$UserPrincipalName': $_"
        Write-Warning $userStatus.LastError
    }

    # Check Cloud Status
    try {
        # Ensure Graph connection is established before calling this
        # Using -ConsistencyLevel eventual and -CountVariable ensures robust filtering in large tenants
        $mgUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -Property AccountEnabled, Id -ConsistencyLevel eventual -CountVariable userCount -ErrorAction SilentlyContinue
        if ($userCount -gt 0 -and $mgUser) {
             # Handle potential multiple results if filtering wasn't precise, take the first one
             $targetUser = $mgUser | Select-Object -First 1
            $userStatus.ExistsInCloud = $true
            $userStatus.CloudEnabled = $targetUser.AccountEnabled
            $userStatus.CloudObjectId = $targetUser.Id
             Write-Verbose "Found Cloud User: $($userStatus.Username) (Enabled: $($userStatus.CloudEnabled))"
        } else {
             Write-Verbose "User not found in Cloud: $UserPrincipalName"
        }
    }
    catch {
        $errorMsg = "Cloud Error checking '$UserPrincipalName': $_"
        if ($userStatus.LastError) {
            $userStatus.LastError += " | $errorMsg"
        } else {
            $userStatus.LastError = $errorMsg
        }
         Write-Warning $errorMsg
    }

    return $userStatus
}

# --- Main Execution Block ---
try {
    # Initialize modules
    Initialize-RequiredModules

    # Determine Action Verb based on switch
    $actionVerb = if ($EnableAccounts) { "Enable" } else { "Disable" }
    $actionState = if ($EnableAccounts) { $true } else { $false }
    $actionVerbGerund = if ($EnableAccounts) { "Enabling" } else { "Disabling" } # For messages

    # Announce Test Mode if active
    if ($TestMode) {
        Write-Host "`n*************************************" -ForegroundColor Yellow
        Write-Host "*** RUNNING IN TEST MODE ***" -ForegroundColor Yellow
        Write-Host "*** No accounts will actually be $($actionVerb.ToLower())d. ***" -ForegroundColor Yellow
        Write-Host "*************************************`n"
    }

    # Get CSV path if not provided
    if (-not $CsvPath) {
        $CsvPath = Read-Host "Don't forget the LOG saves to current Directory!`n`n  (CSV Header must include 'UserPrincipalName')`n`nEnter the full path for the user list CSV"
    }

    # Validate CSV Path
    if (-not (Test-Path $CsvPath -PathType Leaf)) {
        throw "Invalid CSV file path: $CsvPath"
    }

    # Setup logging
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile = Join-Path $PWD.Path "UserAccount-$($actionVerb)Log_$timestamp.csv"
    Write-Host "Log file will be saved to: $logFile" -ForegroundColor Cyan

    # Connect to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    # Specify required permissions clearly
    $requiredScopes = @("User.ReadWrite.All", "Directory.Read.All")
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # Import and validate users
    Write-Host "Importing users from $CsvPath..." -ForegroundColor Cyan
    try {
        $users = Import-Csv -Path $CsvPath
    }
    catch {
        throw "Failed to import CSV '$CsvPath'. Ensure it's a valid CSV file. Error: $_"
    }

    if (-not $users) {
        throw "CSV file '$CsvPath' is empty or could not be read."
    }
    if (-not ($users[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
        throw "CSV must contain a 'UserPrincipalName' column header."
    }

    # Collect current status
    Write-Host "`nCollecting current status for $($users.Count) users..." -ForegroundColor Cyan
    $userStatuses = @()
    $progressCount = 0
    foreach ($user in $users) {
        $progressCount++
        Write-Progress -Activity "Fetching User Status" -Status "Processing user $progressCount of $($users.Count): $($user.UserPrincipalName)" -PercentComplete (($progressCount / $users.Count) * 100)

        if (-not $user.UserPrincipalName -or $user.UserPrincipalName -notlike '*@*') {
            Write-Warning "Skipping invalid UserPrincipalName format in CSV: '$($user.UserPrincipalName)'"
             $status = [PSCustomObject]@{ Username = $user.UserPrincipalName; LastError = "Invalid UPN Format"; ExistsInAD = $false; ExistsInCloud=$false; ADEnabled=$null;CloudEnabled=$null}
        } else {
            $status = Get-UserStatus -UserPrincipalName $user.UserPrincipalName
        }
        $userStatuses += $status
    }
    Write-Progress -Activity "Fetching User Status" -Completed

    # Display current status
    Write-Host "`nCurrent User Statuses (Before $actionVerbGerund):" -ForegroundColor Green
    $userStatuses | Format-Table Username, ExistsInAD, ADEnabled, ExistsInCloud, CloudEnabled, LastError -AutoSize

    # Confirmation
    $confirmMsg = "`nDo you want to proceed with $($actionVerbGerund.ToLower()) these accounts?"
    if ($TestMode) {
         $confirmMsg += " (TEST MODE - NO CHANGES WILL BE MADE)"
    }
    $proceed = Read-Host "$confirmMsg (Y/N)"
    if ($proceed -ne 'Y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        # Don't throw an error, just exit gracefully
        exit 0
    }

    # --- Perform Enable/Disable Actions ---
    Write-Host "`n$($actionVerbGerund) accounts..." -ForegroundColor Cyan
    $results = @()
    $progressCount = 0
    foreach ($status in $userStatuses) {
        $progressCount++
        Write-Progress -Activity "$actionVerbGerund Accounts" -Status "Processing user $progressCount of $($userStatuses.Count): $($status.Username)" -PercentComplete (($progressCount / $userStatuses.Count) * 100)

        # Skip users with previous errors or invalid UPNs
        if ($status.LastError -like 'Invalid UPN Format*' -or (-not $status.ExistsInAD -and -not $status.ExistsInCloud)) {
             Write-Warning "Skipping user $($status.Username) due to previous error or not found in AD/Cloud."
             $result = $status.PSObject.Copy() # Start with existing status info
             $result | Add-Member -MemberType NoteProperty -Name TimeStamp -Value (Get-Date)
             $result | Add-Member -MemberType NoteProperty -Name ActionTaken -Value "Skipped"
             $result | Add-Member -MemberType NoteProperty -Name NewADEnabled -Value $status.ADEnabled # No change attempted
             $result | Add-Member -MemberType NoteProperty -Name NewCloudEnabled -Value $status.CloudEnabled # No change attempted
             $results += $result
             continue # Move to the next user
        }

        $result = $status.PSObject.Copy() # Start with existing status info
        $result | Add-Member -MemberType NoteProperty -Name TimeStamp -Value (Get-Date)
        $result | Add-Member -MemberType NoteProperty -Name ActionTaken -Value "None"
        $result.LastError = $null # Clear previous check errors, track action errors now

        $adActionTaken = $false
        $cloudActionTaken = $false

        # Action in AD
        if ($status.ExistsInAD) {
            # Check if action is actually needed
            if ($status.ADEnabled -ne $actionState) {
                try {
                    Write-Verbose "$($actionVerbGerund) AD account: $($status.Username)"
                    if ($TestMode) {
                        $result.ActionTaken = "Test: $actionVerb AD"
                         $adActionTaken = $true
                        Write-Host "[TEST MODE] Would $($actionVerb.ToLower()) AD account: $($status.Username)" -ForegroundColor Yellow
                    } else {
                        if ($EnableAccounts) {
                             Enable-ADAccount -Identity $status.ADDistinguishedName #-WhatIf:$TestMode # WhatIf doesn't work well with TestMode switch logic
                        } else {
                             Disable-ADAccount -Identity $status.ADDistinguishedName #-WhatIf:$TestMode
                        }
                        $result.ActionTaken = "$actionVerb AD"
                         $adActionTaken = $true
                        Write-Host "$($actionVerb)d AD account: $($status.Username)" -ForegroundColor White
                    }
                }
                catch {
                    $errorMsg = "Failed to $($actionVerb.ToLower()) AD account '$($status.Username)': $_"
                    Write-Warning $errorMsg
                    $result.LastError = $errorMsg
                }
            } else {
                 Write-Verbose "AD account '$($status.Username)' is already $($actionVerb.ToLower())d. No action needed."
                $result.ActionTaken = "AD Already $($actionVerb)d"
                 $adActionTaken = $true # Mark as 'handled' even if no change
            }
        } else {
             Write-Verbose "Skipping AD action for $($status.Username) - User not found in AD."
        }

        # Action in Cloud
        if ($status.ExistsInCloud) {
             # Check if action is actually needed
             if ($status.CloudEnabled -ne $actionState) {
                try {
                    Write-Verbose "$($actionVerbGerund) Cloud account: $($status.Username) (ID: $($status.CloudObjectId))"
                     if ($TestMode) {
                         $cloudActionString = "Test: $actionVerb Cloud"
                         $result.ActionTaken = if ($adActionTaken) { "$($result.ActionTaken) & Cloud" } else { $cloudActionString } # Append action string correctly
                         $cloudActionTaken = $true
                        Write-Host "[TEST MODE] Would $($actionVerb.ToLower()) Cloud account: $($status.Username)" -ForegroundColor Yellow
                    } else {
                        Update-MgUser -UserId $status.CloudObjectId -AccountEnabled:$actionState #-WhatIf:$TestMode # WhatIf handled by TestMode switch
                        $cloudActionString = "$actionVerb Cloud"
                         $result.ActionTaken = if ($adActionTaken) { "$($result.ActionTaken) & Cloud" } else { $cloudActionString } # Append action string correctly
                         $cloudActionTaken = $true
                        Write-Host "$($actionVerb)d Cloud account: $($status.Username)" -ForegroundColor White
                    }
                }
                catch {
                    $errorMsg = "Failed to $($actionVerb.ToLower()) Cloud account '$($status.Username)': $_"
                    # Check for specific permission errors
                     if ($_.Exception.Message -match "Insufficient privileges" -or $_.Exception.Message -match "Authorization_RequestDenied") {
                         $errorMsg += " - CHECK PERMISSIONS (User.ReadWrite.All or appropriate role needed)"
                         Write-Warning "$errorMsg" -ForegroundColor Red # Make permission errors more visible
                     } else {
                        Write-Warning $errorMsg
                     }

                    if ($result.LastError) { $result.LastError += " | $errorMsg" } else { $result.LastError = $errorMsg }
                }
            } else {
                Write-Verbose "Cloud account '$($status.Username)' is already $($actionVerb.ToLower())d. No action needed."
                 $cloudActionString = "Cloud Already $($actionVerb)d"
                 $result.ActionTaken = if ($adActionTaken) { "$($result.ActionTaken) & Cloud" } else { $cloudActionString }
                 $cloudActionTaken = $true # Mark as 'handled' even if no change
            }
        } else {
             Write-Verbose "Skipping Cloud action for $($status.Username) - User not found in Cloud."
        }

        # Get updated status only if not in Test Mode and an action was attempted
        if (-not $TestMode -and ($adActionTaken -or $cloudActionTaken) -and -not $result.LastError) {
            Write-Verbose "Pausing briefly and re-checking status for $($status.Username)..."
            Start-Sleep -Seconds 3 # Allow time for changes to potentially propagate (might need adjustment)
            $newStatus = Get-UserStatus -UserPrincipalName $status.Username
            $result | Add-Member -MemberType NoteProperty -Name NewADEnabled -Value $newStatus.ADEnabled
            $result | Add-Member -MemberType NoteProperty -Name NewCloudEnabled -Value $newStatus.CloudEnabled
            # Add any new errors from the status check
             if ($newStatus.LastError -and ($result.LastError -notlike "*$($newStatus.LastError)*")) { # Avoid duplicate errors
                 if ($result.LastError) { $result.LastError += " | PostCheck Error: $($newStatus.LastError)" } else { $result.LastError = "PostCheck Error: $($newStatus.LastError)" }
             }
        } else {
            # In test mode or if no action taken/error occurred, report original status as the 'new' status
            $result | Add-Member -MemberType NoteProperty -Name NewADEnabled -Value (if ($TestMode) { "No Change (Test)" } else { $status.ADEnabled })
            $result | Add-Member -MemberType NoteProperty -Name NewCloudEnabled -Value (if ($TestMode) { "No Change (Test)" } else { $status.CloudEnabled })
        }

        # If no action was needed or taken in either dir, update ActionTaken field
        if (-not $adActionTaken -and -not $cloudActionTaken) {
             $result.ActionTaken = "No Action Needed"
        } elseif ($result.ActionTaken -eq "None") { # Catch cases where AD/Cloud existed but state was already correct
            $actionStatus = if ($EnableAccounts) { "Enabled" } else { "Disabled" }
            $result.ActionTaken = "Already $actionStatus"
        }


        $results += $result
    }
    Write-Progress -Activity "$actionVerbGerund Accounts" -Completed

    # Export and display results
    if ($results) {
        $results | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
        Write-Host "`nFinal Results:" -ForegroundColor Green
        # Select relevant columns for final display
        $results | Format-Table Username, ExistsInAD, ADEnabled, ExistsInCloud, CloudEnabled, ActionTaken, NewADEnabled, NewCloudEnabled, LastError -AutoSize
        Write-Host "`nLog file saved to: $logFile" -ForegroundColor Cyan
    } else {
        Write-Host "`nNo results to log." -ForegroundColor Yellow
    }
}
# Replace the original main catch block (around line 344) with this:
catch {
    # Catch script-terminating errors
    $errorMessage = "SCRIPT FAILED."
    if ($_) { # Check if there's an actual error object
        $errorMessage += " Error details: $($_.ToString())"
        if ($_.Exception) {
            $errorMessage += " | Exception Message: $($_.Exception.Message)"
        }
        if ($_.ScriptStackTrace) {
            $errorMessage += " | StackTrace: $($_.ScriptStackTrace)"
        }
    } else {
        $errorMessage += " An unknown error occurred."
    }

    Write-Error $errorMessage # Write the composed error message

    # Log the error details if possible
    # Use a separate variable for the log message to avoid issues with Add-Content if $errorMessage is complex
    $logErrorMessage = "SCRIPT FAILED at $(Get-Date): $errorMessage"

    if ($logFile -and (Test-Path (Split-Path $logFile -Parent) -PathType Container)) { # Also check if log directory exists
        try {
            Add-Content -Path $logFile -Value $logErrorMessage -Encoding UTF8 -ErrorAction Stop
            Write-Host "Error details logged to $logFile" -ForegroundColor Red
        } catch {
            Write-Warning "Could not write error details to log file '$logFile'. Error: $($_.ToString())"
        }
    } elseif ($logFile) {
         Write-Warning "Log directory for '$logFile' not found. Cannot log error details."
    } else {
         Write-Warning "Log file variable not set. Cannot log error details."
    }

    exit 1 # Ensure the script exits with a non-zero code
}
# The 'finally' block remains the same after this catch block
finally {
    # Ensure Graph connection is disconnected
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}