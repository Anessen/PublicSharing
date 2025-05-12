# TABL.ps1
# Tenant Allow/Block List Management Script

# ASCII Header
function Show-Header {
    Write-Host "
 _____  _    ____  _     
|_   _|/ \  | __ )| |    
  | | / _ \ |  _ \| |    
  | |/ ___ \| |_) | |___ 
  |_/_/   \_\____/|_____|
                         
 Tenant Allow/Block List Manager
" -ForegroundColor Cyan
}

# Check and import required module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Please install it using:" -ForegroundColor Red
    Write-Host "Install-Module -Name ExchangeOnlineManagement -Force" -ForegroundColor Yellow
    exit
}

# Connect to Exchange Online
function Connect-ExchangeOnline {
    try {
        # Check if already connected
        $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $connectionStatus) {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
            
            # Try to connect with enhanced error handling
            try {
                # Use device code authentication which is more reliable in VSCode
                Connect-ExchangeOnline -UseDeviceAuthentication -ShowBanner:$false
                Write-Host "Connected successfully!" -ForegroundColor Green
            }
            catch {
                $errorMessage = $_.Exception.Message
                
                # Check for common authentication errors
                if ($errorMessage -like "*0x80070520*" -or $errorMessage -like "*Error Acquiring Token*") {
                    Write-Host "`n=== AUTHENTICATION ERROR DETECTED ===`n" -ForegroundColor Red
                    Write-Host "This is a known issue with VSCode and Exchange Online authentication." -ForegroundColor Yellow
                    Write-Host "Try one of these solutions:" -ForegroundColor Yellow
                    Write-Host "1. Run this script in a regular PowerShell window instead of VSCode terminal" -ForegroundColor White
                    Write-Host "2. Try running VSCode as administrator" -ForegroundColor White
                    Write-Host "3. Clear your token cache:" -ForegroundColor White
                    Write-Host "   - Close all PowerShell and VSCode windows" -ForegroundColor White
                    Write-Host "   - Delete files in: %USERPROFILE%\.IdentityService\msal.cache" -ForegroundColor White
                    Write-Host "   - Restart your computer and try again" -ForegroundColor White
                    Write-Host "4. Try using the -UseDeviceAuthentication parameter (already attempted)" -ForegroundColor White
                    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
                    Read-Host
                    exit
                }
                else {
                    # Re-throw the error for other types of errors
                    throw
                }
            }
        }
    }
    catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        Write-Host "`nTry running this script in a regular PowerShell window instead of VSCode terminal." -ForegroundColor Yellow
        exit
    }
}

# Display current allow/block lists
function Show-CurrentLists {
    # Ask for report type
    Write-Host "`nReport Type:" -ForegroundColor Yellow
    Write-Host "1. Concise (key fields only)" -ForegroundColor White
    Write-Host "2. Full (all details)" -ForegroundColor White
    $reportType = Read-Host "`nSelect report type (1-2)"
    
    Write-Host "`n=== CURRENT TENANT ALLOW/BLOCK LISTS ===" -ForegroundColor Cyan
    
    Write-Host "`nSENDER ALLOW LIST:" -ForegroundColor Green
    $senderAllowList = Get-TenantAllowBlockListItems -ListType Sender -Allow
    if ($senderAllowList.Count -eq 0) {
        Write-Host "  No entries found" -ForegroundColor Yellow
    } else {
        foreach ($item in $senderAllowList) {
            Write-Host "`nVALUE: $($item.Value)" -ForegroundColor Green
            Write-Host "  Action:          $($item.Action)" -ForegroundColor White
            Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
            Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
            Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
            Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
            
            if ($reportType -eq "2") {
                # Display all other properties for full report
                $item | Get-Member -MemberType Properties | Where-Object { 
                    $_.Name -notin @('Value', 'Action', 'LastModifiedBy', 'Notes', 'CreatedDateTime', 'ExpirationDateTime') 
                } | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $item.$propName
                    Write-Host "  $($propName): $($propValue ?? 'N/A')" -ForegroundColor Gray
                }
            }
        }
    }
    
    Write-Host "`nSENDER BLOCK LIST:" -ForegroundColor Red
    $senderBlockList = Get-TenantAllowBlockListItems -ListType Sender -Block
    if ($senderBlockList.Count -eq 0) {
        Write-Host "  No entries found" -ForegroundColor Yellow
    } else {
        foreach ($item in $senderBlockList) {
            Write-Host "`nVALUE: $($item.Value)" -ForegroundColor Red
            Write-Host "  Action:          $($item.Action)" -ForegroundColor White
            Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
            Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
            Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
            Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
            
            if ($reportType -eq "2") {
                # Display all other properties for full report
                $item | Get-Member -MemberType Properties | Where-Object { 
                    $_.Name -notin @('Value', 'Action', 'LastModifiedBy', 'Notes', 'CreatedDateTime', 'ExpirationDateTime') 
                } | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $item.$propName
                    Write-Host "  $($propName): $($propValue ?? 'N/A')" -ForegroundColor Gray
                }
            }
        }
    }
    
    Write-Host "`nURL ALLOW LIST:" -ForegroundColor Green
    $urlAllowList = Get-TenantAllowBlockListItems -ListType Url -Allow
    if ($urlAllowList.Count -eq 0) {
        Write-Host "  No entries found" -ForegroundColor Yellow
    } else {
        foreach ($item in $urlAllowList) {
            Write-Host "`nVALUE: $($item.Value)" -ForegroundColor Green
            Write-Host "  Action:          $($item.Action)" -ForegroundColor White
            Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
            Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
            Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
            Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
            
            if ($reportType -eq "2") {
                # Display all other properties for full report
                $item | Get-Member -MemberType Properties | Where-Object { 
                    $_.Name -notin @('Value', 'Action', 'LastModifiedBy', 'Notes', 'CreatedDateTime', 'ExpirationDateTime') 
                } | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $item.$propName
                    Write-Host "  $($propName): $($propValue ?? 'N/A')" -ForegroundColor Gray
                }
            }
        }
    }
    
    Write-Host "`nURL BLOCK LIST:" -ForegroundColor Red
    $urlBlockList = Get-TenantAllowBlockListItems -ListType Url -Block
    if ($urlBlockList.Count -eq 0) {
        Write-Host "  No entries found" -ForegroundColor Yellow
    } else {
        foreach ($item in $urlBlockList) {
            Write-Host "`nVALUE: $($item.Value)" -ForegroundColor Red
            Write-Host "  Action:          $($item.Action)" -ForegroundColor White
            Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
            Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
            Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
            Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
            
            if ($reportType -eq "2") {
                # Display all other properties for full report
                $item | Get-Member -MemberType Properties | Where-Object { 
                    $_.Name -notin @('Value', 'Action', 'LastModifiedBy', 'Notes', 'CreatedDateTime', 'ExpirationDateTime') 
                } | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $item.$propName
                    Write-Host "  $($propName): $($propValue ?? 'N/A')" -ForegroundColor Gray
                }
            }
        }
    }
    
    Write-Host "`n==========================================`n" -ForegroundColor Cyan
}

# Search for emails
function Search-Emails {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Sender", "Recipient")]
        [string]$SearchType,
        
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory=$true)]
        [datetime]$EndDate
    )
    
    $searchQuery = ""
    if ($SearchType -eq "Sender") {
        $searchQuery = "from:$EmailAddress"
    }
    else {
        $searchQuery = "to:$EmailAddress"
    }
    
    try {
        Write-Host "Searching for emails $SearchType $EmailAddress between $StartDate and $EndDate..." -ForegroundColor Cyan
        $emails = Get-MessageTrace -SenderAddress $EmailAddress -StartDate $StartDate -EndDate $EndDate | 
                 Select-Object Received, SenderAddress, RecipientAddress, Subject, Status
        
        if ($emails.Count -eq 0) {
            Write-Host "No emails found matching the criteria." -ForegroundColor Yellow
            return $null
        }
        
        return $emails
    }
    catch {
        Write-Host "Error searching emails: $_" -ForegroundColor Red
        return $null
    }
}

# Display emails with sorting options
function Show-EmailsWithSorting {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Emails
    )
    
    Write-Host "`n=== FOUND EMAILS (Total: $($Emails.Count)) ===`n" -ForegroundColor Cyan
    
    # Sorting options
    $sortOptions = @(
        "Received (newest first)",
        "Received (oldest first)",
        "Sender (A-Z)",
        "Recipient (A-Z)",
        "Subject (A-Z)"
    )
    
    for ($i = 0; $i -lt $sortOptions.Count; $i++) {
        Write-Host "$($i+1). Sort by $($sortOptions[$i])" -ForegroundColor Yellow
    }
    
    $sortChoice = Read-Host "`nChoose sorting option (1-$($sortOptions.Count))"
    $sortedEmails = @()
    
    switch ($sortChoice) {
        "1" { $sortedEmails = $Emails | Sort-Object Received -Descending }
        "2" { $sortedEmails = $Emails | Sort-Object Received }
        "3" { $sortedEmails = $Emails | Sort-Object SenderAddress }
        "4" { $sortedEmails = $Emails | Sort-Object RecipientAddress }
        "5" { $sortedEmails = $Emails | Sort-Object Subject }
        default { 
            Write-Host "Invalid choice. Showing unsorted results." -ForegroundColor Yellow 
            $sortedEmails = $Emails
        }
    }
    
    # Display emails
    $index = 1
    foreach ($email in $sortedEmails) {
        Write-Host "`n$index. " -ForegroundColor Yellow -NoNewline
        Write-Host "Date: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($email.Received)" -ForegroundColor White
        Write-Host "   From: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($email.SenderAddress)" -ForegroundColor White
        Write-Host "   To: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($email.RecipientAddress)" -ForegroundColor White
        Write-Host "   Subject: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($email.Subject)" -ForegroundColor White
        Write-Host "   Status: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($email.Status)" -ForegroundColor White
        $index++
    }
    
    return $sortedEmails
}

# Add address to allow/block list
function Add-ToAllowBlockList {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Allow", "Block")]
        [string]$Action
    )
    
    try {
        $listType = "Sender"
        $currentList = @()
        
        # Get current list for before/after comparison
        if ($Action -eq "Allow") {
            $currentList = Get-TenantAllowBlockListItems -ListType $listType -Allow
        } else {
            $currentList = Get-TenantAllowBlockListItems -ListType $listType -Block
        }
        
        # Ask for expiration preference
        Write-Host "`nExpiration options:" -ForegroundColor Yellow
        Write-Host "1. 45 days (recommended)" -ForegroundColor White
        Write-Host "2. Never expire" -ForegroundColor White
        $expirationChoice = Read-Host "`nSelect expiration option (1-2)"
        
        $expirationDate = $null
        $expirationText = ""
        
        if ($expirationChoice -eq "1") {
            $expirationDate = (Get-Date).AddDays(45)
            $expirationText = "will expire after 45 days ($($expirationDate.ToString('MM/dd/yyyy')))"
        } else {
            $expirationText = "will NEVER expire"
        }
        
        # Get notes
        $notes = Read-Host "`nEnter notes for this entry (optional, press Enter to skip)"
        
        # Confirm changes
        Write-Host "`n=== CONFIRM CHANGES ===" -ForegroundColor Yellow
        Write-Host "You are about to:" -ForegroundColor White
        Write-Host "- Add email address: " -ForegroundColor White -NoNewline
        Write-Host "$EmailAddress" -ForegroundColor Green
        Write-Host "- To the Sender " -ForegroundColor White -NoNewline
        Write-Host "$Action" -ForegroundColor Green -NoNewline
        Write-Host " List" -ForegroundColor White
        Write-Host "- This entry $expirationText" -ForegroundColor White
        if ($notes) {
            Write-Host "- With notes: " -ForegroundColor White -NoNewline
            Write-Host "$notes" -ForegroundColor Green
        }
        
        $confirm = Read-Host "`nAre you sure you want to make these changes? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return $false
        }
        
        # Apply changes
        if ($Action -eq "Allow") {
            Write-Host "Adding $EmailAddress to the sender allow list..." -ForegroundColor Cyan
            if ($expirationChoice -eq "1") {
                if ($notes) {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Allow -ExpirationDate $expirationDate -Notes $notes
                } else {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Allow -ExpirationDate $expirationDate
                }
            } else {
                if ($notes) {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Allow -Notes $notes
                } else {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Allow
                }
            }
            Write-Host "Successfully added $EmailAddress to the sender allow list!" -ForegroundColor Green
        }
        else {
            Write-Host "Adding $EmailAddress to the sender block list..." -ForegroundColor Cyan
            if ($expirationChoice -eq "1") {
                if ($notes) {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Block -ExpirationDate $expirationDate -Notes $notes
                } else {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Block -ExpirationDate $expirationDate
                }
            } else {
                if ($notes) {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Block -Notes $notes
                } else {
                    $result = New-TenantAllowBlockListItems -ListType $listType -Entries $EmailAddress -Block
                }
            }
            Write-Host "Successfully added $EmailAddress to the sender block list!" -ForegroundColor Green
        }
        
        # Show before/after
        Write-Host "`n=== BEFORE CHANGES ===" -ForegroundColor Yellow
        if ($currentList.Count -eq 0) {
            Write-Host "  No previous entries" -ForegroundColor Yellow
        } else {
            foreach ($item in $currentList) {
                Write-Host "`nVALUE: $($item.Value)" -ForegroundColor White
                Write-Host "  Action:          $($item.Action)" -ForegroundColor White
                Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
                Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
                Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
                Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
            }
        }
        
        Write-Host "`n=== AFTER CHANGES ===" -ForegroundColor Green
        $updatedList = @()
        if ($Action -eq "Allow") {
            $updatedList = Get-TenantAllowBlockListItems -ListType $listType -Allow
        } else {
            $updatedList = Get-TenantAllowBlockListItems -ListType $listType -Block
        }
        
        foreach ($item in $updatedList) {
            Write-Host "`nVALUE: $($item.Value)" -ForegroundColor White
            Write-Host "  Action:          $($item.Action)" -ForegroundColor White
            Write-Host "  ModifiedBy:      $($item.LastModifiedBy ?? 'N/A')" -ForegroundColor White
            Write-Host "  Notes:           $($item.Notes ?? 'N/A')" -ForegroundColor White
            Write-Host "  CreatedDateTime: $($item.CreatedDateTime)" -ForegroundColor White
            Write-Host "  ExpirationDate:  $($item.ExpirationDateTime ?? 'Never')" -ForegroundColor White
        }
        
        return $true
    }
    catch {
        Write-Host "Error adding to $Action list: $_" -ForegroundColor Red
        return $false
    }
}

# Main menu
function Show-MainMenu {
    Clear-Host
    Show-Header
    Write-Host "1. Search emails by sender" -ForegroundColor Yellow
    Write-Host "2. Search emails by recipient" -ForegroundColor Yellow
    Write-Host "3. View current allow/block lists" -ForegroundColor Yellow
    Write-Host "4. Quit" -ForegroundColor Yellow
    
    $choice = Read-Host "`nSelect an option (1-4)"
    
    switch ($choice) {
        "1" { Search-EmailsMenu "Sender" }
        "2" { Search-EmailsMenu "Recipient" }
        "3" { 
            Show-CurrentLists
            Read-Host "Press Enter to continue"
            Show-MainMenu 
        }
        "4" { 
            Write-Host "Exiting script..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            exit 
        }
        default { 
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-MainMenu 
        }
    }
}

# Email search menu
function Search-EmailsMenu {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Sender", "Recipient")]
        [string]$SearchType
    )
    
    Clear-Host
    Write-Host "===== SEARCH EMAILS BY $SearchType =====" -ForegroundColor Cyan
    
    $emailAddress = Read-Host "Enter $SearchType email address"
    
    Write-Host "`nDate Range Options:" -ForegroundColor Yellow
    Write-Host "- Enter 0 for current day (until end of today)" -ForegroundColor White
    Write-Host "- Enter number of days to go back (e.g., 7 for last week)" -ForegroundColor White
    Write-Host "- Maximum search range is typically 10 days for standard message trace" -ForegroundColor White
    
    $daysBackStart = Read-Host "`nEnter days back for START date"
    $daysBackEnd = Read-Host "Enter days back for END date (0 for today)"
    
    $startDate = [datetime]::MinValue
    $endDate = [datetime]::MaxValue
    
    try {
        # Convert days back to actual dates
        $endDate = (Get-Date).Date.AddDays(-[int]$daysBackEnd).AddDays(1).AddSeconds(-1) # End of the day
        $startDate = (Get-Date).Date.AddDays(-[int]$daysBackStart)
        
        # Validate that start date is before or equal to end date
        if ($startDate -gt $endDate) {
            Write-Host "Start date cannot be after end date. Swapping dates." -ForegroundColor Yellow
            $tempDate = $startDate
            $startDate = $endDate
            $endDate = $tempDate
        }
        
        Write-Host "`nSearching from $($startDate.ToString('MM/dd/yyyy HH:mm:ss')) to $($endDate.ToString('MM/dd/yyyy HH:mm:ss'))" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Invalid input. Using default date range (last 7 days)." -ForegroundColor Yellow
        $endDate = (Get-Date).Date.AddDays(1).AddSeconds(-1) # End of today
        $startDate = (Get-Date).Date.AddDays(-7) # 7 days ago
    }
    
    $emails = Search-Emails -SearchType $SearchType -EmailAddress $emailAddress -StartDate $startDate -EndDate $endDate
    
    if ($null -eq $emails -or $emails.Count -eq 0) {
        Read-Host "Press Enter to return to main menu"
        Show-MainMenu
        return
    }
    
    $sortedEmails = Show-EmailsWithSorting -Emails $emails
    
    Write-Host "`n=== ACTIONS ===" -ForegroundColor Cyan
    Write-Host "1. Add a sender to the allow list" -ForegroundColor Yellow
    Write-Host "2. Add a sender to the block list" -ForegroundColor Yellow
    Write-Host "3. Return to main menu" -ForegroundColor Yellow
    
    $actionChoice = Read-Host "`nSelect an action (1-3)"
    
    switch ($actionChoice) {
        "1" { 
            $emailIndex = Read-Host "Enter the number of the email to whitelist the sender"
            if ([int]$emailIndex -gt 0 -and [int]$emailIndex -le $sortedEmails.Count) {
                $selectedEmail = $sortedEmails[[int]$emailIndex - 1]
                Add-ToAllowBlockList -EmailAddress $selectedEmail.SenderAddress -Action "Allow"
            }
            else {
                Write-Host "Invalid email selection." -ForegroundColor Red
            }
            Read-Host "Press Enter to continue"
            Show-MainMenu
        }
        "2" { 
            $emailIndex = Read-Host "Enter the number of the email to blacklist the sender"
            if ([int]$emailIndex -gt 0 -and [int]$emailIndex -le $sortedEmails.Count) {
                $selectedEmail = $sortedEmails[[int]$emailIndex - 1]
                Add-ToAllowBlockList -EmailAddress $selectedEmail.SenderAddress -Action "Block"
            }
            else {
                Write-Host "Invalid email selection." -ForegroundColor Red
            }
            Read-Host "Press Enter to continue"
            Show-MainMenu
        }
        "3" { Show-MainMenu }
        default { 
            Write-Host "Invalid option. Returning to main menu." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-MainMenu 
        }
    }
}

# Start the script
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline
    Show-MainMenu
}
catch {
    Write-Host "Error starting script: $_" -ForegroundColor Red
}
