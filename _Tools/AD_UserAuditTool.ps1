#=====================================#
#               AODR                  #
#=====================================#

<#
.SYNOPSIS
    Active Directory Object Dependency Reporter (AODR)
    
.DESCRIPTION
    This script helps identify what objects are connected to or contained within 
    a specified Active Directory object before deletion.
    
.NOTES
    Outline:
    1. Prompt: User account to check? [Enter samAccountName]
    2. Prompt: Include group memberships? [Y/N]
    3. Prompt: Include owned objects? [Y/N]
    4. Prompt: Include linked mailboxes? [Y/N]
    5. Prompt: Include connected SharePoint sites? [Y/N]
#>

# Import required modules
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    Write-Host "ActiveDirectory module is required. Please install RSAT tools." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

function Get-ADObjectContents {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [switch]$IncludeGroups = $true,
        [switch]$IncludeOwnedObjects = $true,
        [switch]$IncludeMailbox = $true,
        [switch]$IncludeSharePoint = $true
    )

    Write-Host "`n`n#=====================================#" -ForegroundColor Cyan
    Write-Host "#               AODR                  #" -ForegroundColor Cyan
    Write-Host "#=====================================#`n" -ForegroundColor Cyan

    # Try to get user by samAccountName first, then by distinguishedName if that fails
    try {
        $user = Get-ADUser -Identity $Identity -Properties * -ErrorAction Stop
        Write-Host "Found user: $($user.Name) ($($user.SamAccountName))" -ForegroundColor Green
    }
    catch {
        try {
            # Try searching by distinguishedName instead
            $user = Get-ADObject -Identity $Identity -Properties * -ErrorAction Stop
            Write-Host "Found AD object: $($user.Name) ($($user.ObjectClass))" -ForegroundColor Green
        }
        catch {
            Write-Host "Error: Could not find AD object with identity '$Identity'." -ForegroundColor Red
            Write-Host "Please try using samAccountName, distinguishedName, or objectGUID." -ForegroundColor Yellow
            return
        }
    }

    Write-Host "`n[EXAMINING OBJECT DEPENDENCIES]" -ForegroundColor Magenta
    Write-Host "--------------------------------" -ForegroundColor Magenta

    # 1. Check for direct child objects
    Write-Host "`n[1] DIRECT CHILD OBJECTS" -ForegroundColor Cyan
    Write-Host "   -------------------" -ForegroundColor Cyan
    
    $childObjects = Get-ADObject -Filter "distinguishedName -like '*,$($user.DistinguishedName)'" -Properties *
    
    if ($childObjects.Count -eq 0) {
        Write-Host "   No direct child objects found." -ForegroundColor Gray
    }
    else {
        Write-Host "   Found $($childObjects.Count) child objects:" -ForegroundColor Yellow
        $childObjects | ForEach-Object {
            Write-Host "   - $($_.Name) ($($_.ObjectClass))" -ForegroundColor White
        }
    }

    # 2. Check group memberships if requested
    if ($IncludeGroups) {
        Write-Host "`n[2] GROUP MEMBERSHIPS" -ForegroundColor Cyan
        Write-Host "   -----------------" -ForegroundColor Cyan
        
        if ($user.MemberOf.Count -eq 0) {
            Write-Host "   No group memberships found." -ForegroundColor Gray
        }
        else {
            Write-Host "   Member of $($user.MemberOf.Count) groups:" -ForegroundColor Yellow
            foreach ($group in $user.MemberOf) {
                try {
                    $groupName = (Get-ADGroup $group).Name
                    Write-Host "   - $groupName" -ForegroundColor White
                }
                catch {
                    Write-Host "   - $group" -ForegroundColor White
                }
            }
        }
    }

    # 3. Check for owned objects if requested
    if ($IncludeOwnedObjects) {
        Write-Host "`n[3] OWNED OBJECTS" -ForegroundColor Cyan
        Write-Host "   -------------" -ForegroundColor Cyan
        
        $ownedObjects = Get-ADObject -Filter "managedBy -eq '$($user.DistinguishedName)'" -Properties *
        
        if ($ownedObjects.Count -eq 0) {
            Write-Host "   No owned objects found." -ForegroundColor Gray
        }
        else {
            Write-Host "   Owns $($ownedObjects.Count) objects:" -ForegroundColor Yellow
            $ownedObjects | ForEach-Object {
                Write-Host "   - $($_.Name) ($($_.ObjectClass))" -ForegroundColor White
            }
        }
    }

    # 4. Check for Exchange mailbox if requested and Exchange module is available
    if ($IncludeMailbox) {
        Write-Host "`n[4] EXCHANGE MAILBOX" -ForegroundColor Cyan
        Write-Host "   ---------------" -ForegroundColor Cyan
        
        if (Get-Module -Name ExchangeOnlineManagement -ListAvailable) {
            try {
                # Check if already connected to Exchange Online
                $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                
                if (-not $exchangeSession) {
                    Write-Host "   Exchange Online session not connected. Please run Connect-ExchangeOnline first." -ForegroundColor Yellow
                }
                else {
                    $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
                    
                    if ($mailbox) {
                        Write-Host "   Mailbox found:" -ForegroundColor Yellow
                        Write-Host "   - Type: $($mailbox.RecipientTypeDetails)" -ForegroundColor White
                        Write-Host "   - Display Name: $($mailbox.DisplayName)" -ForegroundColor White
                        Write-Host "   - Primary SMTP: $($mailbox.PrimarySmtpAddress)" -ForegroundColor White
                        Write-Host "   - Archive Status: $($mailbox.ArchiveStatus)" -ForegroundColor White
                    }
                    else {
                        Write-Host "   No mailbox found for this user." -ForegroundColor Gray
                    }
                }
            }
            catch {
                Write-Host "   Error checking mailbox: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "   ExchangeOnlineManagement module not installed. Cannot check mailbox." -ForegroundColor Yellow
            Write-Host "   Install with: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Gray
        }
    }

    # 5. Check for SharePoint sites if requested and PnP module is available
    if ($IncludeSharePoint) {
        Write-Host "`n[5] SHAREPOINT SITES" -ForegroundColor Cyan
        Write-Host "   ---------------" -ForegroundColor Cyan
        
        if (Get-Module -Name PnP.PowerShell -ListAvailable) {
            try {
                # Check if already connected to SharePoint Online
                $pnpConnection = Get-PnPConnection -ErrorAction SilentlyContinue
                
                if (-not $pnpConnection) {
                    Write-Host "   PnP connection not established. Please run Connect-PnPOnline first." -ForegroundColor Yellow
                }
                else {
                    Write-Host "   Checking for sites where user is owner or member..." -ForegroundColor Gray
                    Write-Host "   Note: This is limited to the current SharePoint connection context." -ForegroundColor Gray
                    
                    # This is a simplified approach and may need customization
                    # A comprehensive check would require iterating through all site collections
                }
            }
            catch {
                Write-Host "   Error checking SharePoint sites: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "   PnP.PowerShell module not installed. Cannot check SharePoint sites." -ForegroundColor Yellow
            Write-Host "   Install with: Install-Module -Name PnP.PowerShell" -ForegroundColor Gray
        }
    }

    # 6. Final section: Deletion recommendations
    Write-Host "`n[DELETION RECOMMENDATIONS]" -ForegroundColor Magenta
    Write-Host "--------------------------" -ForegroundColor Magenta
    
    if ($childObjects.Count -gt 0) {
        Write-Host "`nWARNING: This object contains $($childObjects.Count) child objects." -ForegroundColor Red
        Write-Host "If you delete this object, all child objects will also be deleted." -ForegroundColor Red
        Write-Host "`nRecommendation: Review the child objects and consider moving or" -ForegroundColor Yellow
        Write-Host "reassigning them before deleting this object." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nNo child objects detected that would prevent deletion." -ForegroundColor Green
    }

    if ($user.MemberOf.Count -gt 0 -or $ownedObjects.Count -gt 0) {
        Write-Host "`nBefore deletion, consider:" -ForegroundColor Yellow
        if ($user.MemberOf.Count -gt 0) {
            Write-Host "- Removing from $($user.MemberOf.Count) groups" -ForegroundColor White
        }
        if ($ownedObjects.Count -gt 0) {
            Write-Host "- Reassigning ownership of $($ownedObjects.Count) objects" -ForegroundColor White
        }
    }

    Write-Host "`n#=====================================#" -ForegroundColor Cyan
}

# Main script execution
Clear-Host

Write-Host "#=====================================#" -ForegroundColor Cyan
Write-Host "#               AODR                  #" -ForegroundColor Cyan
Write-Host "#=====================================#" -ForegroundColor Cyan
Write-Host "Active Directory Object Dependency Reporter`n" -ForegroundColor White

# Get user input
$identity = Read-Host "Enter user account name (samAccountName or DN)"
$includeGroups = (Read-Host "Include group memberships? (Y/N) [Y]").ToUpper() -ne "N"
$includeOwned = (Read-Host "Include owned objects? (Y/N) [Y]").ToUpper() -ne "N"
$includeMailbox = (Read-Host "Include linked mailboxes? (Y/N) [Y]").ToUpper() -ne "N"
$includeSP = (Read-Host "Include connected SharePoint sites? (Y/N) [N]").ToUpper() -eq "Y"

# Run the function with the specified parameters
Get-ADObjectContents -Identity $identity -IncludeGroups:$includeGroups -IncludeOwnedObjects:$includeOwned -IncludeMailbox:$includeMailbox -IncludeSharePoint:$includeSP