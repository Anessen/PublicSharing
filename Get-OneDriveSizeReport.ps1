<#
    .SYNOPSIS
    Get-OneDriveSizeReport.ps1

    .DESCRIPTION
    Export OneDrive storage usage in Microsoft 365 to CSV file with PowerShell.

    .LINK
    www.alitajran.com/export-onedrive-usage-report/

    .NOTES
    Written by: ALI TAJRAN
    Website:    www.alitajran.com
    LinkedIn:   linkedin.com/in/alitajran

    .CHANGELOG
    V1.00, 08/10/2024 - Initial version
#>

# Connect to Microsoft Graph with necessary permissions
Connect-MgGraph -NoWelcome -Scopes "User.Read.All", "Reports.Read.All", "ReportSettings.ReadWrite.All"

# Define file paths
$CSVOutputFile = "C:\temp\OneDriveSizeReport.csv"
$TempExportFile = "C:\temp\TempExportFile.csv"

# Remove the temporary export file if it exists
if (Test-Path $TempExportFile) {
    Remove-Item $TempExportFile
}

# Check if tenant reports have concealed user data, and adjust settings if necessary
if ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $true) {
    $Parameters = @{ displayConcealedNames = $false }
    Write-Host "Unhiding concealed report data to retrieve full user information..." -ForegroundColor Cyan
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
    $ConcealedFlag = $true
}
else {
    $ConcealedFlag = $false
    Write-Host "User data is already fully visible in the reports." -ForegroundColor Cyan
}

# Retrieve detailed user account information
Write-Host "Fetching user account details from Microsoft Graph..." -ForegroundColor Cyan

# Define user properties to be retrieved
$Properties = 'Id', 'displayName', 'userPrincipalName', 'city', 'country', 'department', 'jobTitle', 'officeLocation'

# Define parameters for retrieving users with assigned licenses
$userParams = @{
    All              = $true
    Filter           = "assignedLicenses/`$count ne 0 and userType eq 'Member'"
    ConsistencyLevel = 'Eventual'
    CountVariable    = 'UserCount'
    Sort             = 'displayName'
}

# Get user account information and select the desired properties
$Users = Get-MgUser @UserParams -Property $Properties | Select-Object -Property $Properties

# Create a hashtable to map UPNs (User Principal Names) to user details
$UserHash = @{}
foreach ($User in $Users) {
    $UserHash[$User.userPrincipalName] = $User
}

# Retrieve OneDrive for Business site usage details for the last 30 days and export to a temporary CSV file
Write-Host "Retrieving OneDrive for Business site usage details..." -ForegroundColor Cyan
Get-MgReportOneDriveUsageAccountDetail -Period D30 -Outfile $TempExportFile

# Import the data from the temporary CSV file
$ODFBSites = Import-CSV $TempExportFile | Sort-Object 'User display name'

if (-not $ODFBSites) {
    Write-Host "No OneDrive sites found." -ForegroundColor Yellow
    return
}

# Calculate total storage used by all OneDrive for Business accounts
$TotalODFBGBUsed = [Math]::Round(($ODFBSites.'Storage Used (Byte)' | Measure-Object -Sum).Sum / 1GB, 2)

# Initialize a list to store report data
$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

# Populate the report with detailed information for each OneDrive site
foreach ($Site in $ODFBSites) {
    $UserData = $UserHash[$Site.'Owner Principal name']
    $ReportLine = [PSCustomObject]@{
        Owner             = $Site.'Owner display name'
        UserPrincipalName = $Site.'Owner Principal name'
        SiteId            = $Site.'Site Id'
        IsDeleted         = $Site.'Is Deleted'
        LastActivityDate  = $Site.'Last Activity Date'
        FileCount         = $Site.'File Count'
        ActiveFileCount   = $Site.'Active File Count'
        QuotaGB           = [Math]::Round($Site.'Storage Allocated (Byte)' / 1GB, 2)
        UsedGB            = [Math]::Round($Site.'Storage Used (Byte)' / 1GB, 2)
        PercentUsed       = [Math]::Round($Site.'Storage Used (Byte)' / $Site.'Storage Allocated (Byte)' * 100, 2)
        City              = $UserData.city
        Country           = $UserData.country
        Department        = $UserData.department
        JobTitle          = $UserData.jobTitle
    }
    $Report.Add($ReportLine)
}

# Export the report to a CSV file and display the data in a grid view
$Report | Sort-Object UsedGB -Descending | Export-CSV -NoTypeInformation -Encoding utf8 $CSVOutputFile
$Report | Sort-Object UsedGB -Descending | Out-GridView -Title OneDriveUsageReport

Write-Host ("Current OneDrive for Business storage consumption is {0} GB. Report saved to {1}" -f $TotalODFBGBUsed, $CSVOutputFile) -ForegroundColor Cyan

# Reset tenant report data concealment setting if it was modified earlier
if ($ConcealedFlag -eq $true) {
    Write-Host "Re-enabling data concealment in tenant reports..." -ForegroundColor Cyan
    $Parameters = @{ displayConcealedNames = $true }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# Clean up the temporary export file
if (Test-Path $TempExportFile) {
    Remove-Item $TempExportFile
    Write-Host "Temporary export file removed." -ForegroundColor Cyan
}