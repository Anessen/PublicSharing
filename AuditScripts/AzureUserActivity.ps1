# Import the Microsoft.Graph module
Import-Module Microsoft.Graph

# Connect to Microsoft Graph with required scopes
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

Write-Host "`n`n----- Starting Azure User Activity Script -----`n`n"

# Define the select properties
$select = "id,userPrincipalName,displayName,jobTitle,department,officeLocation,mail,accountEnabled,userType,companyName,createdDateTime,signInActivity"

# Define the expand for manager
$expand = "manager(`$select=userPrincipalName,displayName)"

# Initial URI for paginated request
$uri = "https://graph.microsoft.com/beta/users?`$select=$select&`$expand=$expand&`$top=999"

# Initialize users array
$users = @()

# Fetch users in pages
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $users += $response.value
    $uri = $response.'@odata.nextLink'
} while ($uri)

# Process users into custom objects
$userData = $users | ForEach-Object {
    [PSCustomObject]@{
        UserType                     = $_.userType
        AccountEnabled               = $_.accountEnabled
        UserPrincipalName            = $_.userPrincipalName
        DisplayName                  = $_.displayName
        CreationDate                 = $_.createdDateTime
        lastLogin                    = if ($_.signInActivity) { $_.signInActivity.lastSignInDateTime } else { $null }
        lastSuccessfulSignInDateTime = if ($_.signInActivity) { $_.signInActivity.lastSuccessfulSignInDateTime } else { $null }
        lastNonInteractiveSignInDateTime = if ($_.signInActivity) { $_.signInActivity.lastNonInteractiveSignInDateTime } else { $null }
        lastSignInDateTime           = if ($_.signInActivity) { $_.signInActivity.lastSignInDateTime } else { $null }
        JobTitle                     = $_.jobTitle
        Department                   = $_.department
        OfficeLocation               = $_.officeLocation
        Mail                         = $_.mail
        CompanyName                  = $_.companyName
        ManagerDisplayName           = if ($_.manager) { $_.manager.displayName } else { $null }
        ManagerUPN                   = if ($_.manager) { $_.manager.userPrincipalName } else { $null }
        UserId                       = $_.id
    }
}

$date = Get-Date -Format "MMddyy"
$path = "AzureUserActivity$date.csv"

# Export the data to CSV
$userData | Export-Csv -Path $path -NoTypeInformation

# Confirm the export in the terminal
Write-Output "Script finished. Your export has been saved to: $path in the current directory."
