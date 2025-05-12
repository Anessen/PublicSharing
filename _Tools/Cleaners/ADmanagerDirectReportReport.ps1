# ADOrganization.ps1
# PowerShell program to document the organization specified by
# the manager and directReports attributes in Active Directory.
#
# Copyright (c) 2014-2015 Richard L. Mueller
# PowerShell Version 1.0
# October 15, 2014
# Revised February 27, 2015
#
# ----------------------------------------------------------------------
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the copyright owner above has no warranty, obligations,
# or liability for such use.

Trap {"Error: $_"; Break;}

Function Get-Reports($ReportDN, $ManagerDN, $ManagerName, $Offset)
{
    # Recursive function to document the organization.
    # The first time this function is called it considers managers at
    # the top of the organization hierarchy. These are objects with
    # direct reports but no manager.
    If ($ReportDN -eq "Top")
    {
        # Filter on objects with no manager and at least one direct report.
        $Filter = "(&(!manager=*)(directReports=*))"
    }
    Else
    {
        # The function has been called recursively to deal with a direct report.
        # Output the object that reports to the previous manager.
        If ($Output  -eq "DN")
        {
            Switch ($Format)
            {
                "HTML" {"<li>$ReportDN</li>" | Out-File -FilePath $File -Append}
                "Text" {
                            # Direct reports are indented beneath their manager.
                            "$Offset$ReportDN" | Out-File -FilePath $File -Append
                            # Indent the next level of the hierarchy 4 spaces.
                            $Offset = "$Offset    "
                       }
                "CSV" {"""$ManagerDN"",""$ReportDN""" | Out-File -FilePath $File -Append}
            }
        }
        If ($Output -eq "Name")
        {
            # Escape any forward slash characters with the backslash escape character.
            $ReportDN = $ReportDN.Replace("/", "\/")
            # Use ADSI to bind to the direct report object and retrieve names.
            $Object = [ADSI]"LDAP://$ReportDN"
            $Name = $Object.name
            $NTName = $Object.sAMAccountName
            $ReportName = "$Name ($NTName)"
            Switch ($Format)
            {
                "HTML" {"<li>$ReportName</li>" | Out-File -FilePath $File -Append}
                "Text" {
                            # Direct reports are indented beneath their manager.
                            "$Offset$ReportName" | Out-File -FilePath $File -Append
                            # Indent the next level of the hierarchy 4 spaces.
                            $Offset = "$Offset    "
                       }
                "CSV" {"""$ManagerName"",""$ReportName""" | Out-File -FilePath $File -Append}
            }
        }
        # Search for all objects that report to this object.
        $Filter = "(manager=$ReportDN)"
    }

    # Run the query.
    $Searcher.Filter = $Filter

    $Results = $Searcher.FindAll()
    If ($Results.Count -gt 0)
    {
        If ($Format -eq "HTML")
        {
            "<ul>"  | Out-File -FilePath $File -Append
        }
        ForEach ($Result In $Results)
        {
            # Output the object.
            $DN = $Result.Properties.Item("distinguishedName")
            If ($Output -eq "DN")
            {
                Switch ($Format)
                {
                    "HTML" {$Line = "<li>$DN</li>"}
                    "Text" {$Line = "$Offset$DN"}
                    "CSV"
                    {
                        # When $ReportDN is "Top", there is no manager.
                        # Enclose DN values in quotes. There will be embedded commas.
                        If ($ReportDN -ne "Top") {$Line = """$ReportDN"",""$DN"""}
                        Else {$Line = ",""$DN"""}
                    }
                }
            }
            If ($Output -eq "Name")
            {
                # Retrieve name and sAMAccountName.
                $Name = $Result.Properties.Item("name")
                $NTName = $Result.Properties.Item("sAMAccountName")
                Switch ($Format)
                {
                    "HTML" {$Line = "<li>$Name ($NTName)</li>"}
                    "Text" {$Line = "$Offset$Name ($NTName)"}
                    "CSV"
                    {
                        # When $ReportDN is "Top", there is no manager.
                        # Enclose values in quotes. There could be embedded commas.
                        If ($ReportDN -ne "Top")
                        {$Line = """$ReportName"",""$Name ($NTName"""}
                        Else {$Line = ",""$DN"""}
                    }
                }
            }
            $Line | Out-File -FilePath $File -Append

            # Retrieve any direct reports for this object.
            $Reports = $Result.Properties.Item("directReports")
            If ($Reports.Count -gt 0)
            {
                If ($Format -eq "HTML")
                {
                    "<ul>" | Out-File -FilePath $File -Append
                }
                ForEach ($Report In $Reports)
                {
                    # Recursively call this function for each direct report.
                    # Increase any indenting by 4 more spaces.
                    Get-Reports $Report $DN "$Name ($NTName)" "$Offset    "
                }
                If ($Format -eq "HTML")
                {
                    "</ul>" | Out-File -FilePath $File -Append
                }
            }
        }
        If ($Format -eq "HTML")
        {
            "</ul>"  | Out-File -FilePath $File -Append
        }
    }
}

Function GetHelp()
{
    "The ADOrganization.ps1 script documents the organization hierarchy"
    "specified by the manager and directReports attributes in Active Directory."
    "Optional parameters:"
    "  One of the following to specify the output format (default is -text):"
    "    -html      Output in HTML format to be displayed in a browser"
    "    -csv       Output in CSV format to be displayed in a spreadsheet"
    "    -text      Output in text format to be displayed in notepad"
    "    -help      Display this help information"
    "  One of the following to specify the attributes to output (default is -dn):"
    "    -dn        Display distinguished names"
    "    -name      Display common names and sAMAccountNames"
    "Creates output file ""Organization.htm"" or ""Organization.csv"" or"
    """Organization.txt"" in the current directory, depending on the format."
}

# Check optional parameters indicating output format.
# The default is "Text" format and output "DN" distinguished names.
$Format = "Text"
$Output = "DN"
$Abort = $False

# Process any optional parameters.
If ($Args.Count -gt 2)
{
    "Error: Wrong number of parameters."
    GetHelp
    Break
}

If ($Args.Count -gt 0)
{
    ForEach ($Arg In $Args)
    {
        Switch ($Arg.ToLower())
        {
            "-help"
            {
                GetHelp
                $Abort = $True
            }
            "-html" {$Format = "HTML"}
            "-csv" {$Format = "CSV"}
            "-text" {$Format = "Text"}
            "-dn" {$Output = "DN"}
            "-name" {$Output = "Name"}
            Default
            {
                "Error: Invalid paramter"
                GetHelp
                $Abort = $True
            }
        }
    }
}
# Abort the script if invalid parameter found.
If ($Abort -eq $True){Break}

# Specify the output file, with the appropriate extension.
Switch ($Format)
{
    "HTML" {$File = ".\ADOrganization.htm"}
    "Text" {$File = ".\ADOrganization.txt"}
    "CSV" {$File = ".\ADOrganization.csv"}
}

# Setup the DirectorySearcher object.
$D = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$Domain = [ADSI]"LDAP://$D"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.PageSize = 200
$Searcher.SearchScope = "subtree"
$Searcher.PropertiesToLoad.Add("distinguishedName") > $Null
$Searcher.PropertiesToLoad.Add("directReports") > $Null
If ($Output -eq "Name")
{
    $Searcher.PropertiesToLoad.Add("name") > $Null
    $Searcher.PropertiesToLoad.Add("sAMAccountName") > $Null
}
$Searcher.SearchRoot = "LDAP://" + $Domain.distinguishedName

# Output header lines.
If ($Format -eq "HTML")
{
    "<div style=""font-family:Courier New,Courier"">" | Out-File -FilePath $File
    "<h1>Organization: $D</h1>" | Out-File -FilePath $File -Append
}
If ($Format -eq "Text")
{
    "Organization: $D"  | Out-File -FilePath $File
}
If ($Format -eq "CSV")
{
    "Manager,Direct Report" | Out-File -FilePath $File -Encoding ASCII
}

# Retrieve organization hierarchy, starting from the top.
Get-Reports "Top" "" "" ""

# Output final tag for HTML format.
If ($Format -eq "HTML")
{
    "</div>" | Out-File -FilePath $File -Append
}

# Display the output file, in the application appropriate for the file extension.
Invoke-Expression $File