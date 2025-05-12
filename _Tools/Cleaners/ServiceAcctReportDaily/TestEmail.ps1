# =========================================================================
# TITLE: TestEmail_20250211_0800
# AUTHOR: System Administrator
# DATE: February 11, 2025
# VERSION: 1.0
# =========================================================================

# Email configuration
$emailAddress = "jromaine@davidyurman.com"
$smtpServer = "smtp.office365.com"
$port = 587

# If credentials aren't already stored, create them
$credPath = "C:\Scripts\_Automation\ServiceAcctReportDaily\email.xml"
if (-not (Test-Path $credPath)) {
    Write-Host "No stored credentials found. Please enter your email credentials."
    $credential = Get-Credential -Message "Enter credentials for $emailAddress"
    $credential | Export-Clixml -Path $credPath
}

$credential = Import-Clixml -Path $credPath

# Create test email parameters
$emailParams = @{
    From = $emailAddress
    To = $emailAddress
    Subject = "Test Email - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    Body = "This is a test email to verify O365 email connectivity.`n`nSent from: $env:COMPUTERNAME`nTime: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    SmtpServer = $smtpServer
    Port = $port
    UseSsl = $true
    Credential = $credential
}

# Send test email
try {
    Send-MailMessage @emailParams
    Write-Host "Test email sent successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to send test email. Error: $_" -ForegroundColor Red
}