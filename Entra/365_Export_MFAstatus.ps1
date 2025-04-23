#=======================================================================================#
# Script: MFA Report Generator
# Version: 0.1
# Author: Robert Clarkson
# Email: robert.clarkson@neweratech.com
# Description: Exports MFA information for all licensed users in a Microsoft tenant.
#=======================================================================================#


# Install necessary modules only if not already present
$modulesToInstall = @(
    "MSOnline",
    "Microsoft.Graph.Beta"
	"Microsoft.Identity.Client"
)

foreach ($module in $modulesToInstall) {
    if (!(Get-InstalledModule -Name $module -ErrorAction SilentlyContinue)) {
        Write-Host "Installing $module module..." -ForegroundColor Yellow
        Install-Module $module -AllowClobber -Force -ErrorAction SilentlyContinue
        Write-Host "$module module installed."
    }
}

Write-Host "Connecting to MSOnline..."
Connect-MsolService
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "Policy.ReadWrite.AuthenticationMethod", "Organization.Read.All" -NoWelcome
$tenantName = (Get-MsolCompanyInformation).DisplayName


# Configs

# Change the path below if needed
$Csvfile = "C:\temp\MFAUsers_{0}_{1}.csv" -f $tenantName, (Get-Date -f yyyyMMddhhmm)
# Net to capture users Configure as required
$users = Get-MsolUser -All | Where-Object { $_.IsLicensed } | Select-Object DisplayName, UserPrincipalName, ObjectId

# Start Script
$Report = [System.Collections.Generic.List[Object]]::new()

# Loop through each user
$totalUsers = $users.Count
$counter = 0
foreach ($user in $users) {
    $counter++
    Write-Progress -Activity "Generating MFA Report" -Status "Processing user $($counter) of $totalUsers - $($user.DisplayName)..." -PercentComplete ([math]::Round(($counter / $totalUsers) * 100))
    $authMethods = Get-MgUserAuthenticationMethod -UserId $user.ObjectId

    $mfaDetails = [PSCustomObject]@{
        Status = "Disabled"
        AuthenticatorApp = $false
        PhoneAuthentication = $false
        Email = $false
        SoftwareOATH = $false
    }

    # Variables to store Authenticator device info - This is broken currently
    $authDevice = "Not Set"

    foreach ($method in $authMethods) {
        switch ($method.AdditionalProperties["@odata.type"]) {
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $mfaDetails.AuthenticatorApp = $true
                $mfaDetails.Status = "Enabled"

                if ($method.DisplayName) {
                    $authDevice = $method.DisplayName
                }
            }
            "#microsoft.graph.phoneAuthenticationMethod" {
                $mfaDetails.PhoneAuthentication = $true
                $mfaDetails.Status = "Enabled"
            }
            "#microsoft.graph.emailAuthenticationMethod" {
                $mfaDetails.Email = $true
                $mfaDetails.Status = "Enabled"
            }
            "#microsoft.graph.softwareOathAuthenticationMethod" {
                $mfaDetails.SoftwareOATH = $true
                $mfaDetails.Status = "Enabled"
            }
        }
    }

    # Get the users preferred MFA method
    $uri = "https://graph.microsoft.com/beta/users/$($user.ObjectId)/authentication/signInPreferences"
    $mfaPreferredMethod = Invoke-MgGraphRequest -uri $uri -Method GET

    $defaultMethodName = $null 
    foreach ($methodName in @(
        "Microsoft Authenticator - notification",
        "Phone - Call", 
        "Phone - SMS", 
        "Email Authentication", 
        "Software OATH Token"
    )) {
        if ($mfaPreferredMethod.userPreferredMethodForSecondaryAuthentication -eq $methodName) {
            switch ($methodName) {
                "Microsoft Authenticator - notification" {
                    # Check authentication mode for Authenticator app
                    switch ($authMethods | Where-Object { $_.authenticationMethod -eq "Microsoft Authenticator - notification" } | Select-Object -ExpandProperty AdditionalProperties.microsoftAuthenticatorAuthenticationMode) {
                        "push" { $defaultMethodName = "PhoneAppNotification"; break }
                        "oneWaySms" { $defaultMethodName = "SMS"; break }
                        "timeBasedOneTimePassword" { $defaultMethodName = "PhoneAppOTP"; break }
                        default { $defaultMethodName = "SoftwareOTP"; break }
                    }
                }
                default { $defaultMethodName = $methodName; break } 
            }
            break 
        } 
    }
    if (-not $defaultMethodName) {
        $defaultMethodName = $mfaPreferredMethod.userPreferredMethodForSecondaryAuthentication
    }

    $ReportLine = [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        MFAStatus = $mfaDetails.Status
        DefaultMFAMethod = $defaultMethodName
        PhoneAuthentication = $mfaDetails.PhoneAuthentication
        AuthenticatorApp = $mfaDetails.AuthenticatorApp
        Email = $mfaDetails.Email
        SoftwareOATH = $mfaDetails.SoftwareOATH
        "Authenticator Device" = $authDevice
    }

    $Report.Add($ReportLine)
}


# Closing
Write-Progress -Activity "Generating MFA Report" -Status "Completed" -PercentComplete 100
$Report | Export-Csv -Path $Csvfile -NoTypeInformation -Encoding UTF8
Write-Host "Script completed. Results exported to $Csvfile." -ForegroundColor Green