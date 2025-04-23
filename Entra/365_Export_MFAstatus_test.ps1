#=======================================================================================#
# Script: MFA Report Generator
# Version: 0.1
# Author: Robert Clarkson
# Email: robert.clarkson@neweratech.com
# Description: Exports MFA information for all licensed users in a Microsoft tenant.
#=======================================================================================#

# --- Script Configuration ---
$Csvfile = "C:\temp\MFAUsers_{0}_{1}.csv" 

# --- Functions ---

function Install-RequiredModules {
    # Installs required modules if they are not already present
    $modulesToInstall = @(
        "MSOnline",
        "Microsoft.Graph.Beta",
        "Microsoft.Identity.Client"
    )

    foreach ($module in $modulesToInstall) {
        if (!(Get-InstalledModule -Name $module -ErrorAction SilentlyContinue)) {
            Write-Host "Installing $module module..." -ForegroundColor Yellow
            Install-Module $module -AllowClobber -Force -ErrorAction SilentlyContinue
            Write-Host "$module module installed."
        }
    }
}

function Connect-ToGraph {
    # Connects to Microsoft Graph Beta API with the required scopes
    Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "Policy.ReadWrite.AuthenticationMethod", "Organization.Read.All" -NoWelcome 
}

function Get-MFAReport {
    # Retrieves MFA information for all licensed users
    $tenantName = (Get-MgOrganization).displayName 
    $users = Get-MgUser -All | Where-Object { $_.assignedLicenses } | Select-Object DisplayName, UserPrincipalName, Id 

    $Report = [System.Collections.Generic.List[Object]]::new()

    foreach ($user in $users) {
        Write-Progress -Activity "Generating MFA Report" -Status "Processing user $($user.DisplayName)..." 
        
        # ... (rest of the code to get MFA details for each user) ...

        $ReportLine = [PSCustomObject]@{
            # ... (properties for the report) ...
        }

        $Report.Add($ReportLine)
    }

    return $Report
}

# --- End Functions ---

# --- Start Script ---

# Install required modules
Install-RequiredModules

# Connect to Microsoft Graph Beta API
Connect-ToGraph

# Generate the MFA report
$mfaReport = Get-MFAReport

# Set the CSV file path with the tenant name and timestamp
$Csvfile = $Csvfile -f $tenantName, (Get-Date -f yyyyMMddhhmm)

# Export the report to CSV
$mfaReport | Export-Csv -Path $Csvfile -NoTypeInformation -Encoding UTF8
Write-Host "Script completed. Results exported to $Csvfile." -ForegroundColor Green