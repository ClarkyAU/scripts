# Import Exchange Online PowerShell Module
Import-Module ExchangeOnlineManagement

# Connect To Exchange Online (will trigger interactive login prompt)
Connect-ExchangeOnline

Write-Host "`nFetching all mailboxes (this may take a while)..."
$allMailboxes = @{}
try {
    Get-Mailbox -ResultSize Unlimited | Select-Object Identity, PrimarySmtpAddress | ForEach-Object {
        if ($_.PrimarySmtpAddress) {
            $allMailboxes[$_.Identity.ToString()] = $_.PrimarySmtpAddress.ToString()
        } else {
            Write-Warning "Mailbox with Identity $($_.Identity) has no PrimarySmtpAddress. Skipping."
        }
    }
    Write-Host "Done fetching $($allMailboxes.Count) mailboxes with Primary SMTP Addresses." -ForegroundColor Green
}
catch {
    Write-Error "Failed to fetch mailboxes: $($_.Exception.Message)"
}

$Username = Read-Host "`nEnter the primary email address of the user you want to audit"

# --- Show Mailboxes with Full Access Permissions ---

Write-Host "`nChecking Full Access permissions for $Username..."
$fullAccessPermissions = Get-Mailbox -ResultSize Unlimited |
    Get-MailboxPermission -User $Username -ErrorAction SilentlyContinue |
    Where-Object { $_.AccessRights -contains 'FullAccess' -and $_.IsInherited -eq $false }

if ($fullAccessPermissions) {
    Write-Host "`n$Username has Full Access permissions to the following mailboxes:" -ForegroundColor Yellow
    $fullAccessPermissions | ForEach-Object {
        $identityKey = $_.Identity.ToString()
        if ($allMailboxes.ContainsKey($identityKey)) {
            [PSCustomObject]@{
                PrimarySmtpAddress = $allMailboxes[$identityKey].ToLower()
            }
        } else {
            Write-Warning "Could not find details for mailbox identity '$identityKey' in pre-fetched list (Full Access)."
        }
    } | Sort-Object PrimarySmtpAddress | Format-Table -AutoSize -HideTableHeaders
} else {
}


# --- Show Mailboxes with Send As Permissions ---

Write-Host "`nChecking Send As permissions for $Username..."
$sendAsPermissions = Get-Mailbox -ResultSize Unlimited |
    Get-RecipientPermission -Trustee $Username -ErrorAction SilentlyContinue |
    Where-Object { $_.AccessRights -contains 'SendAs' -and $_.IsInherited -eq $false }

if ($sendAsPermissions) {
    Write-Host "`n$Username has Send As permissions to the following mailboxes:" -ForegroundColor Yellow
    $sendAsPermissions | ForEach-Object {
        $identityKey = $_.Identity.ToString()
        if ($allMailboxes.ContainsKey($identityKey)) {
            [PSCustomObject]@{
                PrimarySmtpAddress = $allMailboxes[$identityKey].ToLower()
            }
        } else {
            Write-Warning "Could not find details for mailbox identity '$identityKey' in pre-fetched list (Send As)."
        }
    } | Sort-Object PrimarySmtpAddress | Format-Table -AutoSize -HideTableHeaders
} else {
}


# Disconnect from Exchange Online (with user confirmation prompt)
Write-Host "`nDisconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm