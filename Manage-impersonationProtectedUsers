# -----------------------------------------------
# Helper: prompt user to select an Anti-Phish policy
# -----------------------------------------------
function Select-AntiPhishPolicy {
    $policies = Get-AntiPhishPolicy
    Write-Host ""
    Write-Host "Anti-Phish Policies:"
    for ($i = 0; $i -lt $policies.Count; $i++) {
        Write-Host "  $($i+1). $($policies[$i].Name)"
    }
    $index = Read-Host "Enter the number of the policy"
    if ($index -lt 1 -or $index -gt $policies.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        return $null
    }
    return $policies[$index - 1]
}

# -----------------------------------------------
# Action 1: Import users from CSV
# -----------------------------------------------
function Import-UsersToPolicy {
    $policy = Select-AntiPhishPolicy
    if (-not $policy) { return }

    # Ensure targeted user protection is enabled
    if (-not $policy.EnableTargetedUserProtection) {
        $answer = ""
        while ($answer -notin @("Y","y","N","n")) {
            $answer = Read-Host "Targeted user protection is not enabled for this policy. Enable it? (Y/N)"
        }
        if ($answer -in @("Y","y")) {
            Set-AntiPhishPolicy -Identity $policy.Identity -EnableTargetedUserProtection $true
            Write-Host "Targeted user protection enabled." -ForegroundColor Green
        } else {
            Write-Host "Import cancelled - targeted user protection must be enabled." -ForegroundColor Yellow
            return
        }
    }

    $csvPath = Read-Host "Enter the path to the CSV file"
    if (-not (Test-Path $csvPath)) {
        Write-Host "File not found: $csvPath" -ForegroundColor Red
        return
    }

    $users = Import-Csv $csvPath
    if (-not ($users[0].PSObject.Properties.Name -contains 'username;upn')) {
        Write-Host "Invalid CSV format. Expected a 'username;upn' column." -ForegroundColor Red
        return
    }

    # Get existing entries to detect duplicates
    $existingEntries = (Get-AntiPhishPolicy -Identity $policy.Identity).TargetedUsersToProtect
    if (-not $existingEntries) { $existingEntries = @() }

    $results = @()
    $toAdd   = @()

    foreach ($user in $users) {
        $entry  = $user.'username;upn'
        $parts  = $entry -split ';'
        $name   = $parts[0]
        $upn    = if ($parts.Count -gt 1) { $parts[1] } else { $parts[0] }

        if ($existingEntries -contains $entry) {
            Write-Host "  [SKIPPED] $name ($upn) - already in policy" -ForegroundColor Yellow
            $results += [PSCustomObject]@{ DisplayName=$name; UPN=$upn; Status="Already Exists" }
        } else {
            $toAdd   += [PSCustomObject]@{ DisplayName=$name; UPN=$upn; Entry=$entry }
        }
    }

    if ($toAdd.Count -eq 0) {
        Write-Host "All users already exist in the policy. Nothing to import." -ForegroundColor Yellow
    } else {
        $newEntries  = $toAdd | ForEach-Object { $_.Entry }
        $combinedList = $existingEntries + $newEntries
        Set-AntiPhishPolicy -Identity $policy.Identity -TargetedUsersToProtect $combinedList

        # Verify each new entry was actually saved
        $savedEntries = (Get-AntiPhishPolicy -Identity $policy.Identity).TargetedUsersToProtect
        foreach ($item in $toAdd) {
            if ($savedEntries -contains $item.Entry) {
                Write-Host "  [OK] $($item.DisplayName) ($($item.UPN))" -ForegroundColor Green
                $results += [PSCustomObject]@{ DisplayName=$item.DisplayName; UPN=$item.UPN; Status="Success" }
            } else {
                Write-Host "  [FAILED] $($item.DisplayName) ($($item.UPN))" -ForegroundColor Red
                $results += [PSCustomObject]@{ DisplayName=$item.DisplayName; UPN=$item.UPN; Status="Failed" }
            }
        }
    }

    # Save results CSV alongside the input file
    $csvDir      = Split-Path $csvPath -Parent
    $csvBaseName = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
    $resultsPath = Join-Path $csvDir "$csvBaseName-import-results.csv"
    $results | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Results saved to: $resultsPath"
    $successCount = ($results | Where-Object { $_.Status -eq "Success" }).Count
    $skipCount    = ($results | Where-Object { $_.Status -eq "Already Exists" }).Count
    $failCount    = ($results | Where-Object { $_.Status -eq "Failed" }).Count
    Write-Host "Summary: $successCount added, $skipCount already existed, $failCount failed." -ForegroundColor Cyan
}

# -----------------------------------------------
# Action 2: Delete all users from a policy
# -----------------------------------------------
function Remove-AllUsersFromPolicy {
    $policy = Select-AntiPhishPolicy
    if (-not $policy) { return }

    $currentUsers = (Get-AntiPhishPolicy -Identity $policy.Identity).TargetedUsersToProtect
    if (-not $currentUsers -or $currentUsers.Count -eq 0) {
        Write-Host "No users are currently protected under '$($policy.Name)'." -ForegroundColor Yellow
        return
    }

    Write-Host "This will remove $($currentUsers.Count) user(s) from '$($policy.Name)'." -ForegroundColor Yellow
    $confirm = Read-Host "Are you sure? (Y/N)"
    if ($confirm -notin @("Y","y")) {
        Write-Host "Deletion cancelled." -ForegroundColor Yellow
        return
    }

    Set-AntiPhishPolicy -Identity $policy.Identity -TargetedUsersToProtect @()
    Write-Host "All users removed from impersonation protection in '$($policy.Name)'." -ForegroundColor Green
}

# -----------------------------------------------
# Connect to Exchange Online
# -----------------------------------------------
Connect-ExchangeOnline

# -----------------------------------------------
# Main menu loop
# -----------------------------------------------
do {
    Write-Host ""
    Write-Host "=================================="
    Write-Host "  Anti-Phish Impersonation Menu"
    Write-Host "=================================="
    Write-Host "  1. Import users from CSV"
    Write-Host "  2. Delete all users from a policy"
    Write-Host "  3. Exit"
    Write-Host ""

    $choice = Read-Host "Select an option (1-3)"

    switch ($choice) {
        "1"     { Import-UsersToPolicy }
        "2"     { Remove-AllUsersFromPolicy }
        "3"     { Write-Host "Exiting..." -ForegroundColor Cyan }
        default { Write-Host "Invalid option. Please enter 1, 2 or 3." -ForegroundColor Red }
    }
} while ($choice -ne "3")

# Gracefully disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Cyan
