# TMS_Settings.ps1
# Description: Contains functions for managing application settings, primarily tariff margins.
#              Requires TMS_Helpers.ps1 and potentially TMS_Auth.ps1 (if re-auth needed).
#              This file should be dot-sourced by the main script.
# Usage: . .\TMS_Settings.ps1

# Assumes TMS_Helpers.ps1 functions (like Clear-HostAndDrawHeader, Load-KeysFromFolder, etc.) are available.
# Assumes necessary $script scoped variables (like folder paths) are available.

function Show-SettingsMenu {
   param(
       [Parameter(Mandatory=$true)]
       [hashtable]$UserProfile,
       [Parameter(Mandatory=$true)]
       [hashtable]$AllCentralKeys,
       [Parameter(Mandatory=$true)]
       [hashtable]$AllSAIAKeys,
       [Parameter(Mandatory=$true)]
       [hashtable]$AllRLKeys
   )

    $exitSettingsMenu = $false
    while (-not $exitSettingsMenu) {
        if (Get-Command Clear-HostAndDrawHeader -ErrorAction SilentlyContinue) {
            Clear-HostAndDrawHeader -Title "TMS Settings" -User $UserProfile.Username
        } else {
            Clear-Host
            Write-Host "--- TMS Settings (User: $($UserProfile.Username)) ---"
            Write-Warning "Clear-HostAndDrawHeader function not found."
        }

        Write-Host "`nOptions:" -ForegroundColor Yellow
        Write-Host "  1. Manage Central Transport Margins"
        Write-Host "  2. Manage SAIA Margins"
        Write-Host "  3. Manage R+L Carrier Margins"
        Write-Host "  B. Back to Main Menu"
        Write-Host "--------------------------------------" -ForegroundColor Blue
        $settingsChoice = Read-Host "Enter your choice"

        $currentScriptRoot = $null
        if ($PSScriptRoot) { $currentScriptRoot = $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { $currentScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent }
        else { Write-Error "FATAL (Show-SettingsMenu): Cannot determine script root directory."; return } 

        switch ($settingsChoice.ToUpper()) {
            '1' {
                Manage-TariffMargins -CarrierName "Central Transport" `
                                     -UserProfile $UserProfile `
                                     -AllKeys $AllCentralKeys `
                                     -KeysFolderPath (Join-Path $currentScriptRoot 'keys_central') 
            }
            '2' {
                Manage-TariffMargins -CarrierName "SAIA" `
                                     -UserProfile $UserProfile `
                                     -AllKeys $AllSAIAKeys `
                                     -KeysFolderPath (Join-Path $currentScriptRoot 'keys_saia') 
            }
            '3' { 
                Manage-TariffMargins -CarrierName "RL Carriers" `
                                     -UserProfile $UserProfile `
                                     -AllKeys $AllRLKeys `
                                     -KeysFolderPath (Join-Path $currentScriptRoot 'keys_rl') 
            }
            'B' { $exitSettingsMenu = $true }
            Default { Write-Warning "`nInvalid choice."; Read-Host "Press Enter..."; }
        } 
    } 
}


function Manage-TariffMargins {
    param(
        [Parameter(Mandatory=$true)][string]$CarrierName,
        [Parameter(Mandatory=$true)][hashtable]$UserProfile,
        [Parameter(Mandatory=$true)][hashtable]$AllKeys, 
        [Parameter(Mandatory=$true)][string]$KeysFolderPath 
    )

    $exitMarginMenu = $false
    while (-not $exitMarginMenu) {
        $title = "Manage $CarrierName Margins"
        if (Get-Command Clear-HostAndDrawHeader -ErrorAction SilentlyContinue) {
            Clear-HostAndDrawHeader -Title $title -User $UserProfile.Username
        } else {
            Clear-Host; Write-Host "--- $title (User: $($UserProfile.Username)) ---"
        }

    	$allowedKeyPropertyName = $null
        switch ($CarrierName) {
        	'Central Transport' { $allowedKeyPropertyName = 'AllowedCentralKeys' }
        	'SAIA'              { $allowedKeyPropertyName = 'AllowedSAIAKeys' }
        	'RL Carriers'       { $allowedKeyPropertyName = 'AllowedRLKeys' } 
        default {
            Write-Warning "Generating permission key name dynamically for unknown carrier '$CarrierName'."
            $allowedKeyPropertyName = "Allowed" + ($CarrierName -replace '[^a-zA-Z0-9]', '') + "Keys"
        }
    }
        $permittedKeys = @{}
        if ($UserProfile.ContainsKey($allowedKeyPropertyName)) {
             if (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue) {
                $permittedKeys = Get-PermittedKeys -AllKeys $AllKeys -AllowedKeyNames $UserProfile[$allowedKeyPropertyName]
             } else {
                 Write-Error "Get-PermittedKeys function not found. Cannot list permitted keys."
                 Read-Host "Press Enter to return..."; $exitMarginMenu = $true; continue
             }
        } else {
            Write-Warning "User profile does not contain the expected permission key '$allowedKeyPropertyName'."
        }


        Write-Host "`nPermitted $CarrierName Tariffs/Accounts:" -ForegroundColor Yellow
        if ($permittedKeys.Count -gt 0) {
            $keyNames = @($permittedKeys.Keys | Sort-Object)
            for ($i = 0; $i -lt $keyNames.Count; $i++) {
                $keyName = $keyNames[$i]
                $keyData = $permittedKeys[$keyName]
                $currentMargin = "N/A" 
                if ($keyData -is [hashtable] -and $keyData.ContainsKey('MarginPercent')) {
                    if ($keyData['MarginPercent'] -as [double] -ne $null) {
                        $currentMargin = "{0:N1}%" -f ([double]$keyData['MarginPercent'])
                    } else { $currentMargin = "Invalid!" }
                }
                Write-Host (" [{0,2}] : {1} (Current Margin: {2})" -f ($i + 1), $keyName, $currentMargin)
            }
        } else {
            Write-Host "  No $CarrierName tariffs permitted for this user." -ForegroundColor Gray
        }
        Write-Host "--------------------------------------" -ForegroundColor Blue
        Write-Host "Options:" -ForegroundColor Yellow
        Write-Host "  S. Set Margin for a Tariff (Enter Number)"
        Write-Host "  B. Back to Settings Menu"
        Write-Host "--------------------------------------" -ForegroundColor Blue
        $marginChoice = Read-Host "Enter your choice"

        switch ($marginChoice.ToUpper()) {
            'S' { 
                 if ($permittedKeys.Count -eq 0) { Write-Warning "No permitted tariffs to set margins for."; Read-Host "..."; continue }
                 $idxInput = Read-Host "Enter tariff number to set margin for"
                 if ($idxInput -match '^\d+$') {
                    $idx = [int]$idxInput - 1
                    if ($idx -ge 0 -and $idx -lt $keyNames.Count) {
                        $tariffToUpdate = $keyNames[$idx]
                        Set-SingleTariffMargin -TariffName $tariffToUpdate `
                                               -AllKeysHashtable $AllKeys `
                                               -KeysFolderPath $KeysFolderPath 
                    } else { Write-Warning "Invalid tariff number." }
                 } else { Write-Warning "Invalid input." }
                 Read-Host "`nPress Enter..." 
            }
            'B' { $exitMarginMenu = $true }
            Default { Write-Warning "Invalid choice."; Read-Host "Press Enter..."; }
        } 
    } 
}


function Set-SingleTariffMargin {
    param(
        [Parameter(Mandatory=$true)][string]$TariffName,
        [Parameter(Mandatory=$true)][hashtable]$AllKeysHashtable, 
        [Parameter(Mandatory=$true)][string]$KeysFolderPath
    )

    $filePath = Join-Path -Path $KeysFolderPath -ChildPath "$TariffName.txt"
    if (-not (Test-Path $filePath -PathType Leaf)) { Write-Error "Key file not found '$filePath'. Cannot set margin."; return }

    $newMarginPercent = $null
    while ($newMarginPercent -eq $null) {
        $input = Read-Host "Enter NEW Margin % for '$TariffName' (e.g., 15.5 for 15.5%, 0-99.9)"
        try {
            $marginValue = [double]$input
            if ($marginValue -ge 0 -and $marginValue -lt 100) {
                $newMarginPercent = [math]::Round($marginValue, 1) 
            } else { Write-Warning "Margin must be between 0 and 99.9." }
        } catch { Write-Warning "Invalid number format." }
    }

    Write-Host "Attempting to update '$TariffName' margin to $newMarginPercent%..." -ForegroundColor Gray

    try {
        $fileContent = Get-Content -Path $filePath -Raw -ErrorAction Stop
        $marginLinePattern = '(?im)^\s*MarginPercent\s*=.*$' 
        $newLine = "MarginPercent=$newMarginPercent"

        if ($fileContent -match $marginLinePattern) {
            $updatedContent = $fileContent -replace $marginLinePattern, $newLine
            Write-Verbose "Found and replaced existing MarginPercent line."
        } else {
            Write-Warning "MarginPercent line not found in '$filePath'. Appending."
            $updatedContent = $fileContent.TrimEnd() + "`r`n" + $newLine
        }

        Set-Content -Path $filePath -Value $updatedContent -Encoding UTF8 -Force -ErrorAction Stop

        if ($AllKeysHashtable -ne $null -and $AllKeysHashtable.ContainsKey($TariffName)) {
            if ($AllKeysHashtable[$TariffName] -is [hashtable]) {
                 $AllKeysHashtable[$TariffName].MarginPercent = $newMarginPercent
                 Write-Verbose "In-memory margin updated for '$TariffName'."
            } else {
                 Write-Warning "Value for '$TariffName' in memory is not a hashtable. Cannot update in-memory margin."
            }
        } else {
             Write-Warning "Could not find '$TariffName' in the provided in-memory hashtable. Restart tool to see changes."
        }

        Write-Host " -> Success: '$TariffName' margin updated to $newMarginPercent%." -ForegroundColor Green

    } catch {
        Write-Error " -> Failed update for '$TariffName': $($_.Exception.Message)"
    }
}

function Update-TariffMargin {
    param(
        [Parameter(Mandatory=$true)][string]$TariffName,
        [Parameter(Mandatory=$true)][hashtable]$AllKeysHashtable, 
        [Parameter(Mandatory=$true)][string]$KeysFolderPath,
        [Parameter(Mandatory=$true)][double]$NewMarginPercent 
    )
    if ($NewMarginPercent -lt 0 -or $NewMarginPercent -ge 100) {
        Write-Error "Invalid margin value passed to Update-TariffMargin: ${NewMarginPercent}. Must be 0-99.9."
        return $false 
    }

    $filePath = Join-Path -Path $KeysFolderPath -ChildPath "${TariffName}.txt"
    if (-not (Test-Path $filePath -PathType Leaf)) { Write-Error "Key file not found '${filePath}'. Cannot update margin."; return $false }

    Write-Verbose "Attempting to update '${TariffName}' margin to ${NewMarginPercent}% in file '${filePath}'..."

    try {
        $fileContent = Get-Content -Path $filePath -Raw -ErrorAction Stop
        $marginLinePattern = '(?im)^\s*MarginPercent\s*=.*$' 
        $newMarginString = "{0:F1}" -f $NewMarginPercent 
        $newLine = "MarginPercent=${newMarginString}" 

        if ($fileContent -match $marginLinePattern) {
            $updatedContent = $fileContent -replace $marginLinePattern, $newLine
            Write-Verbose "Found and replaced existing MarginPercent line."
        } else {
            Write-Warning "MarginPercent line not found in '${filePath}'. Appending."
            if (-not ($fileContent.EndsWith("`r`n")) -and -not ($fileContent.EndsWith("`n"))) {
                 $updatedContent = $fileContent + "`r`n" + $newLine
            } else {
                 $updatedContent = $fileContent + $newLine
            }
        }

        Set-Content -Path $filePath -Value $updatedContent -Encoding UTF8 -Force -ErrorAction Stop

        if ($AllKeysHashtable -ne $null -and $AllKeysHashtable.ContainsKey($TariffName)) {
            if ($AllKeysHashtable[$TariffName] -is [hashtable]) {
                 $AllKeysHashtable[$TariffName].MarginPercent = $NewMarginPercent
                 Write-Verbose "In-memory margin updated for '${TariffName}'."
            } else {
                 Write-Warning "Value for '${TariffName}' in memory is not a hashtable. Cannot update in-memory margin."
            }
        } else {
             Write-Warning "Could not find '${TariffName}' in the provided in-memory hashtable. Restart tool or re-load keys to see changes fully reflected in memory."
        }
        Write-Verbose "Successfully updated '${TariffName}' margin to ${NewMarginPercent}%."
        return $true 
    } catch {
        Write-Error "Failed update for '${TariffName}': $($_.Exception.Message)" 
        return $false 
    }
}

Write-Verbose "TMS Settings Functions loaded."
