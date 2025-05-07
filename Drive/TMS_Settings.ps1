# TMS_Settings.ps1
# Description: Contains functions for managing TMS settings, primarily tariff margins.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first (by main entry script).
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# Assumes helper functions like Clear-HostAndDrawHeader, Load-KeysFromFolder, Select-SingleKeyEntry are available.
# Assumes global variables like $Global:allCentralKeys, $Global:allSAIAKeys, $Global:allRLKeys, $Global:allAverittKeys are populated.
# Assumes script-scoped folder paths like $script:centralKeysFolderPath, etc., are set by the main entry script.

function Show-SettingsMenu {
    param(
        [Parameter(Mandatory=$true)][hashtable]$UserProfile, # Logged-in user's profile
        # Pass all loaded keys for all carriers
        [Parameter(Mandatory=$true)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllRLKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllAverittKeys # <<< AVERITT ADDED >>>
    )

    $exitSettingsMenu = $false
    while (-not $exitSettingsMenu) {
        Clear-HostAndDrawHeader -Title "TMS Settings - Manage Margins" -User $UserProfile.Username
        Write-Host "Select a carrier to manage tariff margins or 'B' to go back:" -ForegroundColor Yellow
        Write-Host "  1. Central Transport"
        Write-Host "  2. SAIA"
        Write-Host "  3. R+L Carriers"
        Write-Host "  4. Averitt" # <<< AVERITT ADDED >>>
        Write-Host "  B. Back to Main Menu"
        Write-Host "--------------------------------------" -ForegroundColor Blue
        $carrierChoiceInput = Read-Host "Enter your choice"

        $selectedCarrierName = $null
        $selectedAllCarrierKeys = $null
        $selectedKeysFolderPath = $null

        switch ($carrierChoiceInput.ToUpper()) {
            '1' {
                $selectedCarrierName = "Central Transport"
                $selectedAllCarrierKeys = $AllCentralKeys
                $selectedKeysFolderPath = $script:centralKeysFolderPath # Relies on this being set by TMS_GUI.ps1
            }
            '2' {
                $selectedCarrierName = "SAIA"
                $selectedAllCarrierKeys = $AllSAIAKeys
                $selectedKeysFolderPath = $script:saiaKeysFolderPath
            }
            '3' {
                $selectedCarrierName = "R+L Carriers"
                $selectedAllCarrierKeys = $AllRLKeys
                $selectedKeysFolderPath = $script:rlKeysFolderPath
            }
            '4' { # <<< AVERITT ADDED >>>
                $selectedCarrierName = "Averitt"
                $selectedAllCarrierKeys = $AllAverittKeys
                $selectedKeysFolderPath = $script:averittKeysFolderPath # Relies on this being set
            }
            'B' { $exitSettingsMenu = $true; continue }
            Default { Write-Warning "Invalid carrier selection."; Read-Host "Press Enter..."; continue }
        }

        if ($null -eq $selectedAllCarrierKeys -or $selectedAllCarrierKeys.Count -eq 0) {
            Write-Warning "No keys loaded for $selectedCarrierName. Cannot manage margins."
            Read-Host "Press Enter..."
            continue
        }
        if ([string]::IsNullOrWhiteSpace($selectedKeysFolderPath) -or -not (Test-Path $selectedKeysFolderPath -PathType Container)) {
            Write-Error "FATAL: Keys folder path for $selectedCarrierName is not correctly set or accessible: '$selectedKeysFolderPath'"
            Read-Host "Press Enter..."
            continue
        }

        Manage-TariffMargins -CarrierName $selectedCarrierName -AllKeysForCarrier $selectedAllCarrierKeys -KeysFolderPath $selectedKeysFolderPath
    }
}

function Manage-TariffMargins {
    param(
        [Parameter(Mandatory=$true)][string]$CarrierName,
        [Parameter(Mandatory=$true)][hashtable]$AllKeysForCarrier, # All loaded keys for the selected carrier
        [Parameter(Mandatory=$true)][string]$KeysFolderPath      # Full path to the key files for this carrier
    )

    $exitManageMargins = $false
    while (-not $exitManageMargins) {
        Clear-HostAndDrawHeader -Title "Manage Margins for $CarrierName"

        if ($AllKeysForCarrier.Count -eq 0) {
            Write-Warning "No key files found for $CarrierName in '$KeysFolderPath'."
            Read-Host "Press Enter to return..."
            return
        }

        # Display current margins
        Write-Host "Current Margins for $CarrierName Tariffs/Accounts:" -ForegroundColor Cyan
        # Sort by the 'Name' property of the key's hashtable value for consistent display
        $AllKeysForCarrier.GetEnumerator() | Sort-Object {$_.Value.Name} | ForEach-Object {
            $keyDetails = $_.Value
            $keyDisplayName = if ($keyDetails.ContainsKey('Name')) { $keyDetails.Name } else { $_.Key } # Fallback to filename
            $marginDisplay = if ($keyDetails.ContainsKey('MarginPercent')) { "$($keyDetails.MarginPercent)%" } else { "Not Set (Uses Default)" }
            Write-Host ("  {0,-35}: {1}" -f $keyDisplayName, $marginDisplay)
        }
        Write-Host "----------------------------------------------------"

        Write-Host "`nOptions:"
        Write-Host "  1. Update Margin for a Tariff/Account"
        Write-Host "  B. Back to Carrier Selection"
        $actionChoice = Read-Host "Enter your choice"

        switch ($actionChoice.ToUpper()) {
            '1' {
                $selectedKeyDetails = Select-SingleKeyEntry -AvailableKeys $AllKeysForCarrier -PromptMessage "Select Tariff/Account to Update Margin for $CarrierName"
                if ($null -eq $selectedKeyDetails) { continue } # User cancelled

                $keyFileName = $selectedKeyDetails.TariffFileName # This was stored when keys were loaded
                if ([string]::IsNullOrWhiteSpace($keyFileName)) {
                    Write-Warning "Could not determine original filename for key '$($selectedKeyDetails.Name)'. Cannot update margin."
                    Read-Host "Press Enter..."; continue
                }
                $fullKeyFilePath = Join-Path -Path $KeysFolderPath -ChildPath "$($keyFileName).txt"

                if (-not (Test-Path $fullKeyFilePath -PathType Leaf)) {
                    Write-Warning "Key file '$fullKeyFilePath' not found. Cannot update margin."
                    Read-Host "Press Enter..."; continue
                }

                $currentMargin = if ($selectedKeyDetails.ContainsKey('MarginPercent')) { $selectedKeyDetails.MarginPercent } else { "Not Set (Using Default)" }
                Write-Host "Current margin for '$($selectedKeyDetails.Name)': $currentMargin"
                
                $newMarginInput = $null
                $newMarginDouble = $null
                while ($newMarginDouble -eq $null) {
                    $newMarginInput = Read-Host "Enter new Margin Percentage (e.g., 12.5 for 12.5%, or 'C' to cancel)"
                    if ($newMarginInput.ToUpper() -eq 'C') { break }
                    try {
                        $tempMargin = [double]$newMarginInput
                        if ($tempMargin -ge 0 -and $tempMargin -lt 100) { # Margin typically 0-99.9%
                            $newMarginDouble = $tempMargin
                        } else { Write-Warning "Margin must be between 0 and 99.99." }
                    } catch { Write-Warning "Invalid number format for margin." }
                }

                if ($newMarginInput.ToUpper() -eq 'C') { Write-Host "Margin update cancelled."; continue }
                if ($newMarginDouble -eq $null) { Write-Warning "Margin update failed due to invalid input."; continue }

                # Update the key file content
                try {
                    $fileContent = Get-Content -Path $fullKeyFilePath -Raw -ErrorAction Stop
                    $marginLinePattern = '(?m)^\s*MarginPercent\s*=\s*.*$' # Regex to find existing MarginPercent line

                    if ($fileContent -match $marginLinePattern) {
                        $newFileContent = $fileContent -replace $marginLinePattern, "MarginPercent=$newMarginDouble"
                    } else {
                        # Append if not found (ensure it's on a new line)
                        $newFileContent = $fileContent.TrimEnd() + [System.Environment]::NewLine + "MarginPercent=$newMarginDouble"
                    }
                    Set-Content -Path $fullKeyFilePath -Value $newFileContent -Encoding UTF8 -ErrorAction Stop
                    Write-Host "Margin for '$($selectedKeyDetails.Name)' updated to $newMarginDouble% in file '$fullKeyFilePath'." -ForegroundColor Green
                    
                    # IMPORTANT: Reload keys for this carrier so the change is reflected immediately in the global variable
                    Write-Host "Reloading $CarrierName keys to reflect changes..." -ForegroundColor DarkYellow
                    $reloadedKeys = Load-KeysFromFolder -KeysFolderPath $KeysFolderPath -CarrierName $CarrierName
                    switch ($CarrierName) {
                        "Central Transport" { $Global:allCentralKeys = $reloadedKeys }
                        "SAIA"              { $Global:allSAIAKeys = $reloadedKeys }
                        "R+L Carriers"      { $Global:allRLKeys = $reloadedKeys }
                        "Averitt"           { $Global:allAverittKeys = $reloadedKeys } # <<< AVERITT ADDED >>>
                    }
                    # Update the local $AllKeysForCarrier for the current loop iteration
                    $AllKeysForCarrier = $reloadedKeys

                } catch {
                    Write-Error "Failed to update margin in key file '$fullKeyFilePath': $($_.Exception.Message)"
                }
                Read-Host "Press Enter..."
            }
            'B' { $exitManageMargins = $true }
            Default { Write-Warning "Invalid option."; Read-Host "Press Enter..." }
        }
    }
}

Write-Verbose "TMS Settings Functions loaded."
