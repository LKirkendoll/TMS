# TMS_Reports.ps1
# Description: Contains functions for generating various reports and analyses,
#              refactored to accept parameters for GUI use. Corrected parameter mismatch and added Averitt processing.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first.
#              This file should be dot-sourced by the main script.

# Assumes TMS_Helpers.ps1 functions are available.
# Assumes TMS_Config.ps1 variables are available.
# Assumes TMS_Carrier_*.ps1 functions (GUI versions) are available.

# --- Report Management (Keep for now, but GUI might use different approach) ---
function Manage-UserReports {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserReportsFolder
    )
    # This function retains console interaction and might not be directly used by GUI.
    # GUI will likely implement its own report Browse/management.
    if (-not (Test-Path $UserReportsFolder)) { Write-Warning "No reports folder found at '$UserReportsFolder'."; Read-Host "..."; return }
    $exitReportMenu = $false
    while (-not $exitReportMenu) {
        if (Get-Command Clear-HostAndDrawHeader -ErrorAction SilentlyContinue) { Clear-HostAndDrawHeader -Title "Manage My Reports" -User (Split-Path $UserReportsFolder -Leaf) }
        else { Clear-Host; Write-Host "--- Manage My Reports (User: $(Split-Path $UserReportsFolder -Leaf)) ---"; Write-Warning "Clear-HostAndDrawHeader function not found." }
        $reportFiles = Get-ChildItem -Path $UserReportsFolder -Recurse -Filter "*.txt" -File | Sort-Object LastWriteTime -Descending
        Write-Host "`nAvailable Reports in '$UserReportsFolder':" -ForegroundColor Yellow
        if ($reportFiles.Count -gt 0) { for ($i = 0; $i -lt $reportFiles.Count; $i++) { $relativePath = $reportFiles[$i].FullName.Substring($UserReportsFolder.Length).TrimStart('\/'); Write-Host (" [{0,2}] : {1} ({2:yyyy-MM-dd HH:mm})" -f ($i + 1), $relativePath, $reportFiles[$i].LastWriteTime) } }
        else { Write-Host "  No reports found." -ForegroundColor Gray }
        Write-Host "--------------------------------------" -ForegroundColor Blue; Write-Host "Options:" -ForegroundColor Yellow; Write-Host "  O. Open Report (Number)"; Write-Host "  D. Delete Report (Number)"; Write-Host "  X. Delete ALL Reports (Confirm)"; Write-Host "  E. Open Reports Folder"; Write-Host "  B. Back"; Write-Host "--------------------------------------" -ForegroundColor Blue
        $reportChoice = Read-Host "Enter your choice"
        switch ($reportChoice.ToUpper()) {
            'O' { if ($reportFiles.Count -eq 0) { Write-Warning "No reports."; Read-Host "..."; continue }; $idxInput = Read-Host "Report number to open"; if ($idxInput -match '^\d+$') { $idx = [int]$idxInput - 1; if ($idx -ge 0 -and $idx -lt $reportFiles.Count) { Write-Host "Opening '$($reportFiles[$idx].Name)'..."; if (Get-Command Open-FileExplorer -ErrorAction SilentlyContinue) { Open-FileExplorer -Path $reportFiles[$idx].FullName } else { Write-Warning "Open-FileExplorer not found." } } else { Write-Warning "Invalid number." } } else { Write-Warning "Invalid input." }; Read-Host "`nPress Enter..." }
            'D' { if ($reportFiles.Count -eq 0) { Write-Warning "No reports."; Read-Host "..."; continue }; $idxInput = Read-Host "Report number to DELETE"; if ($idxInput -match '^\d+$') { $idx = [int]$idxInput - 1; if ($idx -ge 0 -and $idx -lt $reportFiles.Count) { $fileToDelete = $reportFiles[$idx]; $confirm = Read-Host "DELETE '$($fileToDelete.Name)'? (Y/N)"; if ($confirm -match '^[Yy]$') { try { Remove-Item -Path $fileToDelete.FullName -Force -ErrorAction Stop; Write-Host "'$($fileToDelete.Name)' deleted." -ForegroundColor Green } catch { Write-Error "Failed delete: $($_.Exception.Message)" } } else { Write-Host "Cancelled." } } else { Write-Warning "Invalid number." } } else { Write-Warning "Invalid input." }; Read-Host "`nPress Enter..." }
            'X' { if ($reportFiles.Count -eq 0) { Write-Warning "No reports."; Read-Host "..."; continue }; $confirm = Read-Host "DELETE ALL $($reportFiles.Count) REPORTS? Type 'YES' to confirm."; if ($confirm -eq 'YES') { Write-Host "Deleting..." -ForegroundColor Red; $deletedCount = 0; $failCount = 0; foreach ($file in $reportFiles) { try { Remove-Item -Path $file.FullName -Force -ErrorAction Stop; $deletedCount++ } catch { Write-Error "Failed delete '$($file.Name)': $($_.Exception.Message)"; $failCount++ } }; Write-Host "Deleted ${deletedCount}. Failed ${failCount}." -ForegroundColor Yellow } else { Write-Host "Cancelled." }; Read-Host "`nPress Enter..." }
            'E' { if (Get-Command Open-FileExplorer -ErrorAction SilentlyContinue) { Open-FileExplorer -Path $UserReportsFolder } else { Write-Warning "Open-FileExplorer not found." }; Read-Host "`nPress Enter..." }
            'B' { $exitReportMenu = $true }
            Default { Write-Warning "Invalid choice."; Read-Host "Press Enter..."; }
        }
    }
}


# --- Cross-Carrier Analysis Functions (Refactored for GUI) ---

function Run-CrossCarrierASPAnalysisGUI {
    # GUI VERSION: Calculates required margin for each carrier based on a desired ASP.
    # Optionally applies calculated margins (requires confirmation).
    param(
        [Parameter(Mandatory)][hashtable]$BrokerProfile,
        [Parameter(Mandatory)][hashtable]$SelectedCustomerProfile,
        [Parameter(Mandatory)][string]$ReportsBaseFolder,
        [Parameter(Mandatory)][string]$UserReportsFolder,
        [Parameter(Mandatory)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory)][hashtable]$AllRLKeys,
        [Parameter(Mandatory)][hashtable]$AllAverittKeys, 
        [Parameter(Mandatory)][decimal]$DesiredASPValue, # Parameter for desired ASP
        [Parameter(Mandatory)][string]$CsvFilePath,
        [Parameter(Mandatory=$false)][bool]$ApplyMargins = $false,
        [Parameter(Mandatory=$false)][bool]$ASPFromHistory = $false
    )
    Write-Host "`nRunning Cross-Carrier ASP Analysis (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASPValue -le 0) { Write-Error "Desired ASP must be positive."; return $null }

    $aspDisplayLabel = if ($ASPFromHistory) { "Avg Booked Price (History)" } else { "Desired Avg Selling Price" }
    Write-Host "Customer: '$($SelectedCustomerProfile.CustomerName)', ${aspDisplayLabel}: $($DesiredASPValue.ToString('C2')), CSV: '$CsvFilePath'"

    $analysisTitle = if ($ASPFromHistory) { "Margin Analysis Based on History ASP for $($SelectedCustomerProfile.CustomerName)" } else { "ASP Cross-Carrier Margin Analysis for $($SelectedCustomerProfile.CustomerName)" }
    $reportTypeSuffix = if ($ASPFromHistory) { "MarginByHistory" } else { "MarginForASP" }

    # --- Data Loading (Needs all columns for all carriers) ---
    Write-Host "`nLoading and normalizing data for all carriers from '$CsvFilePath'..." -ForegroundColor Gray
    if (-not (Get-Command Load-And-Normalize-CentralData -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Load-And-Normalize-SAIAData -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Load-And-Normalize-RLData -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Load-And-Normalize-AverittData -ErrorAction SilentlyContinue) ) { Write-Error "Required data normalization function(s) not found."; return $null }

    $centralShipmentData = @(); $saiaShipmentData = @(); $rlShipmentData = @(); $averittShipmentData = @()

    $centralShipmentData = Load-And-Normalize-CentralData -CsvPath $CsvFilePath
    $saiaShipmentData = Load-And-Normalize-SAIAData -CsvPath $CsvFilePath
    $rlShipmentData = Load-And-Normalize-RLData -CsvPath $CsvFilePath
    $averittShipmentData = Load-And-Normalize-AverittData -CsvPath $CsvFilePath 

    if ($null -eq $centralShipmentData -or $null -eq $saiaShipmentData -or $null -eq $rlShipmentData -or $null -eq $averittShipmentData ) {
         Write-Error "Failed to load/normalize data for one or more carriers (function returned null)."
         if ($null -eq $centralShipmentData) { Write-Warning "Central data loading returned null." }
         if ($null -eq $saiaShipmentData) { Write-Warning "SAIA data loading returned null." }
         if ($null -eq $rlShipmentData) { Write-Warning "R+L data loading returned null." }
         if ($null -eq $averittShipmentData) { Write-Warning "Averitt data loading returned null."} 
         return $null
    }

    $totalRows = 0
    if($rlShipmentData.Count -gt $totalRows) { $totalRows = $rlShipmentData.Count }
    if($saiaShipmentData.Count -gt $totalRows) { $totalRows = $saiaShipmentData.Count }
    if($centralShipmentData.Count -gt $totalRows) { $totalRows = $centralShipmentData.Count }
    if($averittShipmentData.Count -gt $totalRows) { $totalRows = $averittShipmentData.Count }


    if ($totalRows -eq 0) { Write-Warning "CSV resulted in 0 processable rows after normalization for all carriers."; return $null }
    Write-Host " -> Data loaded and normalized (target ${totalRows} rows, individual counts may vary)." -ForegroundColor Green

    # --- Get Permitted Keys for the SELECTED CUSTOMER ---
    if (-not (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue)) { Write-Error "Get-PermittedKeys function not found."; return $null }
    $permittedCentral = Get-PermittedKeys -AllKeys $AllCentralKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedCentralKeys
    $permittedSAIA = Get-PermittedKeys -AllKeys $AllSAIAKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedSAIAKeys
    $permittedRL = Get-PermittedKeys -AllKeys $AllRLKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedRLKeys
    $permittedAveritt = Get-PermittedKeys -AllKeys $AllAverittKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedAverittKeys 

    if ($permittedCentral.Count -eq 0 -and $permittedSAIA.Count -eq 0 -and $permittedRL.Count -eq 0 -and $permittedAveritt.Count -eq 0) { Write-Warning "No permitted keys/accounts found for Customer '$($SelectedCustomerProfile.CustomerName)'. Analysis cannot proceed."; return $null }

    # --- Initialize Results Storage ---
    $aspResults = [System.Collections.Generic.List[object]]::new()
    $overallSkippedCount = 0
    if (-not (Get-Command Invoke-CentralTransportApi -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Invoke-SAIAApi -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Invoke-RLApi -ErrorAction SilentlyContinue) -or `
        -not (Get-Command Invoke-AverittApi -ErrorAction SilentlyContinue)) { Write-Error "Required API invocation function(s) not found."; return $null }

    # --- Process Each Carrier ---
    $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue';
    try {
        # Process Central Keys
        if ($permittedCentral.Count -gt 0 -and $centralShipmentData.Count -gt 0) {
            Write-Host "`n--- Processing Central Transport Accounts for $($SelectedCustomerProfile.CustomerName) ---" -ForegroundColor Yellow
            foreach ($keyName in ($permittedCentral.Keys | Sort-Object)) {
                $keyData = $permittedCentral[$keyName]; Write-Host "Processing Account: ${keyName} (using $($centralShipmentData.Count) Central-normalized rows)..." -ForegroundColor Cyan; $totalCostValue = 0.0; $processedCount = 0; $apiSkippedCount = 0
                $carrierTotalRows = $centralShipmentData.Count
                for ($i = 0; $i -lt $carrierTotalRows; $i++) {
                    $shipment = $centralShipmentData[$i];
                    $apiParams = @{ KeyData = $keyData; ShipmentData = $shipment }
                    $costValue = Invoke-CentralTransportApi @apiParams
                    if ($costValue -eq $null) { $apiSkippedCount++; continue }
                    $processedCount++; $totalCostValue += $costValue
                }
                $avgCost = 0.0; $reqMargin = $null;
                if ($processedCount -gt 0) {
                    $avgCost = $totalCostValue / $processedCount
                    if ($DesiredASPValue -ne 0) {
                        try { $reqMargin = (($DesiredASPValue - $avgCost) / $DesiredASPValue) * 100.0 }
                        catch { Write-Warning "Error calculating margin for ${keyName}: $($_.Exception.Message)"}
                    }
                }
                $aspResults.Add([PSCustomObject]@{ Account = $keyName; Carrier = 'Central'; AvgCost = $avgCost; RequiredMargin = $reqMargin; Processed = $processedCount; Skipped = $apiSkippedCount });
                if ($processedCount -gt 0) { Write-Host " -> Avg Cost: $($avgCost.ToString('C2')). Required Margin: $(if($reqMargin -ne $null){'{0:N2}%' -f $reqMargin}else{'N/A'}) (${processedCount} processed, ${apiSkippedCount} API skips)" -ForegroundColor Green } else { Write-Warning " -> No shipments processed successfully via API for Central account '${keyName}'. (${apiSkippedCount} API skips)"; }
                $overallSkippedCount += $apiSkippedCount
            }
        } elseif ($permittedCentral.Count -gt 0 -and $centralShipmentData.Count -eq 0) {
            Write-Warning "`nNo processable Central Transport data rows from CSV for $($SelectedCustomerProfile.CustomerName), though permitted keys exist."
        } else { Write-Host "`nNo permitted Central Transport accounts for $($SelectedCustomerProfile.CustomerName)." -ForegroundColor Gray }

        # Process SAIA Keys
        if ($permittedSAIA.Count -gt 0 -and $saiaShipmentData.Count -gt 0) {
            Write-Host "`n--- Processing SAIA Accounts for $($SelectedCustomerProfile.CustomerName) ---" -ForegroundColor Yellow
            foreach ($keyName in ($permittedSAIA.Keys | Sort-Object)) {
                $keyData = $permittedSAIA[$keyName]; Write-Host "Processing Account: ${keyName} (using $($saiaShipmentData.Count) SAIA-normalized rows)..." -ForegroundColor Cyan; $totalCostValue = 0.0; $processedCount = 0; $apiSkippedCount = 0
                $carrierTotalRows = $saiaShipmentData.Count
                for ($i = 0; $i -lt $carrierTotalRows; $i++) {
                    $shipment = $saiaShipmentData[$i];
                    $costValue = Invoke-SAIAApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -OriginCity $shipment.OriginCity -OriginState $shipment.OriginState -DestinationCity $shipment.DestinationCity -DestinationState $shipment.DestinationState -Details $shipment.details -KeyData $keyData
                    if ($costValue -eq $null) { $apiSkippedCount++; continue }
                    $processedCount++; $totalCostValue += $costValue
                }
                $avgCost = 0.0; $reqMargin = $null;
                if ($processedCount -gt 0) {
                    $avgCost = $totalCostValue / $processedCount
                    if ($DesiredASPValue -ne 0) {
                        try { $reqMargin = (($DesiredASPValue - $avgCost) / $DesiredASPValue) * 100.0 }
                        catch {Write-Warning "Error calculating margin for ${keyName}: $($_.Exception.Message)"}
                    }
                }
                $aspResults.Add([PSCustomObject]@{ Account = $keyName; Carrier = 'SAIA'; AvgCost = $avgCost; RequiredMargin = $reqMargin; Processed = $processedCount; Skipped = $apiSkippedCount });
                if ($processedCount -gt 0) { Write-Host " -> Avg Cost: $($avgCost.ToString('C2')). Required Margin: $(if($reqMargin -ne $null){'{0:N2}%' -f $reqMargin}else{'N/A'}) (${processedCount} processed, ${apiSkippedCount} API skips)" -ForegroundColor Green } else { Write-Warning " -> No shipments processed successfully via API for SAIA account '${keyName}'. (${apiSkippedCount} API skips)"; }
                $overallSkippedCount += $apiSkippedCount
            }
        } elseif ($permittedSAIA.Count -gt 0 -and $saiaShipmentData.Count -eq 0) {
            Write-Warning "`nNo processable SAIA data rows from CSV for $($SelectedCustomerProfile.CustomerName), though permitted keys exist."
        } else { Write-Host "`nNo permitted SAIA accounts for $($SelectedCustomerProfile.CustomerName)." -ForegroundColor Gray }

        # Process R+L Keys
        if ($permittedRL.Count -gt 0 -and $rlShipmentData.Count -gt 0) {
            Write-Host "`n--- Processing R+L Accounts for $($SelectedCustomerProfile.CustomerName) ---" -ForegroundColor Yellow
            foreach ($keyName in ($permittedRL.Keys | Sort-Object)) {
                $keyData = $permittedRL[$keyName]; Write-Host "Processing Account: ${keyName} (using $($rlShipmentData.Count) R+L-normalized rows)..." -ForegroundColor Cyan; $totalCostValue = 0.0; $processedCount = 0; $apiSkippedCount = 0
                $carrierTotalRows = $rlShipmentData.Count
                for ($i = 0; $i -lt $carrierTotalRows; $i++) {
                    $shipment = $rlShipmentData[$i];
                    $costValue = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Commodities $shipment.Commodities -KeyData $keyData -ShipmentDetails $shipment
                    if ($costValue -eq $null) { $apiSkippedCount++; continue }
                    $processedCount++; $totalCostValue += $costValue
                }
                $avgCost = 0.0; $reqMargin = $null;
                if ($processedCount -gt 0) {
                    $avgCost = $totalCostValue / $processedCount
                    if ($DesiredASPValue -ne 0) {
                        try { $reqMargin = (($DesiredASPValue - $avgCost) / $DesiredASPValue) * 100.0 }
                        catch {Write-Warning "Error calculating margin for ${keyName}: $($_.Exception.Message)"}
                    }
                }
                $aspResults.Add([PSCustomObject]@{ Account = $keyName; Carrier = 'R+L'; AvgCost = $avgCost; RequiredMargin = $reqMargin; Processed = $processedCount; Skipped = $apiSkippedCount });
                if ($processedCount -gt 0) { Write-Host " -> Avg Cost: $($avgCost.ToString('C2')). Required Margin: $(if($reqMargin -ne $null){'{0:N2}%' -f $reqMargin}else{'N/A'}) (${processedCount} processed, ${apiSkippedCount} API skips)" -ForegroundColor Green } else { Write-Warning " -> No shipments processed successfully via API for R+L account '${keyName}'. (${apiSkippedCount} API skips)"; }
                $overallSkippedCount += $apiSkippedCount
            }
        } elseif ($permittedRL.Count -gt 0 -and $rlShipmentData.Count -eq 0) {
            Write-Warning "`nNo processable R+L data rows from CSV for $($SelectedCustomerProfile.CustomerName), though permitted keys exist."
        } else { Write-Host "`nNo permitted R+L accounts for $($SelectedCustomerProfile.CustomerName)." -ForegroundColor Gray }

        # Process Averitt Keys
        if ($permittedAveritt.Count -gt 0 -and $averittShipmentData.Count -gt 0) {
            Write-Host "`n--- Processing Averitt Accounts for $($SelectedCustomerProfile.CustomerName) ---" -ForegroundColor Yellow
            foreach ($keyName in ($permittedAveritt.Keys | Sort-Object)) {
                $keyData = $permittedAveritt[$keyName]; Write-Host "Processing Account: ${keyName} (using $($averittShipmentData.Count) Averitt-normalized rows)..." -ForegroundColor Cyan; $totalCostValue = 0.0; $processedCount = 0; $apiSkippedCount = 0
                $carrierTotalRows = $averittShipmentData.Count
                for ($i = 0; $i -lt $carrierTotalRows; $i++) {
                    $shipment = $averittShipmentData[$i]; 
                    $costValue = Invoke-AverittApi -KeyData $keyData -ShipmentData $shipment
                    if ($costValue -eq $null) { $apiSkippedCount++; continue }
                    $processedCount++; $totalCostValue += $costValue
                }
                $avgCost = 0.0; $reqMargin = $null;
                if ($processedCount -gt 0) {
                    $avgCost = $totalCostValue / $processedCount
                    if ($DesiredASPValue -ne 0) {
                        try { $reqMargin = (($DesiredASPValue - $avgCost) / $DesiredASPValue) * 100.0 }
                        catch {Write-Warning "Error calculating margin for ${keyName}: $($_.Exception.Message)"}
                    }
                }
                $aspResults.Add([PSCustomObject]@{ Account = $keyName; Carrier = 'Averitt'; AvgCost = $avgCost; RequiredMargin = $reqMargin; Processed = $processedCount; Skipped = $apiSkippedCount });
                if ($processedCount -gt 0) { Write-Host " -> Avg Cost: $($avgCost.ToString('C2')). Required Margin: $(if($reqMargin -ne $null){'{0:N2}%' -f $reqMargin}else{'N/A'}) (${processedCount} processed, ${apiSkippedCount} API skips)" -ForegroundColor Green } else { Write-Warning " -> No shipments processed successfully via API for Averitt account '${keyName}'. (${apiSkippedCount} API skips)"; }
                $overallSkippedCount += $apiSkippedCount
            }
        } elseif ($permittedAveritt.Count -gt 0 -and $averittShipmentData.Count -eq 0) {
            Write-Warning "`nNo processable Averitt data rows from CSV for $($SelectedCustomerProfile.CustomerName), though permitted keys exist."
        } else { Write-Host "`nNo permitted Averitt accounts for $($SelectedCustomerProfile.CustomerName)." -ForegroundColor Gray }


    } finally {
        $VerbosePreference = $CurrentVerbosePreference
    }

    # --- Prepare Report Content ---
    $reportContent = [System.Collections.Generic.List[string]]::new()
    $reportContent.Add($analysisTitle); $reportContent.Add("=" * $analysisTitle.Length)
    $reportContent.Add("Date: $(Get-Date), Broker: $($BrokerProfile.Username), Input CSV: $(Split-Path $CsvFilePath -Leaf)")
    $reportContent.Add((" {0,-25}: {1:C2}" -f $aspDisplayLabel, $DesiredASPValue))
    $reportContent.Add("Total Unique CSV Rows Available (approx based on largest carrier set): ${totalRows}")
    $reportContent.Add("Total API Call Skips (sum across all accounts): ${overallSkippedCount}")
    $reportContent.Add("")
    if ($aspResults.Count -gt 0) {
        $formattedTableOutput = ($aspResults | Sort-Object -Property Carrier, Account | Format-Table Carrier, Account, @{N='Avg Cost';E={if ($_.AvgCost -ne $null) {$_.AvgCost.ToString("C2")} else {'N/A'}}}, @{N='Required Margin %';E={if($_.RequiredMargin -ne $null){$_.RequiredMargin.ToString("N2") + "%"}else{"N/A"}}}, Processed, @{N='API Skips';E={$_.Skipped}} -AutoSize | Out-String)
        if ($null -ne $formattedTableOutput) {
            $linesToAdd = $formattedTableOutput.ToString().TrimEnd("`r", "`n").Split([Environment]::NewLine)
            foreach($lineToAdd in $linesToAdd){
                $reportContent.Add($lineToAdd)
            }
        }
    } else { $reportContent.Add("No results generated (no permitted keys or no data processed).") }
    $reportContent.Add("-------------------------------------------------")

    # --- Apply Margins if Requested ---
    if ($ApplyMargins -and $aspResults.Count -gt 0) {
        Write-Host "`nApplying calculated margins..." -ForegroundColor Yellow
        $reportContent.Add(""); $reportContent.Add("--- Margin Update Attempt Summary ---")
        $updateSuccessCount = 0; $updateFailCount = 0; $updateSkippedCount = 0; $updateInvalidMarginCount = 0;
        foreach ($result in $aspResults) {
            if ($result.Processed -eq 0) {
                $reportContent.Add("Skipping margin update for '$($result.Account)': No shipments were processed for cost calculation."); $updateSkippedCount++; continue
            }
            $tariffName = $result.Account; $keysFolderPath = $null; $allKeysToUse = $null

            if ($result.RequiredMargin -eq $null -or -not ($result.RequiredMargin -is [double])) { $reportContent.Add("Skipping '${tariffName}': Invalid required margin calculation."); $updateSkippedCount++; continue }
            $newMargin = [math]::Round($result.RequiredMargin, 1)
            if ($newMargin -lt 0 -or $newMargin -ge 100) { $reportContent.Add("Skipping '${tariffName}': Calculated margin ${newMargin}% outside valid range (0-99.9)."); $updateInvalidMarginCount++; continue }

            switch ($result.Carrier) {
                'Central' { $keysFolderPath = $script:centralKeysFolderPath; $allKeysToUse = $AllCentralKeys }
                'SAIA'    { $keysFolderPath = $script:saiaKeysFolderPath; $allKeysToUse = $AllSAIAKeys }
                'R+L'     { $keysFolderPath = $script:rlKeysFolderPath; $allKeysToUse = $AllRLKeys }
                'Averitt' { $keysFolderPath = $script:averittKeysFolderPath; $allKeysToUse = $AllAverittKeys } 
                default   { Write-Warning "Unknown carrier '$($result.Carrier)' for margin update."; $updateFailCount++; continue }
            }
            if ([string]::IsNullOrWhiteSpace($keysFolderPath)) { Write-Error "Keys folder path not determined for '$($result.Carrier)'."; $updateFailCount++; continue}
            if ($null -eq $allKeysToUse) { Write-Error "Source AllKeys hashtable for carrier '$($result.Carrier)' is null."; $updateFailCount++; continue}

            if (-not (Get-Command Update-TariffMargin -ErrorAction SilentlyContinue)) { Write-Error "Update-TariffMargin function not found!"; $updateFailCount++; continue }

            Write-Host "Attempting to update margin for '${tariffName}' (Carrier: $($result.Carrier)) to ${newMargin} % using folder '${keysFolderPath}'."
            $updateAttemptSuccess = Update-TariffMargin -TariffName $tariffName -AllKeysHashtable $allKeysToUse -KeysFolderPath $keysFolderPath -NewMarginPercent $newMargin
            if ($updateAttemptSuccess) { $reportContent.Add("Success: '${tariffName}' updated to ${newMargin}%."); $updateSuccessCount++ }
            else { $reportContent.Add("FAILED update for '${tariffName}'. Check console/logs."); $updateFailCount++ }
        }
         $reportContent.Add("----------------------------------"); $reportContent.Add("Update Counts - Success: ${updateSuccessCount}, Failed: ${updateFailCount}, Skipped (No Data/Calc): ${updateSkippedCount}, Skipped (Invalid Range): ${updateInvalidMarginCount}")
    } elseif ($ApplyMargins) {
         $reportContent.Add(""); $reportContent.Add("NOTE: Apply Margins requested, but no results were generated to apply from.")
    }

    # --- Save Report ---
    if (-not (Get-Command Get-ReportPath -ErrorAction SilentlyContinue)) { Write-Error "Get-ReportPath function not found."; return $null }

    $customerNameSafe = $SelectedCustomerProfile.CustomerName -replace '[^a-zA-Z0-9.-]', '_'
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $BrokerProfile.Username -Carrier 'CrossCarrier' -ReportType $reportTypeSuffix -FilePrefix $customerNameSafe

    if ($reportFilePath) {
        try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -Force -ErrorAction Stop; Write-Host "`nReport saved: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
        catch { Write-Error "Failed to save report: $($_.Exception.Message)"; return $null }
    } else { Write-Error "Failed to generate report path."; return $null }
}


function Run-MarginsByHistoryAnalysisGUI {
    param(
        [Parameter(Mandatory)][hashtable]$BrokerProfile,
        [Parameter(Mandatory)][hashtable]$SelectedCustomerProfile,
        [Parameter(Mandatory)][string]$ReportsBaseFolder,
        [Parameter(Mandatory)][string]$UserReportsFolder,
        [Parameter(Mandatory)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory)][hashtable]$AllRLKeys,
        [Parameter(Mandatory)][hashtable]$AllAverittKeys, 
        [Parameter(Mandatory)][string]$CsvFilePath,
        [Parameter(Mandatory=$false)][bool]$ApplyMargins = $false
    )
    Write-Host "`nRunning Margins by History Analysis for Customer '$($SelectedCustomerProfile.CustomerName)' (GUI Mode)..." -ForegroundColor Cyan

    $totalBookedPrice = 0.0; $validBookedPriceCount = 0; $averageBookedPrice = $null
    try {
        if(-not (Test-Path $CsvFilePath -PathType Leaf)) { Write-Error "History CSV file not found: '$CsvFilePath'"; return $null }
        $rawData = Import-Csv -Path $CsvFilePath -ErrorAction Stop
        if ($rawData.Count -eq 0) { Write-Warning "History CSV '$CsvFilePath' is empty."; return $null }

        $header = $rawData[0].PSObject.Properties.Name
        $bookedPriceColumnName = $null
        $potentialPriceColumns = @('Booked Price', 'BookedPrice', 'Price', 'FinalQuotedPrice', 'Average Selling Price', 'CustomerRate') 
        foreach($colName in $potentialPriceColumns){
            if($header -contains $colName){
                $bookedPriceColumnName = $colName
                break
            }
        }
        if($null -eq $bookedPriceColumnName){ Write-Error "History CSV '$CsvFilePath' missing a recognizable booked price column (e.g., 'Booked Price', 'Price', 'CustomerRate')."; return $null }

        Write-Host "Calculating Average Booked Price from '$bookedPriceColumnName' column..." -ForegroundColor Gray
        foreach ($row in $rawData) {
            if ($row.$bookedPriceColumnName -ne $null -and $row.$bookedPriceColumnName -ne "") {
                try {
                    $priceString = $row.$bookedPriceColumnName -replace '[$,]'
                    $bookedPrice = [decimal]$priceString
                    if ($bookedPrice -gt 0) {
                        $totalBookedPrice += $bookedPrice
                        $validBookedPriceCount++
                    }
                } catch {
                    Write-Verbose "Skipping row due to conversion error for '${bookedPriceColumnName}': Value '$($row.$bookedPriceColumnName)'"
                }
            }
        }
        if ($validBookedPriceCount -eq 0) { Write-Warning "No valid positive '${bookedPriceColumnName}' values found in '${CsvFilePath}' after attempting conversion."; return $null }
        $averageBookedPrice = $totalBookedPrice / $validBookedPriceCount
        Write-Host " -> Average Booked Price: $($averageBookedPrice.ToString('C2')) (${validBookedPriceCount} valid rows)" -ForegroundColor Green
    } catch { Write-Error "Error processing History CSV '${CsvFilePath}': $($_.Exception.Message)"; return $null }

    if (-not (Get-Command Run-CrossCarrierASPAnalysisGUI -ErrorAction SilentlyContinue)) { Write-Error "Run-CrossCarrierASPAnalysisGUI function not found."; return $null }

    $reportPath = Run-CrossCarrierASPAnalysisGUI -BrokerProfile $BrokerProfile `
                                                 -SelectedCustomerProfile $SelectedCustomerProfile `
                                                 -ReportsBaseFolder $ReportsBaseFolder `
                                                 -UserReportsFolder $UserReportsFolder `
                                                 -AllCentralKeys $AllCentralKeys `
                                                 -AllSAIAKeys $AllSAIAKeys `
                                                 -AllRLKeys $AllRLKeys `
                                                 -AllAverittKeys $AllAverittKeys `
                                                 -DesiredASPValue $averageBookedPrice `
                                                 -CsvFilePath $CsvFilePath `
                                                 -ApplyMargins $ApplyMargins `
                                                 -ASPFromHistory $true

    return $reportPath
}

Write-Verbose "TMS Reports Functions loaded (GUI Refactored)."